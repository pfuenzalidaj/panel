import argparse
import importlib
import json
import mimetypes
import tempfile
import threading
import webbrowser
from email import policy
from email.parser import BytesParser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

import build_excel_pilot


BASE_DIR = Path(__file__).resolve().parent
UI_FILE = BASE_DIR / "Actualizador_Dashboard.html"
DASHBOARD_FILE = BASE_DIR / "CEN_Dashboard_Ejecutivo_Jul2026_Piloto_Excel.html"
PROBLEMS_FILE = BASE_DIR / "20260707_Informe_modulo_Problemas.xlsx"
SHEET_RULE = "Detección automática por columnas"
MAX_UPLOAD_BYTES = 100 * 1024 * 1024
APP_VERSION = "2.1"


class DashboardHandler(BaseHTTPRequestHandler):
    server_version = "CENLocalUpdater/1.0"

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {format % args}")

    def send_json(self, payload, status=HTTPStatus.OK):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path, download=False):
        if not path.is_file():
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        body = path.read_bytes()
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header(
            "Content-Type",
            f"{content_type}; charset=utf-8" if content_type.startswith("text/") else content_type,
        )
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        if download:
            self.send_header("Content-Disposition", f'attachment; filename="{path.name}"')
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        route = urlparse(self.path).path
        if route in ("/", "/index.html"):
            self.send_file(UI_FILE)
        elif route == "/dashboard":
            self.send_file(DASHBOARD_FILE)
        elif route == "/download":
            self.send_file(DASHBOARD_FILE, download=True)
        elif route == "/api/status":
            self.send_json(
                {
                    "ready": UI_FILE.is_file() and DASHBOARD_FILE.is_file(),
                    "dashboard": DASHBOARD_FILE.name,
                    "sheetRule": SHEET_RULE,
                    "version": APP_VERSION,
                }
            )
        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def do_POST(self):
        route = urlparse(self.path).path
        if route == "/api/update":
            self.handle_update()
        elif route == "/api/shutdown":
            self.send_json({"ok": True})
            threading.Thread(target=self.server.shutdown, daemon=True).start()
        else:
            self.send_error(HTTPStatus.NOT_FOUND)

    def handle_update(self):
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError:
            self.send_json(
                {"ok": False, "error": "Tamaño de solicitud inválido."},
                HTTPStatus.BAD_REQUEST,
            )
            return

        if length <= 0 or length > MAX_UPLOAD_BYTES:
            self.send_json(
                {"ok": False, "error": "El archivo debe pesar menos de 100 MB."},
                HTTPStatus.REQUEST_ENTITY_TOO_LARGE,
            )
            return

        content_type = self.headers.get("Content-Type", "")
        if not content_type.lower().startswith("multipart/form-data"):
            self.send_json(
                {"ok": False, "error": "Solicitud de carga inválida."},
                HTTPStatus.BAD_REQUEST,
            )
            return

        raw = self.rfile.read(length)
        message = BytesParser(policy=policy.default).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode("ascii")
            + raw
        )
        parts = list(message.iter_parts())
        upload = next(
            (
                part
                for part in parts
                if part.get_content_disposition() == "form-data"
                and part.get_param("name", header="content-disposition") == "file"
            ),
            None,
        )
        problems_upload = next(
            (
                part
                for part in parts
                if part.get_content_disposition() == "form-data"
                and part.get_param("name", header="content-disposition")
                == "problems"
            ),
            None,
        )
        if upload is None:
            self.send_json(
                {"ok": False, "error": "No se recibió un archivo."},
                HTTPStatus.BAD_REQUEST,
            )
            return

        original_name = Path(upload.get_filename() or "insumo.xlsx").name
        if Path(original_name).suffix.lower() != ".xlsx":
            self.send_json(
                {"ok": False, "error": "Selecciona un archivo con extensión .xlsx."},
                HTTPStatus.BAD_REQUEST,
            )
            return

        file_bytes = upload.get_payload(decode=True) or b""
        if not file_bytes.startswith(b"PK"):
            self.send_json(
                {"ok": False, "error": "El archivo no parece ser un libro XLSX válido."},
                HTTPStatus.BAD_REQUEST,
            )
            return

        temporary_path = None
        problems_temporary_path = None
        try:
            with tempfile.NamedTemporaryFile(
                prefix="cen_upload_", suffix=".xlsx", dir=BASE_DIR, delete=False
            ) as temporary:
                temporary.write(file_bytes)
                temporary_path = Path(temporary.name)

            problems_path = PROBLEMS_FILE if PROBLEMS_FILE.is_file() else None
            problems_original_name = PROBLEMS_FILE.name if problems_path else None
            if problems_upload is not None and problems_upload.get_filename():
                problems_original_name = Path(
                    problems_upload.get_filename()
                ).name
                if Path(problems_original_name).suffix.lower() != ".xlsx":
                    raise ValueError(
                        "El seguimiento de problemas debe ser un archivo .xlsx."
                    )
                problems_bytes = problems_upload.get_payload(decode=True) or b""
                if not problems_bytes.startswith(b"PK"):
                    raise ValueError(
                        "El seguimiento de problemas no parece un XLSX válido."
                    )
                with tempfile.NamedTemporaryFile(
                    prefix="cen_problems_",
                    suffix=".xlsx",
                    dir=BASE_DIR,
                    delete=False,
                ) as problems_temporary:
                    problems_temporary.write(problems_bytes)
                    problems_temporary_path = Path(problems_temporary.name)
                problems_path = problems_temporary_path

            generator = importlib.reload(build_excel_pilot)
            data = generator.update_dashboard(
                xlsx_path=temporary_path,
                html_path=DASHBOARD_FILE,
                source_name=original_name,
                problems_path=problems_path,
                problems_source_name=problems_original_name,
            )
            self.send_json(
                {
                    "ok": True,
                    "source": original_name,
                    "generated": data.get("fuente_excel", {}).get("generado"),
                    "tickets": data.get("kpi", {}).get("tickets26"),
                    "incidents": data.get("kpi", {}).get("inc26"),
                    "sheet": data.get("fuente_excel", {}).get("hoja"),
                    "problemsSource": data.get("fuente_problemas", {}).get(
                        "archivo"
                    ),
                    "problems": data.get("kpi", {}).get("prob_abiertos"),
                    "overdue": data.get("prob_kpi", {}).get(
                        "commit_overdue"
                    ),
                    "withoutCommitment": data.get("prob_kpi", {}).get(
                        "commit_missing"
                    ),
                    "dashboard": DASHBOARD_FILE.name,
                }
            )
        except Exception as error:
            text = str(error)
            self.send_json(
                {"ok": False, "error": text or "No fue posible procesar el Excel."},
                HTTPStatus.BAD_REQUEST,
            )
        finally:
            if temporary_path and temporary_path.exists():
                temporary_path.unlink()
            if (
                problems_temporary_path
                and problems_temporary_path.exists()
            ):
                problems_temporary_path.unlink()


def main():
    parser = argparse.ArgumentParser(
        description="Interfaz local para actualizar el dashboard CEN."
    )
    parser.add_argument("--port", type=int, default=8765)
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args()

    if not UI_FILE.is_file() or not DASHBOARD_FILE.is_file():
        raise SystemExit(
            "Faltan la interfaz HTML o el dashboard en la carpeta de la aplicación."
        )

    try:
        server = ThreadingHTTPServer(("127.0.0.1", args.port), DashboardHandler)
    except OSError:
        server = ThreadingHTTPServer(("127.0.0.1", 0), DashboardHandler)

    port = server.server_address[1]
    url = f"http://127.0.0.1:{port}/"
    print(f"Actualizador disponible en {url}")
    print("Cierra esta ventana o usa el botón de la interfaz para detenerlo.")
    if not args.no_browser:
        threading.Timer(0.7, lambda: webbrowser.open(url)).start()
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
