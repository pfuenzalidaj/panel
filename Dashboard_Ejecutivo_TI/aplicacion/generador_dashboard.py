import argparse
import calendar
import json
import math
import re
from collections import OrderedDict
from datetime import datetime
from html import escape
from pathlib import Path

import pandas as pd


APP_DIR = Path(__file__).resolve().parent
PACKAGE_DIR = APP_DIR.parent
SOURCES_DIR = PACKAGE_DIR / "fuentes"
DASHBOARD_DIR = PACKAGE_DIR / "dashboard"
HTML_IN = DASHBOARD_DIR / "Dashboard_Ejecutivo_Gestion_TI.html"
XLSX_IN = SOURCES_DIR / "Tickets_Incidentes.xlsx"
PROBLEMS_IN = SOURCES_DIR / "Informe_Problemas.xlsx"
CHANGES_IN = SOURCES_DIR / "Reporte_RFC-CAB_2026_237_278_202607101029.xlsx"
HTML_OUT = HTML_IN

PRIOS = ["P1- Critico", "P2 - Alto", "P3 - Normal", "P4 - Baja"]
REPORT_YEAR = 2026
MONTH_HOURS = {
    month: calendar.monthrange(REPORT_YEAR, month)[1] * 24
    for month in range(1, 13)
}
MONTH_NAMES_SHORT = {
    1: "Ene",
    2: "Feb",
    3: "Mar",
    4: "Abr",
    5: "May",
    6: "Jun",
    7: "Jul",
    8: "Ago",
    9: "Sep",
    10: "Oct",
    11: "Nov",
    12: "Dic",
}
REQUIRED_COLUMNS = {
    "Tipo ticket",
    "Prioridad",
    "Cumple resolución",
    "Categoria",
    "SubCategoria",
    "Equipo",
    "Fec. creación",
    "ID Entrada",
    "Título",
    "Tiempo resolución",
}
PROBLEM_REQUIRED_COLUMNS = {
    "Jefe Dpto",
    "ID Problema",
    "Estado",
    "Fec. creación",
    "Fecha de Compromisos",
    "SubCategoria",
    "Asignado a",
    "Breve descripción",
    "Descripción (historial)",
    "Causa raíz",
    "Sub Gerencia",
    "Workaround",
    "Tipo de Compromiso",
    "Seguimiento",
    "Tickets Asociados",
}
CHANGE_REQUIRED_COLUMNS = {
    "ID Entrada",
    "Título",
    "SubCategoria",
    "Detalle / Módulo",
    "Mail",
    "Fecha/hora inicio",
    "Fecha/hora termino",
    "Clasificación Cambio",
    "Cambio",
    "Estado",
    "Asignado a",
    "Categoria",
}


def detect_input_sheet(workbook):
    compatible = []
    for index, sheet_name in enumerate(workbook.sheet_names):
        try:
            columns = set(pd.read_excel(workbook, sheet_name=sheet_name, nrows=0).columns)
        except Exception:
            continue
        if REQUIRED_COLUMNS.issubset(columns):
            normalized_name = sheet_name.casefold()
            score = int("tickets_full" in normalized_name) + int(
                "confidencial" in normalized_name
            )
            compatible.append((score, -index, sheet_name))

    if not compatible:
        available = ", ".join(workbook.sheet_names)
        raise ValueError(
            "No se encontró una hoja con las columnas requeridas. "
            f"Hojas revisadas: {available or '(ninguna)'}."
        )
    return max(compatible)[2]


def detect_change_sheet(workbook):
    compatible = []
    for index, sheet_name in enumerate(workbook.sheet_names):
        try:
            columns = set(pd.read_excel(workbook, sheet_name=sheet_name, nrows=0).columns)
        except Exception:
            continue
        if CHANGE_REQUIRED_COLUMNS.issubset(columns):
            normalized_name = sheet_name.casefold()
            score = int("reporte_rfc-cab" in normalized_name) * 2 + int(
                "confiden" in normalized_name or "cambio" in normalized_name
            )
            compatible.append((score, -index, sheet_name))
    if not compatible:
        available = ", ".join(workbook.sheet_names)
        raise ValueError(
            "No se encontrÃ³ una hoja compatible con GestiÃ³n del Cambio. "
            f"Hojas revisadas: {available or '(ninguna)'}."
        )
    return max(compatible)[2]


def detect_problem_sheet(workbook):
    compatible = []
    for index, sheet_name in enumerate(workbook.sheet_names):
        try:
            columns = set(pd.read_excel(workbook, sheet_name=sheet_name, nrows=0).columns)
        except Exception:
            continue
        if PROBLEM_REQUIRED_COLUMNS.issubset(columns):
            normalized_name = sheet_name.casefold()
            score = int("informe_modulo_problemas" in normalized_name) * 2 + int(
                "consolidado" in normalized_name
            )
            compatible.append((score, -index, sheet_name))
    if not compatible:
        available = ", ".join(workbook.sheet_names)
        raise ValueError(
            "No se encontró una hoja compatible con el módulo de Problemas. "
            f"Hojas revisadas: {available or '(ninguna)'}."
        )
    return max(compatible)[2]


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def compact_text(v):
    text = re.sub(r"\s+", " ", clean(v)).strip()
    if text.casefold() in {"- sin valor -", "sin valor", "nan", "none"}:
        return ""
    return text


def safe_int(v, default=0):
    numeric = pd.to_numeric(v, errors="coerce")
    return default if pd.isna(numeric) else int(numeric)


def excel_datetime_text(v):
    if pd.isna(v):
        return ""
    dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return clean(v)
    return dt.strftime("%d-%m-%Y %H:%M")


def change_state_class(status):
    value = compact_text(status).casefold()
    if not value:
        return "medio"
    if "rollback" in value or "no exitosa" in value or "fallida" in value:
        return "no"
    if "cancel" in value:
        return "medio"
    if "implement" in value or "ingresada" in value:
        return "medio"
    if "exitosa" in value or value in {"cerrada", "cerrado"}:
        return "ok"
    return "medio"


def build_problem_data(
    base_d,
    problems_path,
    problems_source_name,
    incidents,
    as_of_date,
):
    problems_path = Path(problems_path)
    with pd.ExcelFile(problems_path) as workbook:
        problem_sheet = detect_problem_sheet(workbook)
        problems = pd.read_excel(workbook, sheet_name=problem_sheet)

    missing_columns = sorted(PROBLEM_REQUIRED_COLUMNS.difference(problems.columns))
    if missing_columns:
        raise ValueError(
            f'En la hoja "{problem_sheet}" faltan columnas requeridas: '
            + ", ".join(missing_columns)
        )

    state_counts = {
        str(key): int(value)
        for key, value in problems["Estado"]
        .fillna("Sin estado")
        .astype(str)
        .str.strip()
        .value_counts()
        .items()
    }
    closed_states = {"cerrado", "cancelado", "resuelto"}
    open_mask = ~problems["Estado"].fillna("").astype(str).str.strip().str.casefold().isin(
        closed_states
    )
    open_problems = problems[open_mask].copy()
    incident_counts = all_counts(incidents["SubCategoria"])
    risk_map = {"rojo": "Alto", "amarillo": "Medio", "verde": "Bajo"}
    commitment_order = {
        "Vencido": 0,
        "Vence hoy": 1,
        "Sin fecha": 2,
        "Próximo": 3,
        "En plazo": 4,
    }
    rows = []

    for _, row in open_problems.iterrows():
        subcategory = compact_text(row.get("SubCategoria")) or "Sin servicio"
        commitment = pd.to_datetime(
            row.get("Fecha de Compromisos"), dayfirst=True, errors="coerce"
        )
        if pd.isna(commitment):
            commitment_text = ""
            commitment_days = None
            commitment_status = "Sin fecha"
        else:
            commitment = commitment.normalize()
            commitment_text = commitment.strftime("%d-%m-%Y")
            commitment_days = int((commitment - as_of_date).days)
            if commitment_days < 0:
                commitment_status = "Vencido"
            elif commitment_days == 0:
                commitment_status = "Vence hoy"
            elif commitment_days <= 30:
                commitment_status = "Próximo"
            else:
                commitment_status = "En plazo"

        semaphore = compact_text(row.get("Semaforo"))
        priority = compact_text(row.get("Prioridad"))
        priority_key = priority.casefold()
        risk = risk_map.get(semaphore.casefold())
        if not risk:
            if "crít" in priority_key or "critic" in priority_key or "alta" in priority_key:
                risk = "Alto"
            elif "baja" in priority_key:
                risk = "Bajo"
            else:
                risk = "Medio"
        workaround_source = compact_text(row.get("Workaround"))
        has_workaround = workaround_source.casefold().startswith("si")
        age = pd.to_numeric(row.get("Ageing"), errors="coerce")
        row_data = {
            "id": clean(row.get("ID Problema")),
            "desc": compact_text(row.get("Breve descripción")),
            "history": compact_text(row.get("Descripción (historial)")),
            "cause": compact_text(row.get("Causa raíz")),
            "follow": compact_text(row.get("Seguimiento")),
            "sub": subcategory,
            "age": None if pd.isna(age) else int(age),
            "tk": safe_int(row.get("Tickets Asociados")),
            "incsub": int(incident_counts.get(subcategory, 0)),
            "wa": "Sí" if has_workaround else "No",
            "wa_detail": compact_text(row.get("En que Consiste el Workaround")),
            "subger": compact_text(row.get("Sub Gerencia"))
            or "Sin Subgerencia",
            "jefe": compact_text(row.get("Jefe Dpto")) or "Sin asignar",
            "owner": compact_text(row.get("Asignado a")) or "Sin asignar",
            "riesgo": risk,
            "semaforo": semaphore or "Sin semáforo",
            "state": compact_text(row.get("Estado")) or "Sin estado",
            "priority": priority,
            "commit": commitment_text,
            "commit_days": commitment_days,
            "commit_status": commitment_status,
            "commit_type": compact_text(row.get("Tipo de Compromiso"))
            or "Sin compromiso definido",
            "scope": compact_text(row.get("Alcance")),
            "occurrence": compact_text(row.get("Ocurrencia")),
            "created": excel_date_text(row.get("Fec. creación")),
        }
        rows.append(row_data)

    rows.sort(
        key=lambda item: (
            commitment_order[item["commit_status"]],
            item["commit_days"] if item["commit_days"] is not None else 99999,
            {"Alto": 0, "Medio": 1, "Bajo": 2}.get(item["riesgo"], 3),
            -(item["age"] or 0),
        )
    )
    subger_counts = OrderedDict(
        sorted(
            (
                (name, sum(1 for row in rows if row["subger"] == name))
                for name in {row["subger"] for row in rows}
            ),
            key=lambda item: (-item[1], item[0]),
        )
    )
    subger_top = next(iter(subger_counts.items()), ("Sin Subgerencia", 0))
    bubbles = [
        {
            "x": row["age"] or 0,
            "y": row["incsub"],
            "r": max(5, min(16, 5 + row["tk"] * 3)),
            "id": row["id"],
            "sub": row["sub"],
            "riesgo": row["riesgo"],
            "wa": row["wa"],
            "jefe": row["jefe"],
        }
        for row in rows
    ]
    urgent = [
        row
        for row in rows
        if row["commit_status"] in {"Vencido", "Vence hoy", "Sin fecha"}
        or (row["riesgo"] == "Alto" and row["wa"] == "No")
    ][:6]
    unique_incident_services = {row["sub"] for row in rows}

    base_d["prob_estado"] = state_counts
    base_d["prob_list"] = rows
    base_d["prob_bubbles"] = bubbles
    base_d["prob_crit"] = urgent
    base_d["prob_kpi"] = {
        "sin_wa": sum(1 for row in rows if row["wa"] == "No"),
        "age_crit": sum(1 for row in rows if (row["age"] or 0) > 120),
        "inc_atrib": sum(
            int(incident_counts.get(service, 0))
            for service in unique_incident_services
        ),
        "subger": subger_counts,
        "subger_top": list(subger_top),
        "commit_overdue": sum(
            1 for row in rows if row["commit_status"] == "Vencido"
        ),
        "commit_missing": sum(
            1 for row in rows if row["commit_status"] == "Sin fecha"
        ),
        "commit_next30": sum(
            1
            for row in rows
            if row["commit_status"] in {"Vence hoy", "Próximo"}
        ),
    }
    base_d["kpi"]["prob_abiertos"] = len(rows)
    base_d["kpi"]["prob_total"] = len(problems)
    base_d["fuente_problemas"] = {
        "archivo": problems_source_name or problems_path.name,
        "hoja": problem_sheet,
        "corte": as_of_date.strftime("%Y-%m-%d"),
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    return rows


def build_change_data(
    base_d,
    changes_path,
    changes_source_name,
    as_of_date,
):
    changes_path = Path(changes_path)
    with pd.ExcelFile(changes_path) as workbook:
        change_sheet = detect_change_sheet(workbook)
        changes = pd.read_excel(workbook, sheet_name=change_sheet)

    missing_columns = sorted(CHANGE_REQUIRED_COLUMNS.difference(changes.columns))
    if missing_columns:
        raise ValueError(
            f'En la hoja "{change_sheet}" faltan columnas requeridas: '
            + ", ".join(missing_columns)
        )

    start_dt = pd.to_datetime(changes["Fecha/hora inicio"], dayfirst=True, errors="coerce")
    end_dt = pd.to_datetime(changes["Fecha/hora termino"], dayfirst=True, errors="coerce")
    valid_months = start_dt.dt.month.dropna().astype(int)
    last_month = int(valid_months.max()) if not valid_months.empty else 1
    months = list(range(1, last_month + 1))

    state_counts = all_counts(changes["Estado"])
    class_counts = all_counts(changes["Clasificación Cambio"])
    type_counts = all_counts(changes["Cambio"])
    owner_counts = all_counts(changes["Asignado a"])
    category_counts = all_counts(changes["Categoria"])
    module_counts = all_counts(changes["Detalle / Módulo"])

    rows = []
    for idx, row in changes.iterrows():
        start = start_dt.iloc[idx]
        end = end_dt.iloc[idx]
        duration_h = None
        if pd.notna(start) and pd.notna(end):
            delta_h = (end - start).total_seconds() / 3600
            if delta_h >= 0:
                duration_h = round(delta_h, 1)
        duration_issue = pd.notna(end) and duration_h is None
        state = compact_text(row.get("Estado")) or "Sin estado"
        state_kind = change_state_class(state)
        month = int(start.month) if pd.notna(start) else 0
        rows.append(
            {
                "id": clean(row.get("ID Entrada")),
                "title": compact_text(row.get("Título")) or "Sin título",
                "sub": compact_text(row.get("SubCategoria")) or "Sin subcategoría",
                "detail": compact_text(row.get("Detalle / Módulo")) or "Sin módulo",
                "mail": compact_text(row.get("Mail")),
                "start": excel_datetime_text(start),
                "end": excel_datetime_text(end),
                "start_iso": None if pd.isna(start) else pd.to_datetime(start).strftime("%Y-%m-%d %H:%M"),
                "end_iso": None if pd.isna(end) else pd.to_datetime(end).strftime("%Y-%m-%d %H:%M"),
                "classif": compact_text(row.get("Clasificación Cambio")) or "Sin clasificación",
                "change": compact_text(row.get("Cambio")) or "Sin tipo",
                "state": state,
                "state_kind": state_kind,
                "owner": compact_text(row.get("Asignado a")) or "Sin asignar",
                "category": compact_text(row.get("Categoria")) or "Sin categoría",
                "month": month,
                "month_label": MONTH_NAMES_SHORT.get(month, "Sin mes"),
                "duration_h": duration_h,
                "duration_d": None if duration_h is None else round(duration_h / 24, 1),
                "duration_issue": duration_issue,
                "duration_txt": (
                    "Sin término"
                    if pd.isna(end)
                    else "Fechas inconsistentes"
                    if duration_issue
                    else f"{duration_h:.1f} h"
                ),
            }
        )

    rows.sort(
        key=lambda item: (
            0 if item["state_kind"] == "no" else 1 if item["state_kind"] == "medio" else 2,
            item["duration_h"] if item["duration_h"] is not None else 999999,
            -(item["month"] or 0),
            item["id"],
        )
    )

    month_counts = [int((start_dt.dt.month == month).sum()) for month in months]
    month_success = [
        sum(1 for row in rows if row["month"] == month and row["state_kind"] == "ok")
        for month in months
    ]
    month_failed = [
        sum(1 for row in rows if row["month"] == month and row["state_kind"] == "no")
        for month in months
    ]

    success_total = sum(1 for row in rows if row["state_kind"] == "ok")
    failed_total = sum(1 for row in rows if row["state_kind"] == "no")
    cancel_total = sum(1 for row in rows if "cancel" in row["state"].casefold())
    pending_total = sum(
        1
        for row in rows
        if "ingresada" in row["state"].casefold()
        or "implement" in row["state"].casefold()
    )
    inconsistent_total = sum(1 for row in rows if row["duration_issue"])
    durations = [row["duration_h"] for row in rows if row["duration_h"] is not None]
    avg_duration = round(sum(durations) / len(durations), 1) if durations else None
    max_duration = max(durations) if durations else None
    month_peak_idx = month_counts.index(max(month_counts)) + 1 if month_counts else 1
    month_peak = MONTH_NAMES_SHORT.get(month_peak_idx, "Sin mes")

    base_d["change_list"] = rows
    base_d["change_state"] = state_counts
    base_d["change_class"] = class_counts
    base_d["change_type"] = type_counts
    base_d["change_owner"] = owner_counts
    base_d["change_category"] = category_counts
    base_d["change_module"] = module_counts
    base_d["change_month"] = month_counts
    base_d["change_month_success"] = month_success
    base_d["change_month_failed"] = month_failed
    base_d["change_kpi"] = {
        "total": int(len(rows)),
        "success": success_total,
        "failed": failed_total,
        "cancelled": cancel_total,
        "pending": pending_total,
        "inconsistent": inconsistent_total,
        "cab": int((changes["Cambio"].fillna("").astype(str).str.strip() == "Cambio Programado (CAB)").sum()),
        "agile": int((changes["Cambio"].fillna("").astype(str).str.strip() == "Cambio Agile (CAB Agile)").sum()),
        "emergent": int((changes["Cambio"].fillna("").astype(str).str.strip() == "Cambio No Programado (Emergente)").sum()),
        "standard": int((changes["Cambio"].fillna("").astype(str).str.strip() == "Cambio Estándar (Rutinario)").sum()),
        "avg_duration_h": avg_duration,
        "max_duration_h": max_duration,
        "peak_month": month_peak,
        "peak_month_value": max(month_counts) if month_counts else 0,
    }
    base_d["kpi"]["changes26"] = int(len(rows))
    base_d["kpi"]["changes_success"] = success_total
    base_d["kpi"]["changes_failed"] = failed_total
    base_d["kpi"]["changes_cancelled"] = cancel_total
    base_d["kpi"]["changes_inconsistent"] = inconsistent_total
    base_d["fuente_cambios"] = {
        "archivo": changes_source_name or changes_path.name,
        "hoja": change_sheet,
        "corte": (start_dt.dropna().max().normalize() if start_dt.notna().any() else as_of_date).strftime("%Y-%m-%d"),
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    return rows


def top_counts(series, limit=10):
    counts = {}
    for value, count in series.dropna().astype(str).str.strip().value_counts().items():
        if value and value != "nan":
            counts[value] = int(count)
    ordered = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))
    return OrderedDict(ordered[:limit])


def all_counts(series):
    counts = {}
    for value, count in series.dropna().astype(str).str.strip().value_counts().items():
        if value and value != "nan":
            counts[value] = int(count)
    return OrderedDict(sorted(counts.items(), key=lambda kv: (-kv[1], kv[0])))


def count_by_priority(df):
    return {p: int((df["Prioridad"] == p).sum()) for p in PRIOS}


def sla_rate(df):
    if df.empty:
        return 0.0
    vals = df["Cumple resolución"].fillna("").astype(str).str.strip()
    dentro = int((vals == "Dentro del SLA").sum())
    fuera = int((vals == "Fuera del SLA").sum())
    denom = dentro + fuera
    return round((dentro / denom * 100), 1) if denom else 0.0


def mtbf_dict(df, hours, limit=10):
    counts = all_counts(df["SubCategoria"])
    vals = []
    for sub, count in counts.items():
        if count:
            vals.append((sub, round(hours / count, 1)))
    return OrderedDict(sorted(vals, key=lambda kv: (kv[1], kv[0]))[:limit])


def excel_date_text(v):
    if pd.isna(v):
        return ""
    dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return clean(v)
    return dt.strftime("%d-%m-%Y")


def build_data(
    base_d,
    xlsx_path=XLSX_IN,
    source_name=None,
    problems_path=None,
    problems_source_name=None,
    changes_path=CHANGES_IN,
    changes_source_name=None,
):
    xlsx_path = Path(xlsx_path)
    with pd.ExcelFile(xlsx_path) as workbook:
        input_sheet = detect_input_sheet(workbook)
        df = pd.read_excel(workbook, sheet_name=input_sheet)
    missing_columns = sorted(REQUIRED_COLUMNS.difference(df.columns))
    if missing_columns:
        raise ValueError("Faltan columnas requeridas: " + ", ".join(missing_columns))
    df["Tipo ticket"] = df["Tipo ticket"].fillna("").astype(str).str.strip()
    df["Prioridad"] = df["Prioridad"].fillna("").astype(str).str.strip()
    creation_dates = pd.to_datetime(
        df["Fec. creación"], dayfirst=True, errors="coerce"
    )
    if "Mes" in df.columns:
        df["Mes"] = pd.to_numeric(df["Mes"], errors="coerce").fillna(0).astype(int)
    else:
        df["Mes"] = creation_dates.dt.month.fillna(0).astype(int)
    as_of_date = (
        creation_dates.dropna().max().normalize()
        if creation_dates.notna().any()
        else pd.Timestamp.today().normalize()
    )
    valid_months = df.loc[df["Mes"].between(1, 12), "Mes"]
    last_month = int(valid_months.max()) if not valid_months.empty else 1
    months = list(range(1, last_month + 1))
    total_hours = sum(MONTH_HOURS[month] for month in months)

    inc = df[df["Tipo ticket"] == "Incidente"].copy()
    req = df[df["Tipo ticket"] == "Requerimiento"].copy()
    ev = df[df["Tipo ticket"] == "Evento"].copy()
    con = df[df["Tipo ticket"] == "Consultas"].copy()
    p1 = inc[inc["Prioridad"] == "P1- Critico"].copy()

    base_d["kpi"]["tickets26"] = int(len(df))
    base_d["kpi"]["inc26"] = int(len(inc))
    base_d["kpi"]["req26"] = int(len(req))
    base_d["kpi"]["ev26"] = int(len(ev))
    base_d["kpi"]["p1_26"] = int(len(p1))
    base_d["kpi"]["p1_26_fuera"] = int((p1["Cumple resolución"].fillna("").astype(str).str.strip() == "Fuera del SLA").sum())
    base_d["kpi"]["sla26"] = sla_rate(inc)
    base_d["kpi"]["consultas26"] = int(len(con))

    base_d["prio26"] = count_by_priority(inc)
    base_d["sla_prio26"] = {p: sla_rate(inc[inc["Prioridad"] == p]) for p in PRIOS}
    base_d["inc_mes26"] = [int((inc["Mes"] == m).sum()) for m in months]
    base_d["req_mes26"] = [int((req["Mes"] == m).sum()) for m in months]
    base_d["p1_mes26"] = [int((p1["Mes"] == m).sum()) for m in months]
    base_d["cat26"] = all_counts(inc["Categoria"])
    base_d["subcat26"] = top_counts(inc["SubCategoria"])
    base_d["p1_sub26"] = top_counts(p1["SubCategoria"])
    base_d["p1_equipo26"] = top_counts(p1["Equipo"])

    evol26 = {"all": base_d["inc_mes26"]}
    for prio in PRIOS:
        evol26[prio] = [
            int(((inc["Mes"] == m) & (inc["Prioridad"] == prio)).sum())
            for m in months
        ]
    base_d["evol26"] = evol26

    inc_by_month = {}
    sub_mp = {"0": {"all": top_counts(inc["SubCategoria"])}}
    for prio in PRIOS:
        sub_mp["0"][prio] = top_counts(inc[inc["Prioridad"] == prio]["SubCategoria"])
    p1_mp = {
        "0": {
            "sub": top_counts(p1["SubCategoria"]),
            "eq": top_counts(p1["Equipo"]),
            "total": int(len(p1)),
        }
    }
    mtbf = {"0": mtbf_dict(inc, total_hours)}
    for m in months:
        inc_m = inc[inc["Mes"] == m]
        p1_m = p1[p1["Mes"] == m]
        inc_by_month[str(m)] = {
            "total": int(len(inc_m)),
            "prio": count_by_priority(inc_m),
            "sub": top_counts(inc_m["SubCategoria"]),
            "cat": all_counts(inc_m["Categoria"]),
        }
        sub_mp[str(m)] = {"all": top_counts(inc_m["SubCategoria"])}
        for prio in PRIOS:
            sub_mp[str(m)][prio] = top_counts(inc_m[inc_m["Prioridad"] == prio]["SubCategoria"])
        p1_mp[str(m)] = {
            "sub": top_counts(p1_m["SubCategoria"]),
            "eq": top_counts(p1_m["Equipo"]),
            "total": int(len(p1_m)),
        }
        mtbf[str(m)] = mtbf_dict(inc_m, MONTH_HOURS[m])

    base_d["inc_by_month"] = inc_by_month
    base_d["sub_mp"] = sub_mp
    base_d["p1_mp"] = p1_mp
    base_d["mtbf"] = mtbf

    mtbf_all = {}
    for sub, count in all_counts(inc["SubCategoria"]).items():
        mtbf_all[sub] = round(total_hours / count, 1)
    p1_rows = []
    p1_sorted = p1.sort_values(["Mes", "Fec. creación", "ID Entrada"])
    for _, row in p1_sorted.iterrows():
        sub = clean(row.get("SubCategoria"))
        minutes = pd.to_numeric(row.get("Tiempo resolución"), errors="coerce")
        p1_rows.append({
            "id": clean(row.get("ID Entrada")),
            "fecha": excel_date_text(row.get("Fec. creación")),
            "titulo": clean(row.get("Título")),
            "cat": clean(row.get("Categoria")),
            "sub": sub,
            "equipo": clean(row.get("Equipo")),
            "mttr": None if pd.isna(minutes) else round(float(minutes) / 60, 1),
            "sla": clean(row.get("Cumple resolución")),
            "mes": int(row.get("Mes")),
            "mtbf": mtbf_all.get(sub),
        })
    base_d["p1_list"] = p1_rows

    problem_rows = base_d.get("prob_list", [])
    if problems_path:
        problems_path = Path(problems_path)
        if problems_path.is_file():
            problem_rows = build_problem_data(
                base_d,
                problems_path,
                problems_source_name,
                inc,
                as_of_date,
            )

    change_rows = base_d.get("change_list", [])
    if changes_path:
        changes_path = Path(changes_path)
        if changes_path.is_file():
            change_rows = build_change_data(
                base_d,
                changes_path,
                changes_source_name,
                as_of_date,
            )

    risk = base_d.get("riesgo", {})
    vals = risk.get("vals", {})
    vals["P1 incumplido"] = round(base_d["kpi"]["p1_26_fuera"] / max(1, len(p1[p1["Cumple resolución"].isin(["Dentro del SLA", "Fuera del SLA"])])) * 100, 1)
    vals["SLA global"] = round(100 - base_d["kpi"]["sla26"], 1)
    vals["Concentración"] = round((max(base_d["p1_sub26"].values()) if base_d["p1_sub26"] else 0) / max(1, base_d["kpi"]["p1_26"]) * 100, 1)
    if problem_rows:
        vals["Problemas crónicos"] = round(
            sum(1 for row in problem_rows if (row.get("age") or 0) > 120)
            / len(problem_rows)
            * 100,
            1,
        )
        vals["Falta mitigación"] = round(
            sum(1 for row in problem_rows if row.get("wa") == "No")
            / len(problem_rows)
            * 100,
            1,
        )
    risk["vals"] = vals
    risk.setdefault("desc", {})
    risk["desc"].setdefault("P1 incumplido", {})["formula"] = f"{base_d['kpi']['p1_26_fuera']} de {len(p1[p1['Cumple resolución'].isin(['Dentro del SLA', 'Fuera del SLA'])])} P1 fuera de SLA"
    fuera_inc = int((inc["Cumple resolución"].fillna("").astype(str).str.strip() == "Fuera del SLA").sum())
    evaluados_inc = int(inc["Cumple resolución"].fillna("").astype(str).str.strip().isin(["Dentro del SLA", "Fuera del SLA"]).sum())
    risk["desc"].setdefault("SLA global", {})["formula"] = f"{fuera_inc} de {evaluados_inc} incidentes fuera de SLA"
    top_p1 = max(base_d["p1_sub26"].values()) if base_d["p1_sub26"] else 0
    risk["desc"].setdefault("Concentración", {})["formula"] = f"{top_p1} de {base_d['kpi']['p1_26']} P1 en el servicio top"
    if problem_rows:
        chronic = sum(1 for row in problem_rows if (row.get("age") or 0) > 120)
        without_workaround = sum(
            1 for row in problem_rows if row.get("wa") == "No"
        )
        risk["desc"].setdefault("Problemas crónicos", {})[
            "formula"
        ] = f"{chronic} de {len(problem_rows)} problemas >120 días"
        risk["desc"].setdefault("Falta mitigación", {})[
            "formula"
        ] = f"{without_workaround} de {len(problem_rows)} sin mitigación"
    max_eq = max(base_d["p1_equipo26"].values()) if base_d["p1_equipo26"] else 1
    risk["equipo"] = {k: round(v / max_eq * 100) for k, v in base_d["p1_equipo26"].items()}
    max_cat = max(base_d["cat26"].values()) if base_d["cat26"] else 1
    risk["categoria"] = {k: round(v / max_cat * 100) for k, v in base_d["cat26"].items()}
    base_d["riesgo"] = risk

    base_d["fuente_excel"] = {
        "archivo": source_name or xlsx_path.name,
        "hoja": input_sheet,
        "corte": as_of_date.strftime("%Y-%m-%d"),
        "criterio": "Tipo ticket oficial del Excel; Incidente/Requerimiento/Evento/Consultas sin reclasificación manual.",
        "generado": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    return base_d


def update_dashboard(
    xlsx_path=XLSX_IN,
    html_path=HTML_OUT,
    source_name=None,
    problems_path=PROBLEMS_IN,
    problems_source_name=None,
    changes_path=CHANGES_IN,
    changes_source_name=None,
):
    xlsx_path = Path(xlsx_path).resolve()
    html_path = Path(html_path).resolve()
    source_name = source_name or xlsx_path.name
    source_label = escape(source_name)
    html = html_path.read_text(encoding="utf-8")
    match = re.search(r"const D=(.*?);</script>", html, re.S)
    if not match:
        raise RuntimeError("No se encontró const D en el HTML")
    base_d = json.loads(match.group(1))
    data = build_data(
        base_d,
        xlsx_path,
        source_name,
        problems_path,
        problems_source_name,
        changes_path,
        changes_source_name,
    )
    encoded = json.dumps(data, ensure_ascii=False, separators=(",", ":"))
    html = html[: match.start(1)] + encoded + html[match.end(1):]
    html = re.sub(
        r"<title>.*?</title>",
        "<title>CEN · Dashboard Ejecutivo de Gestión TI</title>",
        html,
        count=1,
    )
    html = re.sub(
        r"<h1>Dashboard Ejecutivo de Gestión TI.*?</h1>",
        "<h1>Dashboard Ejecutivo de Gestión TI</h1>",
        html,
        count=1,
    )
    html = html.replace(
        "Estado de la operación · Incidentes · Problemas · Cambios",
        "Estado de la operación · Datos oficiales desde Excel",
    )
    cutoff = data.get("fuente_excel", {}).get("corte", "")
    cutoff_text = (
        datetime.strptime(cutoff, "%Y-%m-%d").strftime("%d-%m-%Y")
        if cutoff
        else "sin fecha"
    )
    html = re.sub(
        r"Corte: [^<·]+ · Fuente: [^<]*?\.xlsx",
        f"Corte: {cutoff_text} · Fuente: {source_label}",
        html,
    )
    html = re.sub(
        r"Dashboard actualizado desde [^<]*?\.xlsx · Datos de carácter confidencial",
        f"Dashboard actualizado desde {source_label} · Datos de carácter confidencial",
        html,
    )
    temporary_output = html_path.with_suffix(html_path.suffix + ".tmp")
    temporary_output.write_text(html, encoding="utf-8", newline="\n")
    temporary_output.replace(html_path)
    return data


def main():
    parser = argparse.ArgumentParser(description="Actualiza el dashboard CEN desde un Excel.")
    parser.add_argument("--xlsx", type=Path, default=XLSX_IN)
    parser.add_argument("--html", type=Path, default=HTML_OUT)
    parser.add_argument("--source-name")
    parser.add_argument("--problems", type=Path, default=PROBLEMS_IN)
    parser.add_argument("--problems-source-name")
    parser.add_argument("--changes", type=Path, default=CHANGES_IN)
    parser.add_argument("--changes-source-name")
    args = parser.parse_args()
    update_dashboard(
        args.xlsx,
        args.html,
        args.source_name,
        args.problems,
        args.problems_source_name,
        args.changes,
        args.changes_source_name,
    )


if __name__ == "__main__":
    main()
