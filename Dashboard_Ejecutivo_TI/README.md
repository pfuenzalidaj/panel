# Dashboard Ejecutivo de Gestión TI

Aplicación local para generar el dashboard ejecutivo del CEN desde las fuentes
oficiales de tickets y gestión de problemas.

## Uso

1. Ejecute `Iniciar_Actualizador_Dashboard.bat`.
2. Cargue el Excel de tickets e incidentes.
3. Cargue, si corresponde, el Excel del módulo de Problemas.
4. Seleccione **Actualizar dashboard**.
5. Abra o descargue el HTML generado.

Los nombres de los archivos y los sufijos de las hojas pueden variar. La
aplicación identifica las fuentes por su estructura de columnas.

## Estructura

```text
Dashboard_Ejecutivo_TI/
├── Iniciar_Actualizador_Dashboard.bat
├── Actualizar_Dashboard_Directo.bat
├── README.md
├── VERSION.txt
├── manifiesto.json
├── aplicacion/
│   ├── servidor_actualizador.py
│   ├── generador_dashboard.py
│   └── interfaz_actualizador.html
├── dashboard/
│   └── Dashboard_Ejecutivo_Gestion_TI.html
├── fuentes/
│   ├── Tickets_Incidentes.xlsx
│   └── Informe_Problemas.xlsx
├── documentacion/
│   ├── ESPECIFICACION_DE_FUENTES.md
│   └── REGISTRO_DE_CAMBIOS.md
└── temporales/
    └── README.txt
```

## Principios del entregable

- Rutas relativas: la carpeta completa puede trasladarse.
- Fuentes variables: no depende del nombre exacto del archivo o de la hoja.
- Actualización segura: el dashboard se reemplaza de forma atómica.
- Temporales aislados: las cargas se procesan en `temporales/` y se eliminan.
- Trazabilidad: el dashboard conserva nombre, hoja y fecha de cada fuente.

## Requisitos

- Windows.
- Python con `pandas` y soporte para Excel.
- Navegador web moderno.

El lanzador utiliza primero el runtime incluido en el entorno Codex y, si no
está disponible, intenta utilizar `python` desde el sistema.
