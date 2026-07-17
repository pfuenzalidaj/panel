# Especificación de fuentes

## 1. Tickets e incidentes

La hoja se detecta automáticamente. Se priorizan nombres que contengan
`Tickets_full` y `CONFIDENCIAL`.

Columnas principales:

- `ID Entrada`
- `Tipo ticket`
- `Prioridad`
- `Fec. creación`
- `Categoria`
- `SubCategoria`
- `Equipo`
- `Cumple resolución`
- `Tiempo resolución`

La columna `Mes` es opcional. Si no existe, se calcula desde `Fec. creación`.

## 2. Módulo de Problemas

La hoja se detecta automáticamente. Se priorizan nombres con el patrón
`Informe_modulo_Problemas_*`.

Columnas principales:

- `ID Problema`
- `Estado`
- `Fec. creación`
- `Fecha de Compromisos`
- `SubCategoria`
- `Breve descripción`
- `Descripción (historial)`
- `Causa raíz`
- `Asignado a`
- `Jefe Dpto`
- `Sub Gerencia`
- `Prioridad`
- `Workaround`
- `Tipo de Compromiso`
- `Seguimiento`
- `Tickets Asociados`

`Semaforo` es opcional. Cuando no existe, el nivel de riesgo se deriva desde
`Prioridad`.

## Criterios de compromiso

- **Vencido:** fecha anterior al corte de tickets.
- **Vence hoy:** fecha igual al corte.
- **Próximo:** vence dentro de 30 días.
- **En plazo:** vence después de 30 días.
- **Sin fecha:** no existe compromiso formal registrado.
