# AQF_10 — Reporte de Mantenciones Preventivas (MP) 2026

Aplicación **HTML autónoma** (un solo archivo) generada a partir de
`ProgramacionMP2026.xlsm`. Funciona sin conexión y permite **importar el `.xlsm`
desde el navegador** para actualizar los datos en vivo.

## Contenido

| Archivo | Descripción |
|---|---|
| `generar_reporte.py` | Lee el `.xlsm` e incrusta los datos + SheetJS en el HTML. |
| `Reporte_MP_2026.html` | App autónoma: matriz por equipo, eventos, importar, filtros, drill-down. |
| `vendor/xlsx.full.min.js` | Librería [SheetJS](https://sheetjs.com) (lectura de `.xlsx/.xlsm` en el navegador). |
| `data/ProgramacionMP2026.xlsm` | Archivo fuente (datos por defecto). |

## Funcionalidades

- **Importar `.xlsm`** (botón «⭱ Importar .xlsm»): re-lee la hoja `Registro_MP-2026`
  y recalcula todo en vivo. Los datos por defecto y los importados pasan por la
  **misma** función `transform()` en JavaScript, garantizando resultados idénticos.
- **Vista «Matriz por equipo»**: una fila por equipo, **una columna por mes**
  (Ene–Dic) coloreada por categoría, más columnas de conteo: **Sí · C1–C8 · No ·
  NU · Baja · No registrado**.
- **Vista «Eventos»**: una fila por evento (equipo × mes con actividad), con todos
  los campos del equipo + Mes, Programa, Resultado, Fecha de Ejecución, Estado del
  Equipo, Cant. Pendientes y Última Actualización.
- **Todo es cliqueable**: celda de mes → ficha del evento; conteo → lista de esos
  eventos; ID/equipo → línea de tiempo anual; fila de evento → ficha; tarjetas y
  chips de leyenda → filtran; encabezados → ordenan.
- Búsqueda, filtros (Servicio, Clasificación, Mes, Resultado) y **exportar CSV**
  de la vista activa.

## Fuente y reglas de lectura

- Hoja **`Registro_MP-2026`** (contiene Programa **P** y Resultado **R** por mes;
  su columna P coincide en 99,99 % con `PMP_2026`).
- Datos desde la **fila 7**, columnas **B–AQ**, ignorando **Q** (Observación) y
  **S** (Responsable MP).
- Los **N° de Serie / Inventario** se conservan tal cual (texto), respetando los
  **ceros a la izquierda** (p. ej. `0024` ≠ `24`).

### Códigos

- **Programa (P):** `X` Programada · `R` Reprogramada · `RA` Reprog. año anterior ·
  `PM` Puesta en Marcha.
- **Resultado (R):** `Si` Realizada · `Si-RA` Año anterior realizada · `C1–C8`
  Reprogramada (causal) · `FS` Fuera de servicio · `No` No realizada · `NU` No
  ubicable · `Baja` Dado de baja.

### Columnas de conteo y derivadas

- **Sí** = realizadas (`Si`/`Si-RA`) · **C1–C8** = reprogramadas · **No / NU /
  Baja** = según resultado · **No registrado** = meses transcurridos con MP
  programada y **sin** resultado registrado.
- **Fecha de Ejecución:** mes cuando el resultado es `Si`/`Si-RA` (la fuente opera
  a nivel de mes, no de día).
- **Estado del Equipo:** derivado del resultado; si sólo hay programación →
  `Pendiente` (mes transcurrido) o `Programada` (futuro).
- Los **meses transcurridos** se calculan con la **fecha del navegador** respecto
  al año `YEAR` (2026).

## Regenerar el HTML (opcional)

Sólo es necesario si cambian la plantilla o los datos por defecto; para el uso
diario basta con abrir el HTML e **importar** el `.xlsm` actualizado.

```bash
pip install openpyxl
python3 generar_reporte.py                 # usa data/ProgramacionMP2026.xlsm
python3 generar_reporte.py /ruta/archivo.xlsm
```
