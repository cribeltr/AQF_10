# AQF_10 — Reporte de Mantenciones Preventivas (MP) 2026

Generación de un reporte HTML autónomo a partir del archivo
`ProgramacionMP2026.xlsm`, con **una fila por evento de mantención**
(cada combinación de equipo × mes que tenga programación y/o resultado).

## Contenido

| Archivo | Descripción |
|---|---|
| `generar_reporte.py` | Script que lee el `.xlsm` y produce el HTML. |
| `Reporte_MP_2026.html` | Reporte autónomo (búsqueda, filtros, orden, exportar CSV). |
| `data/ProgramacionMP2026.xlsm` | Archivo fuente. |

## Columnas del reporte (una fila por evento)

ID · N° de Carpeta · N° de Inventario · Nombre del Equipo · Servicio · Unidad ·
Ubicación · Procedencia · Marca · Modelo · N° de Serie · Año de Instalación ·
Vida Útil Residual · Clasificación · ENU / Baja · **Mes · Programa · Resultado ·
Fecha de Ejecución · Estado del Equipo · Cantidad de Pendientes · Última
Actualización**.

## Fuente y reglas de lectura

- Hoja **`Registro_MP-2026`** (fuente única: contiene Programa **P** y Resultado
  **R** por mes; su columna P coincide en un 99,99 % con la hoja `PMP_2026`).
- Datos desde la **fila 7**, columnas **B–AQ**, ignorando **Q** (Observación) y
  **S** (Responsable MP).
- Los **N° de Serie / Inventario** se conservan tal cual (texto), respetando los
  **ceros a la izquierda** (p. ej. `0024` no se transforma en `24`).

### Códigos

- **Programa (P):** `X` Programada · `R` Reprogramada · `RA` Reprogramada año
  anterior · `PM` Puesta en Marcha.
- **Resultado (R):** `Si` Realizada · `Si-RA` Año anterior realizada · `C1–C8`
  Reprogramada (causal) · `FS` Fuera de servicio · `No` No realizada · `NU` No
  ubicable · `Baja` Dado de baja.

### Columnas derivadas (ver «Notas metodológicas» dentro del HTML)

- **Fecha de Ejecución:** mes de ejecución cuando el resultado es `Si`/`Si-RA`
  (la fuente opera a nivel de mes, no de día).
- **Estado del Equipo:** derivado del código de resultado del evento; si sólo
  hay programación, `Pendiente` (mes transcurrido) o `Programada` (mes futuro).
- **Cantidad de Pendientes (por equipo):** meses transcurridos con MP programada
  cuyo resultado no es `Si`/`Si-RA`.
- **Última Actualización (por equipo):** mes más reciente con resultado
  registrado.

## Regenerar el reporte

```bash
pip install openpyxl
python3 generar_reporte.py            # usa data/ProgramacionMP2026.xlsm
# o bien:
python3 generar_reporte.py /ruta/al/archivo.xlsm
```

La fecha de referencia para los cálculos de pendientes/meses transcurridos está
en la constante `REF_DATE` dentro de `generar_reporte.py`.
