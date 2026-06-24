# -*- coding: utf-8 -*-
"""
Generador del Reporte de Mantenciones Preventivas (MP) 2026
============================================================

Lee el archivo ``ProgramacionMP2026.xlsm`` y genera un archivo HTML autónomo
(``Reporte_MP_2026.html``) en el que se produce **una fila por evento de
mantención**, es decir, una fila por cada combinación de (equipo × mes) que
tenga programación y/o resultado registrado.

Cada fila contiene los campos de la base de datos del equipo más los campos
del evento:

    ID, N° de Carpeta, N° de Inventario, Nombre del Equipo, Servicio, Unidad,
    Ubicación, Procedencia, Marca, Modelo, N° de Serie, Año de Instalación,
    Vida Útil Residual, Clasificación, ENU / Baja, Mes, Programa, Resultado,
    Fecha de Ejecución, Estado del Equipo, Cantidad de pendientes,
    Última actualización.

Fuente de datos
---------------
Se utiliza la hoja ``Registro_MP-2026`` como fuente única de verdad, porque
contiene tanto la subcolumna de Programa (P) como la de Resultado (R) para los
doce meses. (Su columna P coincide en un 99,99 % con la programación de la hoja
``PMP_2026``.) Los datos se leen desde la fila 8 (encabezado en la fila 7),
columnas B a AQ, ignorando las columnas Q (Observación) y S (Responsable MP),
tal como indica la especificación.

Números de serie / inventario
------------------------------
Se conservan exactamente como están registrados (son texto en el origen), de
modo que se respetan los ceros a la izquierda (p. ej. "0024" no se transforma
en "24").
"""

from __future__ import annotations

import datetime
import html
import json
import os

import openpyxl

# --------------------------------------------------------------------------- #
# Configuración
# --------------------------------------------------------------------------- #

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSM = os.path.join(BASE_DIR, "data", "ProgramacionMP2026.xlsm")
OUTPUT_HTML = os.path.join(BASE_DIR, "Reporte_MP_2026.html")
SHEET = "Registro_MP-2026"

# Fecha de referencia para clasificar meses transcurridos vs. futuros y para
# calcular la "Cantidad de pendientes". Corresponde a la fecha de generación.
REF_DATE = datetime.date(2026, 6, 24)
CURRENT_MONTH = REF_DATE.month  # 6 = junio
YEAR = 2026

# Encabezado en la fila 7, datos desde la fila 8.
HEADER_ROW = 7
FIRST_DATA_ROW = 8

# Mapa de columnas de identidad (índices de columna de openpyxl, 1-based).
# Se ignoran Q (17, Observación) y S (19, Responsable MP).
IDENT_COLS = {
    "id":           2,   # B  - ID
    "carpeta":      3,   # C  - N° Carpeta
    "inventario":   4,   # D  - N° Inventario
    "equipo":       5,   # E  - Equipo
    "servicio":     6,   # F  - Servicio
    "unidad":       7,   # G  - Unidad
    "ubicacion":    8,   # H  - Ubicación
    "procedencia":  9,   # I  - Procedencia
    "marca":        10,  # J  - Marca
    "modelo":       11,  # K  - Modelo
    "serie":        12,  # L  - Serie
    "anio":         13,  # M  - Año Instalación
    "vida_util":    14,  # N  - Vida Útil Residual
    "clasif":       15,  # O  - Clasificación
    "enu_baja":     16,  # P  - ENU / Baja
    "frecuencia":   18,  # R  - Frecuencia MP
}

# Doce meses: para cada mes, columna P (Programa) y columna R (Resultado).
# T=20,U=21 (Ene) ... AP=42,AQ=43 (Dic).
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
MES_ABR = ["Ene", "Feb", "Mar", "Abr", "May", "Jun",
           "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
MONTH_P_COLS = list(range(20, 43, 2))   # 20,22,...,42
MONTH_R_COLS = list(range(21, 44, 2))   # 21,23,...,43

# Columnas placeholder donde el literal "0" significa "vacío".
PLACEHOLDER_ZERO = {"carpeta", "inventario", "servicio", "unidad",
                    "ubicacion", "procedencia", "enu_baja", "frecuencia"}

# --------------------------------------------------------------------------- #
# Diccionarios de códigos (según especificación)
# --------------------------------------------------------------------------- #

PROGRAMA_LABEL = {
    "X":  "MP Programada",
    "R":  "MP Reprogramada",
    "RA": "MP Reprogramada de año anterior",
    "PM": "Puesta en Marcha",
}

CAUSAS = {
    "C1": "Imposibilidad de desocupar el equipo del paciente por indicación clínica",
    "C2": "Equipo en servicio técnico",
    "C3": "Equipo no operativo, a la espera de repuestos o accesorios",
    "C4": "Equipo en préstamo a otro hospital o institución",
    "C5": "No disponibilidad de horas hombre del funcionario SEC por alta carga laboral",
    "C6": "No disponibilidad de horas hombre del servicio técnico externo",
    "C7": "Ausencia justificada del funcionario SEC superior a 15 días",
    "C8": "Contingencia hospitalaria",
}

RESULTADO_LABEL = {
    "Si":    "MP Preventiva Realizada",
    "Si-RA": "MP de año anterior realizada",
    "FS":    "Fuera de Servicio",
    "No":    "No Realizada",
    "NU":    "No Ubicable",
    "Baja":  "Equipo Dado de Baja",
}
for _c, _d in CAUSAS.items():
    RESULTADO_LABEL[_c] = f"Reprogramada — {_d}"

# Estado del equipo derivado del código de resultado.
# (clave_estado, etiqueta visible)
ESTADO_POR_RESULTADO = {
    "Si":    ("ok",       "Operativo (MP realizada)"),
    "Si-RA": ("ok",       "Operativo (MP año anterior realizada)"),
    "Baja":  ("baja",     "Dado de baja"),
    "FS":    ("critico",  "Fuera de servicio"),
    "NU":    ("alerta",   "No ubicable"),
    "No":    ("critico",  "MP no realizada"),
    "C1":    ("aviso",    "En uso clínico (no liberable)"),
    "C2":    ("alerta",   "En servicio técnico"),
    "C3":    ("critico",  "No operativo (espera repuestos)"),
    "C4":    ("info",     "En préstamo a otra institución"),
    "C5":    ("aviso",    "Operativo (sin HH SEC)"),
    "C6":    ("aviso",    "Operativo (sin HH externo)"),
    "C7":    ("aviso",    "Operativo (ausencia SEC)"),
    "C8":    ("aviso",    "Operativo (contingencia)"),
}

# Resultados que cuentan como ejecución exitosa de la MP.
REALIZADOS = {"Si", "Si-RA"}


# --------------------------------------------------------------------------- #
# Utilidades de lectura
# --------------------------------------------------------------------------- #

def cell_text(value) -> str:
    """Convierte un valor de celda a texto limpio, preservando el contenido.

    Importante: NO altera números de serie / inventario; los ceros a la
    izquierda se preservan porque en el origen son cadenas de texto.
    """
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_field(key: str, value) -> str:
    """Limpia un campo de identidad. En columnas placeholder, "0" => vacío."""
    txt = cell_text(value)
    if key in PLACEHOLDER_ZERO and txt == "0":
        return ""
    return txt


def last_data_row(ws) -> int:
    last = HEADER_ROW
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        if cell_text(ws.cell(row=r, column=IDENT_COLS["id"]).value):
            last = r
    return last


# --------------------------------------------------------------------------- #
# Construcción de eventos
# --------------------------------------------------------------------------- #

def build_events(xlsm_path: str):
    wb = openpyxl.load_workbook(xlsm_path, data_only=True)
    ws = wb[SHEET]
    last = last_data_row(ws)

    events = []
    equipos_con_eventos = set()

    for r in range(FIRST_DATA_ROW, last + 1):
        ident = {k: clean_field(k, ws.cell(row=r, column=c).value)
                 for k, c in IDENT_COLS.items()}

        # Saltar filas que no representan un equipo real (slot disponible).
        if not ident["equipo"] and not ident["serie"] and not ident["inventario"]:
            continue

        # Leer los 12 pares (Programa, Resultado).
        meses_pr = []
        for i in range(12):
            p = cell_text(ws.cell(row=r, column=MONTH_P_COLS[i]).value)
            res = cell_text(ws.cell(row=r, column=MONTH_R_COLS[i]).value)
            meses_pr.append((p, res))

        # --- Agregados por equipo ---
        # Cantidad de pendientes: meses transcurridos (Ene..mes actual) con
        # programación y sin ejecución exitosa (resultado no en {Si, Si-RA}),
        # excluyendo celdas dadas de baja.
        pendientes = 0
        ultima_idx = -1  # índice del mes más reciente con resultado registrado
        tiene_baja = False
        for i, (p, res) in enumerate(meses_pr):
            if res:
                ultima_idx = i
            if res == "Baja":
                tiene_baja = True
            mes_num = i + 1
            if p and mes_num <= CURRENT_MONTH:
                if res not in REALIZADOS and res != "Baja":
                    pendientes += 1

        ultima_actualizacion = (
            f"{MESES[ultima_idx]} {YEAR}" if ultima_idx >= 0 else ""
        )

        # --- Un evento por mes con programación o resultado ---
        for i, (p, res) in enumerate(meses_pr):
            if not p and not res:
                continue
            mes_num = i + 1
            equipos_con_eventos.add(ident["id"])

            # Fecha de ejecución: la fuente registra a nivel de mes. Si la MP
            # fue realizada (Si / Si-RA), la ejecución ocurrió en ese mes.
            if res in REALIZADOS:
                fecha_ejec = f"{MESES[i]} {YEAR}"
            else:
                fecha_ejec = ""

            # Estado del equipo.
            if res in ESTADO_POR_RESULTADO:
                estado_key, estado_txt = ESTADO_POR_RESULTADO[res]
            elif res:  # resultado desconocido pero presente
                estado_key, estado_txt = ("info", res)
            else:
                # Sólo programado, sin resultado.
                if mes_num <= CURRENT_MONTH:
                    estado_key, estado_txt = ("pendiente", "Pendiente")
                else:
                    estado_key, estado_txt = ("programado", "Programada")

            programa_label = PROGRAMA_LABEL.get(p, p)
            resultado_label = RESULTADO_LABEL.get(res, res)

            events.append({
                "id": ident["id"],
                "carpeta": ident["carpeta"],
                "inventario": ident["inventario"],
                "equipo": ident["equipo"],
                "servicio": ident["servicio"],
                "unidad": ident["unidad"],
                "ubicacion": ident["ubicacion"],
                "procedencia": ident["procedencia"],
                "marca": ident["marca"],
                "modelo": ident["modelo"],
                "serie": ident["serie"],
                "anio": ident["anio"],
                "vida_util": ident["vida_util"],
                "clasif": ident["clasif"],
                "enu_baja": ident["enu_baja"],
                "mes": MESES[i],
                "mes_idx": mes_num,
                "programa": p,
                "programa_label": programa_label,
                "resultado": res,
                "resultado_label": resultado_label,
                "fecha_ejec": fecha_ejec,
                "estado_key": estado_key,
                "estado_txt": estado_txt,
                "pendientes": pendientes,
                "ultima_actualizacion": ultima_actualizacion,
            })

    stats = {
        "total_eventos": len(events),
        "total_equipos": len({e["id"] for e in events}),
        "realizadas": sum(1 for e in events if e["resultado"] in REALIZADOS),
        "reprogramadas": sum(1 for e in events if e["resultado"] in CAUSAS),
        "pendientes_eventos": sum(
            1 for e in events
            if not e["resultado"] and e["mes_idx"] <= CURRENT_MONTH
        ),
        "programadas_total": sum(1 for e in events if e["programa"]),
    }
    return events, stats


# --------------------------------------------------------------------------- #
# Render HTML
# --------------------------------------------------------------------------- #

COLUMNS = [
    ("id",                   "ID"),
    ("carpeta",              "N° Carpeta"),
    ("inventario",           "N° Inventario"),
    ("equipo",               "Nombre del Equipo"),
    ("servicio",             "Servicio"),
    ("unidad",               "Unidad"),
    ("ubicacion",            "Ubicación"),
    ("procedencia",          "Procedencia"),
    ("marca",                "Marca"),
    ("modelo",               "Modelo"),
    ("serie",                "N° de Serie"),
    ("anio",                 "Año Inst."),
    ("vida_util",            "Vida Útil Res."),
    ("clasif",               "Clasificación"),
    ("enu_baja",             "ENU / Baja"),
    ("mes",                  "Mes"),
    ("programa",             "Programa"),
    ("resultado",            "Resultado"),
    ("fecha_ejec",           "Fecha de Ejecución"),
    ("estado_txt",           "Estado del Equipo"),
    ("pendientes",           "Cant. Pendientes"),
    ("ultima_actualizacion", "Última Actualización"),
]


def render_html(events, stats, xlsm_path: str) -> str:
    data_json = json.dumps(events, ensure_ascii=False)
    columns_json = json.dumps(COLUMNS, ensure_ascii=False)
    gen_fecha = REF_DATE.strftime("%d-%m-%Y")
    src_name = os.path.basename(xlsm_path)

    causas_rows = "\n".join(
        f"<tr><td><span class='code'>{c}</span></td><td>{html.escape(d)}</td></tr>"
        for c, d in CAUSAS.items()
    )
    programa_rows = "\n".join(
        f"<tr><td><span class='code'>{html.escape(c)}</span></td><td>{html.escape(d)}</td></tr>"
        for c, d in PROGRAMA_LABEL.items()
    )
    resultado_rows = "\n".join(
        f"<tr><td><span class='code'>{html.escape(c)}</span></td><td>{html.escape(d)}</td></tr>"
        for c, d in [
            ("Si", "Mantención Preventiva Realizada"),
            ("C1–C8", "Mantención Preventiva Reprogramada (ver causales)"),
            ("Si-RA", "Mantención de Año Anterior Realizada"),
            ("FS", "Fuera de Servicio"),
            ("No", "No Realizada"),
            ("NU", "No Ubicable"),
            ("Baja", "Equipo Dado de Baja"),
        ]
    )

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Reporte Mantenciones Preventivas 2026 — H.H.H.A.</title>
<style>
  :root {{
    --bg: #f4f6f9; --panel: #ffffff; --ink: #1f2733; --muted: #647084;
    --line: #e3e8ef; --brand: #1565c0; --brand-dark: #0d3c75;
    --ok: #1b873f; --ok-bg: #e7f6ec; --pend: #b54708; --pend-bg: #fff4e5;
    --crit: #c01525; --crit-bg: #fdecec; --aviso: #8a6d00; --aviso-bg: #fff8db;
    --info: #155e9c; --info-bg: #e7f1fb; --baja: #475467; --baja-bg: #eceef2;
    --prog: #344054; --prog-bg: #eef1f5; --alerta:#9a4a00; --alerta-bg:#ffefe0;
  }}
  * {{ box-sizing: border-box; }}
  body {{
    margin: 0; background: var(--bg); color: var(--ink);
    font-family: "Segoe UI", Roboto, system-ui, -apple-system, sans-serif;
    font-size: 13px;
  }}
  header.top {{
    background: linear-gradient(135deg, var(--brand-dark), var(--brand));
    color: #fff; padding: 18px 24px;
  }}
  header.top h1 {{ margin: 0 0 4px; font-size: 20px; }}
  header.top p {{ margin: 0; opacity: .9; font-size: 12.5px; }}
  .wrap {{ padding: 16px 24px 60px; }}
  .cards {{ display: flex; flex-wrap: wrap; gap: 12px; margin: 16px 0; }}
  .card {{
    background: var(--panel); border: 1px solid var(--line); border-radius: 10px;
    padding: 12px 16px; min-width: 130px; flex: 1; box-shadow: 0 1px 2px rgba(16,24,40,.04);
  }}
  .card .n {{ font-size: 24px; font-weight: 700; color: var(--brand-dark); }}
  .card .l {{ font-size: 11.5px; color: var(--muted); text-transform: uppercase; letter-spacing: .03em; }}
  .toolbar {{
    background: var(--panel); border: 1px solid var(--line); border-radius: 10px;
    padding: 12px; display: flex; flex-wrap: wrap; gap: 10px; align-items: center;
    margin-bottom: 14px; position: sticky; top: 0; z-index: 20;
    box-shadow: 0 1px 2px rgba(16,24,40,.04);
  }}
  .toolbar input[type=search], .toolbar select {{
    padding: 7px 10px; border: 1px solid var(--line); border-radius: 7px;
    font-size: 13px; background: #fff; color: var(--ink);
  }}
  .toolbar input[type=search] {{ min-width: 240px; flex: 1; }}
  .toolbar .count {{ margin-left: auto; color: var(--muted); font-size: 12px; }}
  .btn {{
    padding: 7px 12px; border: 1px solid var(--line); background: #fff;
    border-radius: 7px; cursor: pointer; font-size: 12.5px; color: var(--ink);
  }}
  .btn:hover {{ background: #f0f4f9; }}
  .table-scroll {{
    overflow: auto; max-height: 70vh; border: 1px solid var(--line);
    border-radius: 10px; background: var(--panel);
  }}
  table {{ border-collapse: collapse; width: 100%; font-size: 12.3px; }}
  thead th {{
    position: sticky; top: 0; background: #f0f3f8; color: var(--brand-dark);
    text-align: left; padding: 9px 10px; border-bottom: 2px solid var(--line);
    white-space: nowrap; cursor: pointer; user-select: none; z-index: 5;
  }}
  thead th:hover {{ background: #e6ecf5; }}
  thead th .arrow {{ opacity: .4; font-size: 10px; }}
  tbody td {{ padding: 7px 10px; border-bottom: 1px solid var(--line); white-space: nowrap; }}
  tbody tr:hover {{ background: #f7f9fc; }}
  td.serie {{ font-family: "Consolas", "Courier New", monospace; }}
  .badge {{
    display: inline-block; padding: 2px 8px; border-radius: 999px;
    font-size: 11px; font-weight: 600; white-space: nowrap;
  }}
  .b-ok {{ color: var(--ok); background: var(--ok-bg); }}
  .b-pendiente {{ color: var(--pend); background: var(--pend-bg); }}
  .b-programado {{ color: var(--prog); background: var(--prog-bg); }}
  .b-critico {{ color: var(--crit); background: var(--crit-bg); }}
  .b-aviso {{ color: var(--aviso); background: var(--aviso-bg); }}
  .b-alerta {{ color: var(--alerta); background: var(--alerta-bg); }}
  .b-info {{ color: var(--info); background: var(--info-bg); }}
  .b-baja {{ color: #fff; background: var(--baja); }}
  .pill {{
    display:inline-block; min-width:20px; text-align:center; padding:1px 7px;
    border-radius:999px; font-weight:600; font-size:11px;
  }}
  .pill.zero {{ color: var(--muted); background: #f0f2f5; }}
  .pill.some {{ color: var(--pend); background: var(--pend-bg); }}
  details.legend {{
    margin-top: 22px; background: var(--panel); border: 1px solid var(--line);
    border-radius: 10px; padding: 4px 16px;
  }}
  details.legend > summary {{
    cursor: pointer; font-weight: 600; padding: 10px 0; color: var(--brand-dark);
    font-size: 14px;
  }}
  .legend-grid {{ display: flex; flex-wrap: wrap; gap: 24px; padding: 8px 0 16px; }}
  .legend-grid > div {{ flex: 1; min-width: 280px; }}
  .legend h4 {{ margin: 6px 0; color: var(--brand-dark); }}
  .legend table {{ width: 100%; font-size: 12px; }}
  .legend td {{ padding: 4px 8px; border-bottom: 1px solid var(--line); white-space: normal; }}
  .code {{
    font-family: "Consolas", monospace; font-weight: 700; color: var(--brand-dark);
    background: #eef3fb; padding: 1px 7px; border-radius: 5px;
  }}
  .notes {{ font-size: 12.3px; color: var(--muted); line-height: 1.6; }}
  .notes b {{ color: var(--ink); }}
  footer {{ text-align:center; color:var(--muted); font-size:11.5px; margin-top:24px; }}
  .muted {{ color: var(--muted); }}
</style>
</head>
<body>
<header class="top">
  <h1>Programa de Mantención Preventiva de Equipos Médicos Críticos — 2026</h1>
  <p>Hospital Hernán Henríquez Aravena · Una fila por evento de mantención (equipo × mes)</p>
</header>

<div class="wrap">
  <div class="cards">
    <div class="card"><div class="n">{stats['total_eventos']}</div><div class="l">Eventos</div></div>
    <div class="card"><div class="n">{stats['total_equipos']}</div><div class="l">Equipos</div></div>
    <div class="card"><div class="n">{stats['programadas_total']}</div><div class="l">Programadas</div></div>
    <div class="card"><div class="n">{stats['realizadas']}</div><div class="l">Realizadas</div></div>
    <div class="card"><div class="n">{stats['reprogramadas']}</div><div class="l">Reprogramadas</div></div>
    <div class="card"><div class="n">{stats['pendientes_eventos']}</div><div class="l">Pendientes (Ene–Jun)</div></div>
  </div>

  <div class="toolbar">
    <input type="search" id="q" placeholder="Buscar por equipo, serie, inventario, marca, servicio…">
    <select id="f_servicio"><option value="">Servicio: todos</option></select>
    <select id="f_mes"><option value="">Mes: todos</option></select>
    <select id="f_programa"><option value="">Programa: todos</option></select>
    <select id="f_resultado"><option value="">Resultado: todos</option></select>
    <select id="f_estado"><option value="">Estado: todos</option></select>
    <button class="btn" id="clear">Limpiar</button>
    <button class="btn" id="csv">Exportar CSV</button>
    <span class="count" id="count"></span>
  </div>

  <div class="table-scroll">
    <table id="tbl">
      <thead><tr id="head"></tr></thead>
      <tbody id="body"></tbody>
    </table>
  </div>

  <details class="legend">
    <summary>Leyenda de códigos y notas metodológicas</summary>
    <div class="legend-grid">
      <div class="legend">
        <h4>Programa (P)</h4>
        <table>{programa_rows}</table>
        <h4 style="margin-top:14px">Resultado (R)</h4>
        <table>{resultado_rows}</table>
      </div>
      <div class="legend">
        <h4>Causales de reprogramación</h4>
        <table>{causas_rows}</table>
      </div>
    </div>
    <div class="notes">
      <p><b>Fuente:</b> hoja <code>Registro_MP-2026</code> del archivo
      <code>{html.escape(src_name)}</code> (datos desde la fila 7, columnas B–AQ,
      ignorando las columnas Q «Observación» y S «Responsable MP»).</p>
      <p><b>Definición de evento:</b> se genera una fila por cada combinación de
      equipo y mes que tenga programación (P) y/o resultado (R). Los meses sin
      ninguno de los dos no generan fila.</p>
      <p><b>N° de Serie / N° de Inventario:</b> se reproducen exactamente como
      están registrados, conservando los ceros a la izquierda.</p>
      <p><b>Fecha de Ejecución:</b> el registro de origen opera a nivel de mes
      (no de día). Cuando el resultado es «Si» o «Si-RA» se muestra el mes de
      ejecución; en caso contrario queda vacía.</p>
      <p><b>Estado del Equipo:</b> se deriva del código de resultado del evento
      (p. ej. C2 → «En servicio técnico», C3 → «No operativo, espera repuestos»,
      FS → «Fuera de servicio», Baja → «Dado de baja»). Si el mes sólo tiene
      programación sin resultado: «Pendiente» si el mes ya transcurrió o
      «Programada» si es futuro.</p>
      <p><b>Cantidad de Pendientes (por equipo):</b> número de meses ya
      transcurridos (enero a {MESES[CURRENT_MONTH-1].lower()} de {YEAR}) con
      mantención programada cuyo resultado no es «Si» ni «Si-RA». Es un valor por
      equipo, por lo que se repite en todas sus filas.</p>
      <p><b>Última Actualización (por equipo):</b> mes más reciente en que el
      equipo tiene un resultado registrado.</p>
      <p><b>Reprogramación:</b> según las reglas, C2/C3/C4 no fijan nueva fecha
      (la MP se registra en el mes real de reingreso); C1/C5/C6/C7/C8 se
      reprograman dentro de 30 días — el mes original conserva «X» en P y la
      causal en R, y el mes destino lleva «R» en P.</p>
      <p class="muted">Reporte generado el {gen_fecha} · fecha de referencia para
      cálculos: {gen_fecha}.</p>
    </div>
  </details>

  <footer>Reporte autónomo · {stats['total_eventos']} eventos · generado el {gen_fecha}</footer>
</div>

<script>
const DATA = {data_json};
const COLUMNS = {columns_json};
const CURRENT_MONTH = {CURRENT_MONTH};

const body = document.getElementById('body');
const head = document.getElementById('head');
const countEl = document.getElementById('count');

// Encabezados
COLUMNS.forEach(([key, label]) => {{
  const th = document.createElement('th');
  th.dataset.key = key;
  th.innerHTML = label + ' <span class="arrow">↕</span>';
  head.appendChild(th);
}});

// Poblar filtros
function fillSelect(id, key) {{
  const sel = document.getElementById(id);
  const vals = [...new Set(DATA.map(d => d[key]).filter(v => v !== '' && v != null))];
  // ordenar meses por índice
  if (key === 'mes') {{
    const order = {json.dumps(MESES, ensure_ascii=False)};
    vals.sort((a,b) => order.indexOf(a) - order.indexOf(b));
  }} else {{
    vals.sort((a,b) => String(a).localeCompare(String(b), 'es'));
  }}
  vals.forEach(v => {{
    const o = document.createElement('option');
    o.value = v; o.textContent = v; sel.appendChild(o);
  }});
}}
fillSelect('f_servicio','servicio');
fillSelect('f_mes','mes');
fillSelect('f_programa','programa');
fillSelect('f_resultado','resultado');
fillSelect('f_estado','estado_txt');

let sortKey = null, sortDir = 1;
const ESTADO_CLASS = {{
  ok:'b-ok', pendiente:'b-pendiente', programado:'b-programado',
  critico:'b-critico', aviso:'b-aviso', alerta:'b-alerta', info:'b-info', baja:'b-baja'
}};

function esc(s) {{
  return String(s == null ? '' : s)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
}}

function currentFilters() {{
  return {{
    q: document.getElementById('q').value.trim().toLowerCase(),
    servicio: document.getElementById('f_servicio').value,
    mes: document.getElementById('f_mes').value,
    programa: document.getElementById('f_programa').value,
    resultado: document.getElementById('f_resultado').value,
    estado: document.getElementById('f_estado').value,
  }};
}}

function applyFilters() {{
  const f = currentFilters();
  let rows = DATA.filter(d => {{
    if (f.servicio && d.servicio !== f.servicio) return false;
    if (f.mes && d.mes !== f.mes) return false;
    if (f.programa && d.programa !== f.programa) return false;
    if (f.resultado && d.resultado !== f.resultado) return false;
    if (f.estado && d.estado_txt !== f.estado) return false;
    if (f.q) {{
      const hay = (d.equipo+' '+d.serie+' '+d.inventario+' '+d.carpeta+' '+
                   d.marca+' '+d.modelo+' '+d.servicio+' '+d.unidad+' '+
                   d.ubicacion+' '+d.procedencia).toLowerCase();
      if (!hay.includes(f.q)) return false;
    }}
    return true;
  }});
  if (sortKey) {{
    rows.sort((a,b) => {{
      let x=a[sortKey], y=b[sortKey];
      if (sortKey==='mes') {{ x=a.mes_idx; y=b.mes_idx; }}
      const nx=parseFloat(x), ny=parseFloat(y);
      if (!isNaN(nx) && !isNaN(ny) && String(x).trim()!=='' && String(y).trim()!=='') {{ x=nx; y=ny; }}
      if (x<y) return -1*sortDir; if (x>y) return 1*sortDir; return 0;
    }});
  }}
  render(rows);
}}

function render(rows) {{
  const frag = document.createDocumentFragment();
  for (const d of rows) {{
    const tr = document.createElement('tr');
    let h = '';
    for (const [key] of COLUMNS) {{
      if (key === 'estado_txt') {{
        const cls = ESTADO_CLASS[d.estado_key] || 'b-info';
        h += `<td><span class="badge ${{cls}}">${{esc(d.estado_txt)}}</span></td>`;
      }} else if (key === 'pendientes') {{
        const cls = d.pendientes > 0 ? 'some' : 'zero';
        h += `<td><span class="pill ${{cls}}">${{esc(d.pendientes)}}</span></td>`;
      }} else if (key === 'serie' || key === 'inventario') {{
        h += `<td class="serie">${{esc(d[key])}}</td>`;
      }} else if (key === 'programa') {{
        h += `<td title="${{esc(d.programa_label)}}">${{esc(d.programa)}}</td>`;
      }} else if (key === 'resultado') {{
        h += `<td title="${{esc(d.resultado_label)}}">${{esc(d.resultado) || '—'}}</td>`;
      }} else {{
        h += `<td>${{esc(d[key]) || '<span class=muted>—</span>'}}</td>`;
      }}
    }}
    tr.innerHTML = h;
    frag.appendChild(tr);
  }}
  body.innerHTML = '';
  body.appendChild(frag);
  countEl.textContent = rows.length + ' de ' + DATA.length + ' eventos';
}}

// Ordenamiento por encabezado
head.addEventListener('click', e => {{
  const th = e.target.closest('th'); if (!th) return;
  const key = th.dataset.key;
  if (sortKey === key) sortDir *= -1; else {{ sortKey = key; sortDir = 1; }}
  [...head.children].forEach(h => {{
    const a = h.querySelector('.arrow');
    a.textContent = (h.dataset.key===sortKey) ? (sortDir>0?'▲':'▼') : '↕';
    a.style.opacity = (h.dataset.key===sortKey) ? 1 : .4;
  }});
  applyFilters();
}});

['q','f_servicio','f_mes','f_programa','f_resultado','f_estado'].forEach(id => {{
  document.getElementById(id).addEventListener('input', applyFilters);
}});
document.getElementById('clear').addEventListener('click', () => {{
  document.getElementById('q').value='';
  ['f_servicio','f_mes','f_programa','f_resultado','f_estado'].forEach(id=>document.getElementById(id).value='');
  sortKey=null; applyFilters();
}});

// Exportar CSV (de lo filtrado)
document.getElementById('csv').addEventListener('click', () => {{
  const f = currentFilters();
  const rows = DATA.filter(d => {{
    if (f.servicio && d.servicio!==f.servicio) return false;
    if (f.mes && d.mes!==f.mes) return false;
    if (f.programa && d.programa!==f.programa) return false;
    if (f.resultado && d.resultado!==f.resultado) return false;
    if (f.estado && d.estado_txt!==f.estado) return false;
    if (f.q) {{
      const hay=(d.equipo+' '+d.serie+' '+d.inventario+' '+d.carpeta+' '+d.marca+' '+
                 d.modelo+' '+d.servicio+' '+d.unidad+' '+d.ubicacion+' '+d.procedencia).toLowerCase();
      if(!hay.includes(f.q)) return false;
    }}
    return true;
  }});
  const headers = COLUMNS.map(c => c[1]);
  const lines = [headers.join(';')];
  for (const d of rows) {{
    lines.push(COLUMNS.map(([k]) => {{
      let v = (k==='estado_txt')?d.estado_txt:d[k];
      v = (v==null)?'':String(v);
      if (v.includes(';')||v.includes('"')||v.includes('\\n')) v='"'+v.replace(/"/g,'""')+'"';
      return v;
    }}).join(';'));
  }}
  const blob = new Blob(['\\ufeff'+lines.join('\\r\\n')], {{type:'text/csv;charset=utf-8;'}});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'reporte_mp_2026.csv';
  a.click();
}});

applyFilters();
</script>
</body>
</html>"""


def main():
    import sys
    xlsm = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSM
    events, stats = build_events(xlsm)
    htmlout = render_html(events, stats, xlsm)
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(htmlout)
    print(f"OK: {stats['total_eventos']} eventos / {stats['total_equipos']} equipos")
    print(f"    Realizadas={stats['realizadas']}  Reprogramadas={stats['reprogramadas']}"
          f"  Pendientes(Ene-Jun)={stats['pendientes_eventos']}")
    print(f"    -> {OUTPUT_HTML}")


if __name__ == "__main__":
    main()
