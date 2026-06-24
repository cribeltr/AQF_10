# -*- coding: utf-8 -*-
"""
Generador del Reporte interactivo de Mantenciones Preventivas (MP) 2026
=======================================================================

Lee ``ProgramacionMP2026.xlsm`` y genera ``Reporte_MP_2026.html``: una
aplicación HTML autónoma (un solo archivo) que:

* **Importa el .xlsm en el navegador** (botón «Importar .xlsm») y se actualiza
  en vivo — usa la librería SheetJS embebida, sin conexión a internet.
* Vista única **«Eventos»**: una fila por combinación de equipo × mes con
  programación (P) y/o resultado (R). Clic en la fila abre la ficha del evento.
* UX/UI: columna «Equipo» fija al hacer scroll, filas alternadas, render por
  bloques (windowing) para fluidez, chips de filtros activos, estados vacíos,
  impresión/PDF, búsqueda con debounce y accesibilidad por teclado/lectores
  (modal con foco atrapado, encabezados con aria-sort, controles etiquetados).

Diseño clave
------------
Los datos del .xlsm (hoja ``Registro_MP-2026``) se incrustan como filas crudas
y TODA la transformación ocurre en JavaScript (función ``transform``). Así, los
datos por defecto y los datos importados pasan exactamente por la misma lógica,
garantizando resultados idénticos.

Se conservan los N° de Serie / Inventario tal cual (texto), respetando los
ceros a la izquierda.
"""

from __future__ import annotations

import datetime
import json
import os

import openpyxl

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_XLSM = os.path.join(BASE_DIR, "data", "ProgramacionMP2026.xlsm")
VENDOR_SHEETJS = os.path.join(BASE_DIR, "vendor", "xlsx.full.min.js")
OUTPUT_HTML = os.path.join(BASE_DIR, "Reporte_MP_2026.html")
SHEET = "Registro_MP-2026"
YEAR = 2026

# Rango: encabezado en la fila 7 de la planilla; se incrustan las columnas
# A..AR (índices 0..43, base 0) desde la fila 7 hacia abajo.
HEADER_ROW = 7          # 1-based (openpyxl)
LAST_COL = 44           # columnas A..AR (1..44)

# ---- Diccionarios de códigos (se inyectan al JS) ----------------------------
MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
         "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
PROGRAMA_LABEL = {
    "X": "MP Programada",
    "R": "MP Reprogramada",
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
    "Si": "MP Preventiva Realizada",
    "Si-RA": "MP de año anterior realizada",
    "FS": "Fuera de Servicio",
    "No": "No Realizada",
    "NU": "No Ubicable",
    "Baja": "Equipo Dado de Baja",
}
for _c, _d in CAUSAS.items():
    RESULTADO_LABEL[_c] = f"Reprogramada — {_d}"

# Estado del equipo derivado del código de resultado: r -> [claveColor, texto]
ESTADO = {
    "Si":    ["ok",     "Operativo (MP realizada)"],
    "Si-RA": ["ok",     "Operativo (MP año anterior realizada)"],
    "Baja":  ["baja",   "Dado de baja"],
    "FS":    ["fs",     "Fuera de servicio"],
    "NU":    ["nu",     "No ubicable"],
    "No":    ["no",     "MP no realizada"],
    "C1":    ["reprog", "En uso clínico (no liberable)"],
    "C2":    ["reprog", "En servicio técnico"],
    "C3":    ["reprog", "No operativo (espera repuestos)"],
    "C4":    ["reprog", "En préstamo a otra institución"],
    "C5":    ["reprog", "Operativo (sin HH SEC)"],
    "C6":    ["reprog", "Operativo (sin HH externo)"],
    "C7":    ["reprog", "Operativo (ausencia SEC)"],
    "C8":    ["reprog", "Operativo (contingencia)"],
}


def cell_text(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def read_raw(xlsm_path: str):
    """Devuelve filas crudas (lista de listas de texto), encabezado primero."""
    wb = openpyxl.load_workbook(xlsm_path, data_only=True, read_only=True)
    ws = wb[SHEET]
    rows = []
    for r_idx, row in enumerate(ws.iter_rows(min_row=HEADER_ROW, max_col=LAST_COL,
                                             values_only=True), start=HEADER_ROW):
        rows.append([cell_text(v) for v in row])
    # Recortar filas finales totalmente vacías.
    while rows and not any(c for c in rows[-1]):
        rows.pop()
    return rows


def js_safe(s: str) -> str:
    """Evita que cualquier '</script' rompa el documento."""
    return s.replace("</script", "<\\/script").replace("</SCRIPT", "<\\/SCRIPT")


def main():
    import sys
    xlsm = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_XLSM
    raw = read_raw(xlsm)
    sheetjs = open(VENDOR_SHEETJS, encoding="utf-8").read()
    gen_date = datetime.date(2026, 6, 24).strftime("%d-%m-%Y")

    data_json = js_safe(json.dumps(raw, ensure_ascii=False))
    template = HTML_TEMPLATE
    html_out = (
        template
        .replace("/*__SHEETJS__*/", js_safe(sheetjs))
        .replace("/*__DATA__*/", data_json)
        .replace("/*__MESES__*/", json.dumps(MESES, ensure_ascii=False))
        .replace("/*__PROG__*/", json.dumps(PROGRAMA_LABEL, ensure_ascii=False))
        .replace("/*__CAUSAS__*/", json.dumps(CAUSAS, ensure_ascii=False))
        .replace("/*__RESULT__*/", json.dumps(RESULTADO_LABEL, ensure_ascii=False))
        .replace("/*__ESTADO__*/", json.dumps(ESTADO, ensure_ascii=False))
        .replace("__YEAR__", str(YEAR))
        .replace("__GEN_DATE__", gen_date)
        .replace("__SRC__", os.path.basename(xlsm))
        .replace("__NROWS__", str(len(raw) - 1))
    )
    with open(OUTPUT_HTML, "w", encoding="utf-8") as f:
        f.write(html_out)
    print(f"OK: {len(raw)-1} filas de equipo incrustadas -> {OUTPUT_HTML}")
    print(f"    Tamaño: {os.path.getsize(OUTPUT_HTML)/1024:.0f} KB")


# --------------------------------------------------------------------------- #
# Plantilla HTML (se usa .replace con tokens; NO es f-string).
# --------------------------------------------------------------------------- #
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Reporte Mantenciones Preventivas __YEAR__ — H.H.H.A.</title>
<style>
  :root{
    --bg:#eef1f6; --panel:#fff; --ink:#1f2733; --muted:#647084; --line:#e1e6ee;
    --brand:#1565c0; --brand-dark:#0d3c75;
    --ok:#1b873f; --ok-bg:#e6f6ec; --reprog:#8a6d00; --reprog-bg:#fff7d6;
    --no:#c01525; --no-bg:#fde7e7; --nu:#9a4a00; --nu-bg:#ffe9d6;
    --baja:#344054; --noreg:#b54708; --noreg-bg:#ffeede;
    --prog:#155e9c; --prog-bg:#e6f0fb; --fs:#c01525; --fs-bg:#fde7e7;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--ink);font-size:13px;
    font-family:"Segoe UI",Roboto,system-ui,-apple-system,sans-serif}
  header.top{background:linear-gradient(135deg,var(--brand-dark),var(--brand));
    color:#fff;padding:16px 22px}
  header.top h1{margin:0 0 3px;font-size:19px}
  header.top p{margin:0;opacity:.9;font-size:12px}
  .wrap{padding:14px 20px 70px}
  .cards{display:flex;flex-wrap:wrap;gap:10px;margin:14px 0}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:10px;
    padding:10px 14px;min-width:120px;flex:1;cursor:pointer;transition:.12s;
    box-shadow:0 1px 2px rgba(16,24,40,.04)}
  .card:hover{transform:translateY(-1px);box-shadow:0 3px 10px rgba(16,24,40,.10);border-color:var(--brand)}
  .card .n{font-size:22px;font-weight:700;color:var(--brand-dark)}
  .card .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.03em}
  .toolbar{background:var(--panel);border:1px solid var(--line);border-radius:10px;
    padding:11px;display:flex;flex-wrap:wrap;gap:9px;align-items:center;margin:8px 0 12px}
  .toolbar input[type=search],.toolbar select{padding:7px 9px;border:1px solid var(--line);
    border-radius:7px;font-size:12.5px;background:#fff;color:var(--ink)}
  .toolbar input[type=search]{min-width:220px;flex:1}
  .toolbar .sp{margin-left:auto}
  .btn{padding:7px 12px;border:1px solid var(--line);background:#fff;border-radius:7px;
    cursor:pointer;font-size:12.5px;color:var(--ink)}
  .btn:hover{background:#eef3fb}
  .btn.primary{background:var(--brand);color:#fff;border-color:var(--brand)}
  .btn.primary:hover{background:var(--brand-dark)}
  .count{color:var(--muted);font-size:12px}
  .table-scroll{overflow:auto;max-height:72vh;border:1px solid var(--line);
    border-radius:10px;background:var(--panel)}
  table{border-collapse:collapse;width:100%;font-size:12.2px}
  thead th{position:sticky;top:0;background:#eef2f8;color:var(--brand-dark);text-align:left;
    padding:8px 9px;border-bottom:2px solid var(--line);white-space:nowrap;cursor:pointer;
    user-select:none;z-index:5}
  thead th:hover{background:#e2e9f4}
  thead th .ar{opacity:.4;font-size:9px}
  tbody td{padding:6px 9px;border-bottom:1px solid var(--line);white-space:nowrap}
  tbody tr:hover{background:#f6f9fd}
  td.mono{font-family:Consolas,"Courier New",monospace}
  tr.clk{cursor:pointer}
  .badge{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:600}
  .b-ok{color:var(--ok);background:var(--ok-bg)}
  .b-reprog{color:var(--reprog);background:var(--reprog-bg)}
  .b-no{color:var(--no);background:var(--no-bg)}
  .b-nu{color:var(--nu);background:var(--nu-bg)}
  .b-baja{color:#fff;background:var(--baja)}
  .b-noreg{color:var(--noreg);background:var(--noreg-bg)}
  .b-prog{color:var(--prog);background:var(--prog-bg)}
  .b-fs{color:var(--fs);background:var(--fs-bg)}
  .b-info{color:var(--prog);background:var(--prog-bg)}
  .pill{display:inline-block;min-width:22px;text-align:center;padding:1px 7px;border-radius:999px;
    font-weight:700;font-size:11px;cursor:pointer}
  .pill.zero{color:var(--muted);background:#eef0f4;cursor:default;font-weight:600}
  .cnt-ok{color:var(--ok);background:var(--ok-bg)}
  .cnt-reprog{color:var(--reprog);background:var(--reprog-bg)}
  .cnt-no{color:var(--no);background:var(--no-bg)}
  .cnt-nu{color:var(--nu);background:var(--nu-bg)}
  .cnt-baja{color:#fff;background:var(--baja)}
  .cnt-noreg{color:var(--noreg);background:var(--noreg-bg)}
  .muted{color:var(--muted)}
  /* Modal */
  .ov{position:fixed;inset:0;background:rgba(16,24,40,.55);display:none;align-items:center;
    justify-content:center;z-index:100;padding:20px}
  .ov.show{display:flex}
  .modal{background:#fff;border-radius:12px;max-width:880px;width:100%;max-height:86vh;
    overflow:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
  .modal h3{margin:0;padding:15px 20px;background:linear-gradient(135deg,var(--brand-dark),var(--brand));
    color:#fff;border-radius:12px 12px 0 0;font-size:16px;position:sticky;top:0;display:flex;justify-content:space-between}
  .modal h3 .x{cursor:pointer;opacity:.85;font-weight:400}
  .modal .body{padding:16px 20px}
  .kv{display:grid;grid-template-columns:auto 1fr;gap:4px 14px;font-size:12.5px;margin-bottom:12px}
  .kv b{color:var(--muted);font-weight:600}
  .legend{margin-top:20px;background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:2px 16px}
  .legend>summary{cursor:pointer;font-weight:600;padding:11px 0;color:var(--brand-dark);font-size:14px}
  .lg{display:flex;flex-wrap:wrap;gap:24px;padding:6px 0 14px}
  .lg>div{flex:1;min-width:260px}
  .lg h4{margin:6px 0;color:var(--brand-dark)}
  .lg table{width:100%;font-size:12px}
  .lg td{padding:4px 8px;border-bottom:1px solid var(--line);white-space:normal}
  .code{font-family:Consolas,monospace;font-weight:700;color:var(--brand-dark);
    background:#eef3fb;padding:1px 7px;border-radius:5px}
  .chip{display:inline-block;padding:3px 10px;border-radius:999px;font-size:11.5px;font-weight:600;
    cursor:pointer;border:1px solid transparent;margin:2px}
  .notes{font-size:12.2px;color:var(--muted);line-height:1.6}
  .notes b{color:var(--ink)}
  footer{text-align:center;color:var(--muted);font-size:11.5px;margin-top:22px}
  .toast{position:fixed;bottom:22px;left:50%;transform:translateX(-50%);background:#0d3c75;color:#fff;
    padding:11px 20px;border-radius:9px;box-shadow:0 8px 24px rgba(0,0,0,.25);display:none;z-index:200}
  .toast.show{display:block;animation:fade .3s}
  @keyframes fade{from{opacity:0;transform:translate(-50%,8px)}}
  /* ---- mejoras UX/UI ---- */
  tbody tr:nth-child(even) td{background:#f6f8fc}
  tbody tr:hover td{background:#e9f1fb}
  thead th:first-child{left:0;z-index:7}
  tbody td:first-child{position:sticky;left:0;background:var(--panel);z-index:2}
  tbody tr:nth-child(even) td:first-child{background:#f6f8fc}
  tbody tr:hover td:first-child{background:#e9f1fb}
  .firstcol{max-width:260px;overflow:hidden;text-overflow:ellipsis}
  .firstcol .rowbtn{max-width:100%;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;display:inline-block;vertical-align:bottom}
  .rowbtn{background:none;border:none;font:inherit;color:var(--brand-dark);font-weight:600;cursor:pointer;text-align:left;padding:0;width:100%}
  .lnk{background:none;border:none;color:var(--brand);cursor:pointer;font:inherit;padding:0;text-decoration:underline}
  :focus-visible{outline:2px solid var(--brand);outline-offset:1px;border-radius:3px}
  thead th:focus-visible{outline-offset:-2px}
  .active-chips{display:flex;flex-wrap:wrap;gap:6px;margin:-2px 0 12px;align-items:center;min-height:0}
  .active-chips:empty{display:none}
  .fchip{display:inline-flex;align-items:center;gap:6px;background:#e7eef9;color:var(--brand-dark);
    border:1px solid #cfe0f5;border-radius:999px;padding:3px 6px 3px 11px;font-size:12px}
  .fchip .fx{background:none;border:none;cursor:pointer;color:var(--brand-dark);font-size:13px;line-height:1;padding:0}
  td.empty{text-align:center;padding:44px}
  .empty-box{color:var(--muted);font-size:14px;line-height:2.2}
  #busy{position:fixed;inset:0;background:rgba(255,255,255,.62);display:none;align-items:center;justify-content:center;z-index:300}
  #busy.show{display:flex}
  .busy-card{background:#fff;border:1px solid var(--line);border-radius:10px;padding:16px 24px;
    box-shadow:0 12px 32px rgba(0,0,0,.18);font-weight:600;color:var(--brand-dark)}
  .print-only{display:none}
  @media print{
    header.top,.toolbar,.tabs,.cards,.legend,footer,.toast,.ov,#busy,.active-chips{display:none!important}
    .wrap{padding:0}
    .table-scroll{max-height:none;overflow:visible;border:none;border-radius:0}
    thead th{position:static}
    tbody td:first-child{position:static}
    .firstcol{max-width:none}
    .rowbtn,a{color:#000!important}
    body{font-size:10px;background:#fff}
    .print-only{display:block;margin:0 0 8px;font-weight:700;color:#000}
    tbody tr:hover td,tbody tr:nth-child(even) td{background:#fff}
  }
</style>
</head>
<body>
<header class="top">
  <h1>Programa de Mantención Preventiva de Equipos Médicos Críticos — __YEAR__</h1>
  <p>Hospital Hernán Henríquez Aravena · una fila por evento de mantención · importable y cliqueable</p>
</header>

<div class="wrap">
  <div class="cards" id="cards"></div>

  <div class="toolbar" role="search">
    <input type="search" id="q" aria-label="Buscar" placeholder="Buscar: equipo, serie, inventario, marca, servicio…">
    <select id="f_servicio" aria-label="Filtrar por servicio"><option value="">Servicio: todos</option></select>
    <select id="f_clasif" aria-label="Filtrar por clasificación"><option value="">Clasificación: todas</option></select>
    <select id="f_mes" aria-label="Filtrar por mes"><option value="">Mes: todos</option></select>
    <select id="f_resultado" aria-label="Filtrar por resultado"><option value="">Resultado: todos</option></select>
    <button class="btn" id="clear">Limpiar</button>
    <span class="sp"></span>
    <button class="btn primary" id="importBtn" title="Cargar un .xlsm actualizado">⭱ Importar .xlsm</button>
    <input type="file" id="file" accept=".xlsm,.xlsx,.xls" style="display:none">
    <button class="btn" id="csv" title="Exportar a CSV">Exportar CSV</button>
    <button class="btn" id="print" title="Imprimir o guardar como PDF">Imprimir</button>
    <span class="count" id="count" aria-live="polite"></span>
  </div>

  <div class="active-chips" id="activeFilters"></div>
  <div class="print-only" id="printHead"></div>

  <div class="table-scroll" id="scroller">
    <table id="tbl"><thead><tr id="head"></tr></thead><tbody id="body"></tbody></table>
  </div>

  <details class="legend">
    <summary>Leyenda de códigos, filtros rápidos y notas metodológicas</summary>
    <div style="padding:4px 0 10px">
      <b>Filtros rápidos por resultado (clic):</b><br>
      <span class="chip cnt-ok"   data-r="Si (Realizada)" role="button" tabindex="0">Sí · Realizada</span>
      <span class="chip cnt-reprog" data-r="C1–C8 (Reprogramada)" role="button" tabindex="0">C1–C8 · Reprogramada</span>
      <span class="chip cnt-no"   data-r="No (No realizada)" role="button" tabindex="0">No realizada</span>
      <span class="chip cnt-nu"   data-r="NU (No ubicable)" role="button" tabindex="0">No ubicable</span>
      <span class="chip cnt-baja" data-r="Baja" role="button" tabindex="0">Baja</span>
      <span class="chip cnt-noreg" data-r="No registrado" role="button" tabindex="0">No registrado</span>
    </div>
    <div class="lg">
      <div><h4>Programa (P)</h4><table id="lg-prog"></table>
           <h4 style="margin-top:12px">Resultado (R)</h4><table id="lg-res"></table></div>
      <div><h4>Causales de reprogramación</h4><table id="lg-cau"></table></div>
    </div>
    <div class="notes" id="notes"></div>
  </details>

  <footer id="foot"></footer>
</div>

<div class="ov" id="ov" role="dialog" aria-modal="true" aria-labelledby="m-titletext" aria-hidden="true">
  <div class="modal"><h3 id="m-title"><span id="m-titletext"></span><button class="x" id="m-x" aria-label="Cerrar">✕</button></h3>
  <div class="body" id="m-body"></div></div>
</div>
<div id="busy"><div class="busy-card">Procesando…</div></div>
<div class="toast" id="toast" role="status" aria-live="polite"></div>

<script>/*__SHEETJS__*/</script>
<script>
"use strict";
// ---- datos crudos por defecto (hoja Registro_MP-2026, encabezado primero) ----
const RAW = /*__DATA__*/;
const MESES = /*__MESES__*/;
const PROGRAMA_LABEL = /*__PROG__*/;
const CAUSAS = /*__CAUSAS__*/;
const RESULTADO_LABEL = /*__RESULT__*/;
const ESTADO = /*__ESTADO__*/;
const YEAR = __YEAR__;
const GEN_DATE = "__GEN_DATE__";
let SRC = "__SRC__";

const P_COLS=[19,21,23,25,27,29,31,33,35,37,39,41];
const R_COLS=[20,22,24,26,28,30,32,34,36,38,40,42];
const COL={id:1,carpeta:2,inventario:3,equipo:4,servicio:5,unidad:6,ubicacion:7,
  procedencia:8,marca:9,modelo:10,serie:11,anio:12,vida_util:13,clasif:14,enu_baja:15,frecuencia:17};
const REALIZADOS={"Si":1,"Si-RA":1};

function refMonth(){ const d=new Date(); const y=d.getFullYear();
  if(y>YEAR) return 12; if(y<YEAR) return 0; return d.getMonth()+1; }
const REF_MONTH = refMonth();

function txt(v){ if(v===null||v===undefined) return ""; return String(v).trim(); }
function esc(s){ return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;"); }

// ---- estado del equipo derivado del resultado del evento ----
function estadoInfo(p,r,mesIdx){
  if(r && ESTADO[r]) return {key:ESTADO[r][0],txt:ESTADO[r][1]};
  if(r) return {key:"info",txt:r};
  if(mesIdx<=REF_MONTH) return {key:"noreg",txt:"Pendiente (sin registro)"};
  return {key:"prog",txt:"Programada"};
}

// ---- TRANSFORM: filas crudas -> {events, equipos, stats} ----
function transform(rows){            // rows[0] = encabezado
  const events=[], equipos=[];
  for(let i=1;i<rows.length;i++){
    const row=rows[i]||[];
    const id=txt(row[COL.id]);
    if(!id) continue;
    const ident={};
    for(const k in COL) ident[k]=txt(row[COL[k]]);
    // placeholder "0" => vacío (sólo el literal "0"; preserva ceros a la izquierda como "0024")
    ["carpeta","inventario","serie","servicio","unidad","ubicacion","procedencia","enu_baja","frecuencia"]
      .forEach(k=>{ if(ident[k]==="0") ident[k]=""; });
    if(!ident.equipo && !ident.serie && !ident.inventario) continue;

    const meses=[];
    for(let m=0;m<12;m++) meses.push([txt(row[P_COLS[m]]), txt(row[R_COLS[m]])]);

    let pend=0, ultimaIdx=-1;
    const counts={si:0,reprog:0,no:0,nu:0,baja:0,noreg:0};
    for(let m=0;m<12;m++){
      const [p,r]=meses[m];
      if(r) ultimaIdx=m;
      if(REALIZADOS[r]) counts.si++;
      else if(CAUSAS[r]) counts.reprog++;
      else if(r==="No") counts.no++;
      else if(r==="NU") counts.nu++;
      else if(r==="Baja") counts.baja++;
      if(p && (m+1)<=REF_MONTH){
        if(!REALIZADOS[r] && r!=="Baja") pend++;
        if(!r) counts.noreg++;
      }
    }
    const ultima = ultimaIdx>=0 ? (MESES[ultimaIdx]+" "+YEAR) : "";

    for(let m=0;m<12;m++){
      const [p,r]=meses[m];
      if(!p && !r) continue;
      const es=estadoInfo(p,r,m+1);
      events.push(Object.assign({}, ident, {
        mes:MESES[m], mesIdx:m+1, programa:p, programaLabel:PROGRAMA_LABEL[p]||p,
        resultado:r, resultadoLabel:RESULTADO_LABEL[r]||r,
        fechaEjec: REALIZADOS[r] ? (MESES[m]+" "+YEAR) : "",
        estadoKey:es.key, estadoTxt:es.txt,
        pendientes:pend, ultimaActualizacion:ultima
      }));
    }
    equipos.push(Object.assign({}, ident, {counts, pendientes:pend, ultimaActualizacion:ultima}));
  }
  const stats={
    eventos:events.length,
    equipos:new Set(events.map(e=>e.id)).size,
    programadas:events.filter(e=>e.programa).length,
    realizadas:events.filter(e=>REALIZADOS[e.resultado]).length,
    reprogramadas:events.filter(e=>CAUSAS[e.resultado]).length,
    noreg:equipos.reduce((a,e)=>a+e.counts.noreg,0),
  };
  return {events,equipos,stats};
}

// ---- estado global ----
let DATA, sortKey=null, sortDir=1;
const $=id=>document.getElementById(id);
function setData(d){ DATA=d; }
setData(transform(RAW));

// ---- definición de columnas ----
const EVENT_COLS=[
  ["equipo","Nombre del Equipo"],["id","ID"],["carpeta","N° Carpeta"],["inventario","N° Inventario"],
  ["servicio","Servicio"],["unidad","Unidad"],["ubicacion","Ubicación"],["procedencia","Procedencia"],
  ["marca","Marca"],["modelo","Modelo"],["serie","N° de Serie"],["anio","Año Inst."],
  ["vida_util","Vida Útil Res."],["clasif","Clasificación"],["enu_baja","ENU / Baja"],
  ["mes","Mes"],["programa","Programa"],["resultado","Resultado"],["fechaEjec","Fecha de Ejecución"],
  ["estadoTxt","Estado del Equipo"],["pendientes","Cant. Pend."],["ultimaActualizacion","Última Actualización"]
];

// ---- filtros ----
function fillSelect(id, vals, prefix){
  const sel=$(id); const cur=sel.value;
  sel.innerHTML='<option value="">'+prefix+'</option>';
  vals.forEach(v=>{ const o=document.createElement("option"); o.value=v; o.textContent=v; sel.appendChild(o); });
  sel.value=cur;
}
function refreshFilters(){
  const servicios=[...new Set(DATA.equipos.map(e=>e.servicio).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
  const clasifs=[...new Set(DATA.equipos.map(e=>e.clasif).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
  fillSelect("f_servicio",servicios,"Servicio: todos");
  fillSelect("f_clasif",clasifs,"Clasificación: todas");
  fillSelect("f_mes",MESES,"Mes: todos");
  fillSelect("f_resultado",["Si (Realizada)","C1–C8 (Reprogramada)","No (No realizada)","NU (No ubicable)","Baja","No registrado"],"Resultado: todos");
}
function filters(){
  return {q:$("q").value.trim().toLowerCase(), servicio:$("f_servicio").value, clasif:$("f_clasif").value,
    mes:$("f_mes").value, resultado:$("f_resultado").value};
}
function matchEquipoText(e,q){
  if(!q) return true;
  return (e.equipo+" "+e.serie+" "+e.inventario+" "+e.carpeta+" "+e.marca+" "+e.modelo+" "+
          e.servicio+" "+e.unidad+" "+e.ubicacion+" "+e.procedencia).toLowerCase().includes(q);
}
function resultMatchesEvent(ev,resFilter){
  if(!resFilter) return true;
  if(resFilter.startsWith("Si")) return REALIZADOS[ev.resultado];
  if(resFilter.startsWith("C1")) return !!CAUSAS[ev.resultado];
  if(resFilter.startsWith("No (")) return ev.resultado==="No";
  if(resFilter.startsWith("NU")) return ev.resultado==="NU";
  if(resFilter.startsWith("Baja")) return ev.resultado==="Baja";
  if(resFilter.startsWith("No reg")) return !ev.resultado && ev.mesIdx<=REF_MONTH;
  return true;
}

// ---- filtrado ----
function filteredEvents(){
  const f=filters();
  return DATA.events.filter(ev=>{
    if(f.servicio && ev.servicio!==f.servicio) return false;
    if(f.clasif && ev.clasif!==f.clasif) return false;
    if(f.mes && ev.mes!==f.mes) return false;
    if(!resultMatchesEvent(ev,f.resultado)) return false;
    if(!matchEquipoText(ev,f.q)) return false;
    return true;
  });
}

// ---- ordenamiento ----
function sortRows(rows,getter){
  if(!sortKey) return rows;
  return rows.slice().sort((a,b)=>{
    let x=getter(a,sortKey), y=getter(b,sortKey);
    const nx=parseFloat(x), ny=parseFloat(y);
    if(!isNaN(nx)&&!isNaN(ny)&&String(x).trim()!==""&&String(y).trim()!==""){x=nx;y=ny;}
    if(x<y) return -sortDir; if(x>y) return sortDir; return 0;
  });
}

// ---- render encabezado (ordenable + accesible) ----
function renderHead(cols){
  let h="";
  cols.forEach(([k,l])=>{
    const sorted=(k===sortKey)?(sortDir>0?"ascending":"descending"):"none";
    const ar=(k===sortKey)?(sortDir>0?"▲":"▼"):"↕";
    h+='<th data-k="'+k+'" role="button" tabindex="0" aria-sort="'+sorted+'" title="Ordenar por '+esc(l)+'">'+esc(l)+' <span class="ar">'+ar+'</span></th>';
  });
  $("head").innerHTML=h;
}

// ---- ventana incremental (rinde por bloques para fluidez) ----
const PAGE=300;
let winRows=[], winShown=0, winRender=null;
function startWindow(rows, rowHtml){ winRows=rows; winShown=0; winRender=rowHtml; $("body").innerHTML=""; growWindow(); }
function growWindow(){
  if(!winRender || winShown>=winRows.length) return;
  const end=Math.min(winShown+PAGE, winRows.length); let h="";
  for(let i=winShown;i<end;i++) h+=winRender(winRows[i], i);
  $("body").insertAdjacentHTML("beforeend", h); winShown=end;
}
function showEmpty(cols){
  $("body").innerHTML='<tr><td class="empty" colspan="'+cols+'"><div class="empty-box">Sin resultados con los filtros actuales.<br>'+
    '<button class="btn" id="emptyClear">Limpiar filtros</button></div></td></tr>';
}

// ---- render eventos ----
function renderEventos(){
  renderHead(EVENT_COLS);
  const get=(e,k)=> k==="mes"? e.mesIdx : e[k];
  const rows=sortRows(filteredEvents(),get);
  window.__evRows=rows;
  $("count").textContent=rows.length+" de "+DATA.events.length+" eventos";
  if(!rows.length){ showEmpty(EVENT_COLS.length); return; }
  startWindow(rows, eventoRowHtml);
}
function eventoRowHtml(d, idx){
  let h='<tr class="clk" data-ev="'+idx+'">';
  EVENT_COLS.forEach(([k])=>{
    if(k==="equipo") h+='<td class="firstcol" title="'+esc(d.equipo)+'"><button class="rowbtn" aria-label="Ver detalle del evento">'+(esc(d.equipo)||"—")+'</button></td>';
    else if(k==="estadoTxt") h+='<td><span class="badge b-'+d.estadoKey+'">'+esc(d.estadoTxt)+'</span></td>';
    else if(k==="pendientes"){ const c=d.pendientes>0?"cnt-noreg":""; h+='<td><span class="pill '+(c||"zero")+'">'+esc(d.pendientes)+'</span></td>'; }
    else if(k==="serie"||k==="inventario") h+='<td class="mono">'+esc(d[k])+'</td>';
    else if(k==="programa") h+='<td title="'+esc(d.programaLabel)+'">'+(esc(d.programa)||"—")+'</td>';
    else if(k==="resultado") h+='<td title="'+esc(d.resultadoLabel)+'">'+(esc(d.resultado)||'<span class=muted>—</span>')+'</td>';
    else h+='<td>'+(esc(d[k])||'<span class=muted>—</span>')+'</td>';
  });
  return h+"</tr>";
}

function render(){ renderActiveChips(); renderEventos(); }
function sortBy(k){ if(sortKey===k) sortDir*=-1; else {sortKey=k;sortDir=1;} render(); }
function clearFilters(){ $("q").value=""; ["f_servicio","f_clasif","f_mes","f_resultado"].forEach(id=>$(id).value=""); sortKey=null; render(); }

// ---- chips de filtros activos (removibles) ----
function renderActiveChips(){
  const f=filters(); const items=[];
  if($("q").value.trim()) items.push(["q","Búsqueda: "+$("q").value.trim()]);
  if(f.servicio) items.push(["f_servicio","Servicio: "+f.servicio]);
  if(f.clasif) items.push(["f_clasif","Clasificación: "+f.clasif]);
  if(f.mes) items.push(["f_mes","Mes: "+f.mes]);
  if(f.resultado) items.push(["f_resultado","Resultado: "+f.resultado]);
  $("activeFilters").innerHTML = items.length
    ? items.map(([id,lbl])=>'<span class="fchip">'+esc(lbl)+'<button class="fx" data-clear="'+id+'" aria-label="Quitar filtro">✕</button></span>').join("")
      +'<button class="lnk" id="chipClearAll">Limpiar todo</button>'
    : "";
}

// ---- tarjetas resumen (cliqueables) ----
function renderCards(){
  const s=DATA.stats;
  const cards=[
    ["eventos",s.eventos,"Eventos",null],
    ["equipos",s.equipos,"Equipos",null],
    ["realizadas",s.realizadas,"Realizadas (Sí)","Si (Realizada)"],
    ["reprogramadas",s.reprogramadas,"Reprogramadas","C1–C8 (Reprogramada)"],
    ["noreg",s.noreg,"No registradas","No registrado"],
    ["programadas",s.programadas,"Programadas",null],
  ];
  $("cards").innerHTML=cards.map(c=>
    '<div class="card" data-rf="'+(c[3]||"")+'"><div class="n">'+c[1]+'</div><div class="l">'+c[2]+'</div></div>').join("");
}

// ---- modales (con manejo de foco) ----
let lastFocus=null;
function openModal(title, bodyHtml){
  lastFocus=document.activeElement;
  $("m-titletext").textContent=title;
  $("m-body").innerHTML=bodyHtml;
  const ov=$("ov"); ov.classList.add("show"); ov.setAttribute("aria-hidden","false");
  const f=ov.querySelector("textarea,select,button,input,[tabindex]"); if(f&&f.focus) f.focus();
}
function closeModal(){ const ov=$("ov"); ov.classList.remove("show"); ov.setAttribute("aria-hidden","true");
  if(lastFocus&&lastFocus.focus){ try{lastFocus.focus();}catch(e){} } }
function kv(pairs){ return '<div class="kv">'+pairs.map(p=>'<b>'+esc(p[0])+'</b><span>'+(esc(p[1])||'—')+'</span>').join("")+'</div>'; }

function modalEvento(ev){
  openModal((ev.equipo||"Evento")+" — "+ev.mes+" "+YEAR, kv([
    ["ID",ev.id],["N° Carpeta",ev.carpeta],["N° Inventario",ev.inventario],["Equipo",ev.equipo],
    ["Servicio",ev.servicio],["Unidad",ev.unidad],["Ubicación",ev.ubicacion],["Procedencia",ev.procedencia],
    ["Marca",ev.marca],["Modelo",ev.modelo],["N° de Serie",ev.serie],["Año Instalación",ev.anio],
    ["Vida Útil Residual",ev.vida_util],["Clasificación",ev.clasif],["ENU / Baja",ev.enu_baja],
    ["Mes",ev.mes+" "+YEAR],["Programa",ev.programa?(ev.programa+" — "+ev.programaLabel):"—"],
    ["Resultado",(ev.resultado||"—")+(ev.resultado?(" — "+ev.resultadoLabel):"")],
    ["Fecha de Ejecución",ev.fechaEjec],["Estado del Equipo",ev.estadoTxt],
    ["Cant. Pendientes",ev.pendientes],["Última Actualización",ev.ultimaActualizacion]]));
}

// ---- eventos de UI ----
$("head").addEventListener("click",ev=>{ const th=ev.target.closest("th"); if(th) sortBy(th.dataset.k); });
$("head").addEventListener("keydown",ev=>{ if(ev.key==="Enter"||ev.key===" "){ const th=ev.target.closest("th"); if(th){ ev.preventDefault(); sortBy(th.dataset.k); } } });
$("body").addEventListener("click",ev=>{
  if(ev.target.id==="emptyClear"){ clearFilters(); return; }
  const evRow=ev.target.closest("tr[data-ev]");
  if(evRow){ modalEvento(window.__evRows[+evRow.dataset.ev]); }
});
$("scroller").addEventListener("scroll",()=>{ const s=$("scroller");
  if(s.scrollTop+s.clientHeight >= s.scrollHeight-300) growWindow(); });
$("cards").addEventListener("click",ev=>{ const c=ev.target.closest(".card"); if(!c) return;
  if(c.dataset.rf){ $("f_resultado").value=c.dataset.rf; render(); } });
$("activeFilters").addEventListener("click",ev=>{
  const x=ev.target.closest("[data-clear]"); if(x){ $(x.dataset.clear).value=""; render(); return; }
  if(ev.target.id==="chipClearAll") clearFilters();
});
let qT; $("q").addEventListener("input",()=>{ clearTimeout(qT); qT=setTimeout(render,180); });
["f_servicio","f_clasif","f_mes","f_resultado"].forEach(id=>$(id).addEventListener("change",render));
$("clear").addEventListener("click",clearFilters);
$("m-x").addEventListener("click",closeModal);
$("ov").addEventListener("click",ev=>{ if(ev.target===$("ov")) closeModal(); });
document.addEventListener("keydown",ev=>{ if(ev.key==="Escape" && $("ov").classList.contains("show")) closeModal(); });
$("ov").addEventListener("keydown",ev=>{                 // trampa de foco en el modal
  if(ev.key!=="Tab") return;
  const f=[...$("ov").querySelectorAll('button,textarea,select,input,[tabindex]:not([tabindex="-1"])')].filter(el=>!el.disabled);
  if(!f.length) return; const first=f[0], last=f[f.length-1];
  if(ev.shiftKey && document.activeElement===first){ ev.preventDefault(); last.focus(); }
  else if(!ev.shiftKey && document.activeElement===last){ ev.preventDefault(); first.focus(); }
});

// filtros rápidos (chips de leyenda): filtran por resultado -> vista Eventos
function chipApply(ch){ $("f_resultado").value=ch.dataset.r||""; render(); window.scrollTo({top:0,behavior:"smooth"}); }
document.querySelectorAll(".chip").forEach(ch=>{
  ch.addEventListener("click",()=>chipApply(ch));
  ch.addEventListener("keydown",e=>{ if(e.key==="Enter"||e.key===" "){ e.preventDefault(); chipApply(ch); } });
});

// ---- importar .xlsm (con indicador de proceso) ----
function setBusy(on,msg){ const b=$("busy"); if(msg) b.firstElementChild.textContent=msg; b.classList.toggle("show",!!on); }
$("importBtn").addEventListener("click",()=>$("file").click());
$("file").addEventListener("change",ev=>{
  const file=ev.target.files[0]; if(!file) return;
  setBusy(true,"Procesando archivo…");
  const rd=new FileReader();
  rd.onload=e=>{ setTimeout(()=>{ try{
      const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:false});
      if(wb.SheetNames.indexOf("Registro_MP-2026")<0){ toast("No se encontró la hoja «Registro_MP-2026»."); return; }
      const ws=wb.Sheets["Registro_MP-2026"];
      const aoa=XLSX.utils.sheet_to_json(ws,{header:1,raw:false,defval:""});
      setData(transform(aoa.slice(6))); SRC=file.name;
      refreshFilters(); renderCards(); render(); renderNotes();
      toast("Datos actualizados desde «"+file.name+"»: "+DATA.stats.eventos+" eventos / "+DATA.stats.equipos+" equipos.");
    }catch(err){ toast("Error al leer el archivo: "+err.message); } finally{ setBusy(false); } }, 30);
  };
  rd.onerror=()=>{ setBusy(false); toast("No se pudo leer el archivo seleccionado."); };
  rd.readAsArrayBuffer(file); ev.target.value="";
});

// ---- exportar CSV (eventos filtrados) ----
$("csv").addEventListener("click",()=>{
  const lines=[EVENT_COLS.map(c=>c[1]).join(";")];
  filteredEvents().forEach(d=>{ lines.push(EVENT_COLS.map(c=>csv(c[0]==="estadoTxt"?d.estadoTxt:d[c[0]])).join(";")); });
  const blob=new Blob(["﻿"+lines.join("\r\n")],{type:"text/csv;charset=utf-8;"});
  const a=document.createElement("a"); a.href=URL.createObjectURL(blob); a.download="reporte_mp_"+YEAR+"_eventos.csv"; a.click();
});
function csv(v){ v=(v==null)?"":String(v); return (/[;"\n]/.test(v))?'"'+v.replace(/"/g,'""')+'"':v; }

// ---- imprimir / PDF (rinde todas las filas antes de imprimir) ----
$("print").addEventListener("click",()=>{
  while(winShown<winRows.length) growWindow();
  $("printHead").textContent="Reporte MP "+YEAR+" · Eventos · "+new Date().toLocaleString("es-CL");
  window.print();
});

// ---- toast ----
let toastT;
function toast(msg){ const t=$("toast"); t.textContent=msg; t.classList.add("show");
  clearTimeout(toastT); toastT=setTimeout(()=>t.classList.remove("show"),4200); }

// ---- leyenda + notas ----
function renderNotes(){
  $("lg-prog").innerHTML=Object.entries(PROGRAMA_LABEL).map(([c,d])=>"<tr><td><span class='code'>"+c+"</span></td><td>"+esc(d)+"</td></tr>").join("");
  $("lg-res").innerHTML=[["Si","Mantención Preventiva Realizada"],["C1–C8","Mantención Preventiva Reprogramada (causales)"],
    ["Si-RA","Mantención de Año Anterior Realizada"],["FS","Fuera de Servicio"],["No","No Realizada"],
    ["NU","No Ubicable"],["Baja","Equipo Dado de Baja"]].map(([c,d])=>"<tr><td><span class='code'>"+c+"</span></td><td>"+esc(d)+"</td></tr>").join("");
  $("lg-cau").innerHTML=Object.entries(CAUSAS).map(([c,d])=>"<tr><td><span class='code'>"+c+"</span></td><td>"+esc(d)+"</td></tr>").join("");
  const mesRef = REF_MONTH>=1&&REF_MONTH<=12 ? MESES[REF_MONTH-1].toLowerCase() : "—";
  $("notes").innerHTML=
    "<p><b>Fuente:</b> hoja <code>Registro_MP-2026</code> de <code>"+esc(SRC)+"</code> (encabezado fila 7, columnas B–AQ, ignorando Q «Observación» y S «Responsable MP»).</p>"+
    "<p><b>Una fila por evento</b>: cada combinación de equipo × mes con programación (P) y/o resultado (R). <b>Clic</b> en una fila para ver el detalle del evento.</p>"+
    "<p><b>No registrado</b> (resultado vacío) = meses transcurridos (enero a "+mesRef+" "+YEAR+") con MP programada y sin resultado registrado; se refleja en «Pendientes» y en el estado «Pendiente (sin registro)».</p>"+
    "<p><b>N° de Serie / Inventario:</b> se conservan tal cual, respetando ceros a la izquierda.</p>"+
    "<p><b>Fecha de Ejecución:</b> el registro es a nivel de mes; se muestra el mes cuando el resultado es Si/Si-RA.</p>"+
    "<p class='muted'>Reporte generado el "+GEN_DATE+". Cálculos de meses transcurridos según la fecha del navegador (mes de referencia actual: "+mesRef+" "+YEAR+"). Puede importar un .xlsm actualizado con el botón «Importar .xlsm».</p>";
  $("foot").textContent="Reporte autónomo · "+DATA.stats.eventos+" eventos · "+DATA.stats.equipos+" equipos · generado "+GEN_DATE;
}

// ---- init ----
refreshFilters(); renderCards(); renderNotes(); render();
</script>
</body>
</html>"""


if __name__ == "__main__":
    main()
