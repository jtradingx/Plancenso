#!/usr/bin/env python3
"""
generar_dashboard.py
Descarga el Excel desde OneDrive y genera index.html con el dashboard del Censo 2026.
"""

import os
import sys
import json
import requests
import pandas as pd
from datetime import datetime
from io import BytesIO

# ── CONFIGURACIÓN ──────────────────────────────────────────────────────────────
# URL de descarga directa del Excel en OneDrive.
# Instrucciones para obtenerla:
#   1. Abre el archivo en OneDrive/SharePoint
#   2. Haz clic en los tres puntos (...) → "Detalles"
#   3. Copia la "Ruta de acceso directa" o usa "Compartir → Copiar vínculo"
#      y reemplaza "?web=1" por "?download=1" al final de la URL
# Luego guarda esa URL como secret en GitHub:
#   Settings → Secrets → Actions → New secret → ONEDRIVE_URL
ONEDRIVE_URL = os.environ.get("ONEDRIVE_URL", "")

# Si no hay URL de OneDrive, busca el Excel localmente (para pruebas)
LOCAL_EXCEL = "Censo_2026_Ejecutivo.xlsx"

# ── CARGA DEL EXCEL ────────────────────────────────────────────────────────────
def cargar_excel():
    if ONEDRIVE_URL:
        print(f"Descargando Excel desde OneDrive...")
        r = requests.get(ONEDRIVE_URL, timeout=60)
        r.raise_for_status()
        return BytesIO(r.content)
    elif os.path.exists(LOCAL_EXCEL):
        print(f"Usando archivo local: {LOCAL_EXCEL}")
        return LOCAL_EXCEL
    else:
        print("ERROR: No se encontró ONEDRIVE_URL ni archivo local.")
        sys.exit(1)

# ── PARSEO DEL EXCEL ───────────────────────────────────────────────────────────
def parsear_excel(fuente):
    meta_df = pd.read_excel(fuente, sheet_name="Tareas de proyecto", header=None)
    proyecto_nombre = str(meta_df.iloc[0, 1])
    proyecto_inicio = str(meta_df.iloc[2, 1])[:10]
    proyecto_fin    = str(meta_df.iloc[3, 1])[:10]
    proyecto_pct    = float(meta_df.iloc[5, 1])
    exportado       = str(meta_df.iloc[6, 1])[:16]

    df = pd.read_excel(fuente, sheet_name="Tareas de proyecto", header=8)

    def get_depth(scheme):
        s = str(scheme)
        if s in ("nan", "None", ""):
            return 0
        return len(s.split("."))

    df["depth"] = df["Número de esquema"].apply(get_depth)

    groups = []
    current_group = None

    for _, row in df.iterrows():
        depth  = row["depth"]
        scheme = str(row["Número de esquema"])
        name   = str(row["Nombre"]).strip() if pd.notna(row["Nombre"]) else ""
        pct    = float(row["% completado"]) if pd.notna(row["% completado"]) else 0.0
        asig   = str(row["Asignado a"]) if pd.notna(row["Asignado a"]) else ""
        prio   = str(row["Prioridad"]) if pd.notna(row["Prioridad"]) else "Media"
        inicio = row["Inicio"].strftime("%Y-%m-%d") if pd.notna(row["Inicio"]) else None
        fin    = row["Finalización"].strftime("%Y-%m-%d") if pd.notna(row["Finalización"]) else None

        if depth == 1:
            current_group = {
                "id": "g" + scheme.replace(".", "_"),
                "scheme": scheme,
                "label": name,
                "progress": round(pct * 100),
                "responsible": asig,
                "start": inicio,
                "end": fin,
                "tasks": [],
            }
            groups.append(current_group)
        elif depth >= 2 and current_group is not None:
            current_group["tasks"].append({
                "scheme": scheme,
                "depth": depth,
                "name": name,
                "progress": round(pct * 100),
                "responsible": asig,
                "prioridad": prio,
                "start": inicio,
                "end": fin,
            })

    for g in groups:
        sc = {t["scheme"] for t in g["tasks"]}
        leaves = [t for t in g["tasks"]
                  if not any(o.startswith(t["scheme"] + ".") for o in sc)]
        g["tasksTotal"]      = len(leaves)
        g["tasksDone"]       = sum(1 for t in leaves if t["progress"] == 100)
        g["tasksInProgress"] = sum(1 for t in leaves if 0 < t["progress"] < 100)

    return {
        "groups": groups,
        "proyecto_nombre": proyecto_nombre,
        "proyecto_inicio": proyecto_inicio,
        "proyecto_fin": proyecto_fin,
        "proyecto_pct": round(proyecto_pct * 100),
        "exportado": exportado,
    }

# ── GENERACIÓN HTML ────────────────────────────────────────────────────────────
HTML_TEMPLATE = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Dashboard — Censo 2026 Ejecutivo</title>
<link href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=IBM+Plex+Sans:wght@400;500;600;700&display=swap" rel="stylesheet">
<style>
:root {
  --bg:#f2f4f8; --s1:#fff; --s3:#eef0f6;
  --border:#dde1ed; --bd2:#c5cce0;
  --text:#1b1f30; --t2:#4d5475; --t3:#9aa0bc;
  --blue:#2563eb; --blue2:#1d4ed8;
  --green:#059669; --yellow:#b45309; --red:#dc2626; --purple:#7c3aed;
  --font:'IBM Plex Sans',sans-serif; --mono:'IBM Plex Mono',monospace;
}
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0}
body{background:var(--bg);color:var(--text);font-family:var(--font);font-size:13px;min-height:100vh}
header{background:var(--s1);border-bottom:1px solid var(--border);height:52px;padding:0 24px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 1px 4px rgba(0,0,0,.06);position:sticky;top:0;z-index:10}
.hl{display:flex;align-items:center;gap:12px}
.logo{width:28px;height:28px;border-radius:5px;background:var(--blue2);display:flex;align-items:center;justify-content:center;font-family:var(--mono);font-size:11px;font-weight:600;color:#fff}
.htitle{font-size:13px;font-weight:600}
.hsep{width:1px;height:16px;background:var(--bd2)}
.htag{font-family:var(--mono);font-size:10px;color:var(--t3);letter-spacing:.06em;text-transform:uppercase}
.hdate{font-family:var(--mono);font-size:10px;color:var(--t3)}
.hdot{width:7px;height:7px;border-radius:50%;background:var(--green);box-shadow:0 0 6px var(--green)}
.hexp{font-family:var(--mono);font-size:9px;color:var(--t3);background:var(--s3);padding:2px 7px;border-radius:3px;border:1px solid var(--border)}
.wrap{padding:20px 24px;max-width:1420px;margin:0 auto}
.slbl{font-family:var(--mono);font-size:9px;font-weight:600;letter-spacing:.12em;text-transform:uppercase;color:var(--t3);display:flex;align-items:center;gap:8px;margin-bottom:10px}
.slbl::after{content:'';flex:1;height:1px;background:var(--border)}
.krow{display:grid;grid-template-columns:repeat(4,1fr);gap:10px;margin-bottom:14px}
.kpi{background:var(--s1);border:1px solid var(--border);border-radius:6px;padding:14px 16px;position:relative;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.05)}
.kbar{position:absolute;left:0;top:0;bottom:0;width:3px;border-radius:6px 0 0 6px}
.klbl{font-family:var(--mono);font-size:9px;letter-spacing:.1em;text-transform:uppercase;color:var(--t3);margin-bottom:8px}
.knum{font-family:var(--mono);font-size:26px;font-weight:600;letter-spacing:-.03em;line-height:1;margin-bottom:6px}
.ksub{font-size:11px;color:var(--t2);display:flex;align-items:center;gap:5px}
.dot{width:6px;height:6px;border-radius:50%;flex-shrink:0}
.tline-wrap{background:var(--s1);border:1px solid var(--border);border-radius:6px;padding:14px 16px;margin-bottom:14px;box-shadow:0 1px 3px rgba(0,0,0,.05)}
.tline-track{position:relative;height:8px;background:var(--s3);border-radius:99px;overflow:hidden;margin-top:8px}
.tline-fill{position:absolute;left:0;top:0;height:100%;background:linear-gradient(90deg,var(--blue2),var(--blue));border-radius:99px}
.tline-marker{position:absolute;top:-4px;width:2px;height:16px;background:var(--red);border-radius:1px}
.tline-now{position:absolute;top:-10px;font-family:var(--mono);font-size:8px;color:var(--red);transform:translateX(-50%);white-space:nowrap}
.tline-meta{display:flex;justify-content:space-between;margin-top:6px;font-family:var(--mono);font-size:9px;color:var(--t3)}
.fbar{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:12px}
.flbl{font-family:var(--mono);font-size:9px;letter-spacing:.08em;text-transform:uppercase;color:var(--t3)}
.fb{font-family:var(--mono);font-size:10px;padding:4px 10px;border-radius:4px;border:1px solid var(--border);background:var(--s1);color:var(--t2);cursor:pointer;transition:all .12s;white-space:nowrap}
.fb:hover{border-color:var(--bd2);color:var(--text)}
.fb.on{background:var(--blue2);border-color:var(--blue);color:#fff}
.pf-group{display:flex;gap:4px}
.pfb{font-family:var(--mono);font-size:10px;padding:4px 8px;border-radius:4px;border:1px solid var(--border);background:var(--s1);color:var(--t2);cursor:pointer;transition:all .12s}
.pfb.on{border-color:var(--blue2);color:var(--blue2);background:rgba(37,99,235,.08)}
.view-toggle{display:flex;align-items:center;gap:2px;background:var(--s3);border:1px solid var(--border);border-radius:6px;padding:3px;margin-bottom:14px;width:fit-content}
.vbtn{font-family:var(--mono);font-size:10px;font-weight:500;letter-spacing:.04em;padding:5px 16px;border-radius:4px;border:none;background:none;color:var(--t2);cursor:pointer;transition:all .15s}
.vbtn.on{background:var(--s1);color:var(--blue);box-shadow:0 1px 3px rgba(0,0,0,.1);font-weight:600}
.layout{display:grid;grid-template-columns:1fr 400px;gap:16px;align-items:start}
.card{background:var(--s1);border:1px solid var(--border);border-radius:6px;padding:14px 16px;cursor:pointer;transition:border-color .15s,box-shadow .15s;position:relative;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.05);margin-bottom:8px}
.card:hover{border-color:var(--bd2);box-shadow:0 3px 12px rgba(0,0,0,.08)}
.card.active{border-color:var(--blue);box-shadow:0 0 0 2px rgba(37,99,235,.15)}
.card-top{position:absolute;top:0;left:0;right:0;height:3px}
.card-hdr{display:flex;align-items:flex-start;justify-content:space-between;gap:8px;margin-bottom:10px}
.cnum{font-family:var(--mono);font-size:9px;color:var(--t3);background:var(--s3);padding:2px 5px;border-radius:3px;border:1px solid var(--border);flex-shrink:0;margin-top:1px}
.cname{font-size:13px;font-weight:600;line-height:1.3}
.cdates{font-family:var(--mono);font-size:9px;color:var(--t3);margin-top:2px}
.ptrack{height:3px;background:var(--s3);border-radius:99px;overflow:hidden;margin-bottom:10px}
.pfill{height:100%;border-radius:99px}
.stats{display:grid;grid-template-columns:repeat(3,1fr);gap:4px}
.sb{background:var(--s3);border-radius:4px;padding:6px 8px}
.sv{font-family:var(--mono);font-size:13px;font-weight:600;line-height:1;margin-bottom:2px}
.sl{font-family:var(--mono);font-size:8px;letter-spacing:.06em;text-transform:uppercase;color:var(--t3)}
.cfoot{display:flex;align-items:center;justify-content:space-between;margin-top:10px;padding-top:10px;border-top:1px solid var(--border);font-family:var(--mono);font-size:9px}
.ccta{color:var(--blue)}
.detail{background:var(--s1);border:1px solid var(--border);border-radius:6px;box-shadow:0 1px 3px rgba(0,0,0,.05);position:sticky;top:68px;overflow:hidden;max-height:calc(100vh - 80px);display:flex;flex-direction:column}
.detail-empty{padding:40px 20px;text-align:center;font-family:var(--mono);font-size:11px;color:var(--t3);line-height:1.8}
.detail-head{padding:14px 16px;border-bottom:1px solid var(--border);flex-shrink:0}
.dh-title{font-size:14px;font-weight:700;letter-spacing:-.02em;line-height:1.3}
.dh-sub{font-family:var(--mono);font-size:9px;color:var(--t2);margin-top:3px}
.dh-dates{font-family:var(--mono);font-size:9px;color:var(--t3);margin-top:2px}
.dkpis{display:grid;grid-template-columns:repeat(3,1fr);gap:6px;padding:10px 16px;border-bottom:1px solid var(--border);flex-shrink:0}
.dk{background:var(--s3);border-radius:5px;padding:8px 10px;text-align:center}
.dk-val{font-family:var(--mono);font-size:17px;font-weight:600;line-height:1;margin-bottom:2px}
.dk-lbl{font-family:var(--mono);font-size:8px;letter-spacing:.08em;text-transform:uppercase;color:var(--t3)}
.d-progress-wrap{padding:8px 16px;border-bottom:1px solid var(--border);flex-shrink:0}
.d-ptrack{height:5px;background:var(--s3);border-radius:99px;overflow:hidden}
.tasklist{padding:10px 16px 16px;overflow-y:auto;flex:1}
.tasklist::-webkit-scrollbar{width:3px}
.tasklist::-webkit-scrollbar-thumb{background:var(--bd2);border-radius:2px}
.trow{display:grid;grid-template-columns:1fr 80px 36px;align-items:center;gap:8px;padding:7px 10px;background:var(--s3);border-radius:5px;margin-bottom:4px}
.trow.depth-3{padding-left:20px;background:rgba(238,240,246,.6)}
.trow.depth-4{padding-left:32px;background:rgba(238,240,246,.35)}
.tname{font-size:12px;font-weight:500;line-height:1.3}
.tsub{font-family:var(--mono);font-size:9px;color:var(--t3);margin-top:1px}
.tsub-date{font-family:var(--mono);font-size:8px;color:var(--t3);margin-top:1px}
.tbar-wrap{display:flex;flex-direction:column;gap:3px}
.tbar{height:3px;background:var(--border);border-radius:99px;overflow:hidden}
.tbar-fill{height:100%;border-radius:99px}
.tpct{font-family:var(--mono);font-size:12px;font-weight:600;text-align:right}
.tno-tasks{font-family:var(--mono);font-size:10px;color:var(--t3);padding:20px;text-align:center}
.task-view{display:none}
.task-view.active{display:block}
.area-view.hidden{display:none}
.tv-filters{display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:12px}
.ttable{width:100%;border-collapse:collapse}
.ttable-head th{background:var(--s3);border:1px solid var(--border);padding:8px 12px;font-family:var(--mono);font-size:9px;font-weight:600;letter-spacing:.08em;text-transform:uppercase;color:var(--t3);text-align:left;white-space:nowrap;position:sticky;top:52px;z-index:4}
.ttable tbody tr{border-bottom:1px solid var(--border);transition:background .1s}
.ttable tbody tr:hover{background:var(--s3)}
.ttable td{padding:8px 12px;vertical-align:middle}
.tt-scheme{font-family:var(--mono);font-size:9px;color:var(--t3);background:var(--s3);padding:2px 6px;border-radius:3px;border:1px solid var(--border);white-space:nowrap}
.tt-name{font-size:12px;font-weight:500;line-height:1.3}
.tt-group{font-family:var(--mono);font-size:9px;color:var(--t2);background:var(--s3);padding:2px 7px;border-radius:3px;border:1px solid var(--border);white-space:nowrap;max-width:140px;overflow:hidden;text-overflow:ellipsis;display:inline-block}
.tt-bar-wrap{display:flex;align-items:center;gap:8px;min-width:100px}
.tt-bar{flex:1;height:4px;background:var(--border);border-radius:99px;overflow:hidden}
.tt-fill{height:100%;border-radius:99px}
.tt-pct{font-family:var(--mono);font-size:11px;font-weight:600;width:34px;text-align:right;flex-shrink:0}
.tt-resp{font-family:var(--mono);font-size:9px;color:var(--t2);white-space:nowrap;max-width:130px;overflow:hidden;text-overflow:ellipsis;display:block}
.tgroup-row td{background:var(--bg);padding:6px 12px;font-family:var(--mono);font-size:9px;font-weight:600;letter-spacing:.1em;text-transform:uppercase;color:var(--t3);border-bottom:1px solid var(--border)}
@media(max-width:1000px){.layout{grid-template-columns:1fr}.detail{position:static;max-height:none}}
@media(max-width:700px){.krow{grid-template-columns:1fr 1fr}}
</style>
</head>
<body>
<header>
  <div class="hl">
    <div class="logo">C26</div>
    <div class="hsep"></div>
    <div class="htitle">Dashboard de Planificación</div>
    <div class="hsep"></div>
    <div class="htag">CENSO 2026 — EJECUTIVO</div>
  </div>
  <div class="hl">
    <span class="hexp" id="hExp"></span>
    <span class="hdate" id="hDate"></span>
    <div class="hdot"></div>
  </div>
</header>
<div class="wrap">
  <div class="slbl">Indicadores generales</div>
  <div class="krow">
    <div class="kpi"><div class="kbar" style="background:var(--blue)"></div><div class="klbl">Avance General</div><div class="knum" style="color:var(--blue)" id="kAvance">—</div><div class="ksub"><span class="dot" style="background:var(--blue)"></span><span id="kAvanceSub"></span></div></div>
    <div class="kpi"><div class="kbar" style="background:var(--green)"></div><div class="klbl">Proyectos / Tareas hoja</div><div class="knum" style="color:var(--green)" id="kTareas">—</div><div class="ksub"><span class="dot" style="background:var(--yellow)"></span><span id="kTareasSub"></span></div></div>
    <div class="kpi"><div class="kbar" style="background:var(--yellow)"></div><div class="klbl">En progreso</div><div class="knum" style="color:var(--yellow)" id="kProg">—</div><div class="ksub"><span class="dot" style="background:var(--t3)"></span><span id="kProgSub"></span></div></div>
    <div class="kpi"><div class="kbar" style="background:var(--purple)"></div><div class="klbl">Plazo del Proyecto</div><div class="knum" style="color:var(--purple)" id="kPlazo">—</div><div class="ksub"><span class="dot" style="background:var(--purple)"></span><span id="kPlazoSub"></span></div></div>
  </div>
  <div class="tline-wrap">
    <div style="display:flex;justify-content:space-between;align-items:center">
      <span class="klbl" style="margin-bottom:0">Línea de tiempo del proyecto</span>
      <span style="font-family:var(--mono);font-size:9px;color:var(--t3)" id="tlinePct"></span>
    </div>
    <div class="tline-track">
      <div class="tline-fill" id="tlineFill"></div>
      <div class="tline-marker" id="tlineMarker"><span class="tline-now">HOY</span></div>
    </div>
    <div class="tline-meta"><span id="tlineStart"></span><span id="tlineEnd"></span></div>
  </div>
  <div class="fbar">
    <span class="flbl">Avance</span>
    <div class="pf-group">
      <button class="pfb on" data-pf="all">Todos</button>
      <button class="pfb" data-pf="done">Completado</button>
      <button class="pfb" data-pf="prog">En progreso</button>
      <button class="pfb" data-pf="none">Sin iniciar</button>
    </div>
  </div>
  <div class="view-toggle">
    <button class="vbtn on" id="vbtn-area">Por Proyecto</button>
    <button class="vbtn" id="vbtn-tareas">Por Tareas</button>
  </div>
  <div class="task-view" id="taskView">
    <div class="tv-filters">
      <span class="flbl">Proyecto</span>
      <button class="fb on" data-tg="all">Todos</button>
      <span class="flbl" style="margin-left:6px">Avance</span>
      <button class="fb on" data-tp="all">Todos</button>
      <button class="fb" data-tp="done">Completado</button>
      <button class="fb" data-tp="prog">En progreso</button>
      <button class="fb" data-tp="none">Sin iniciar</button>
    </div>
    <div style="background:var(--s1);border:1px solid var(--border);border-radius:6px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.05)">
      <table class="ttable">
        <thead class="ttable-head"><tr>
          <th style="width:60px">Nº</th><th style="width:150px">Proyecto</th><th>Tarea</th>
          <th style="width:160px">Avance</th><th style="width:140px">Responsable</th>
          <th style="width:100px">Inicio</th><th style="width:100px">Término</th>
        </tr></thead>
        <tbody id="taskTableBody"></tbody>
      </table>
    </div>
  </div>
  <div class="area-view" id="areaView">
    <div class="slbl">Por proyecto — selecciona uno para ver sus tareas</div>
    <div class="layout">
      <div id="cards"></div>
      <div class="detail" id="detail"><div class="detail-empty">← Selecciona un proyecto<br>para ver el detalle de tareas</div></div>
    </div>
  </div>
</div>
<script>
var GROUPS = __GROUPS_JSON__;
var PROJECT_START = '__PROJECT_START__';
var PROJECT_END   = '__PROJECT_END__';
var PROJECT_PCT   = __PROJECT_PCT__;
var EXPORTED      = '__EXPORTED__';
var COLORS = ['#2563eb','#059669','#7c3aed','#dc2626','#b45309','#0891b2','#16a34a','#9333ea','#ea580c','#0d9488','#4f46e5','#db2777','#65a30d','#d97706','#7c3aed','#2563eb','#059669','#dc2626','#0891b2','#9333ea','#16a34a','#ea580c','#4f46e5'];
function colorFor(g){return COLORS[parseInt(g.scheme)-1]||'#2563eb';}
function progColor(p){return p===100?'var(--green)':p>0?'var(--blue)':'var(--t3)';}
function fmtDate(d){if(!d)return '—';var p=d.split('-');return p[2]+'/'+p[1]+'/'+p[0].slice(2);}
function progBucket(p){return p===100?'done':p>0?'prog':'none';}
document.getElementById('hDate').textContent=new Date().toLocaleDateString('es-CL',{weekday:'short',day:'numeric',month:'short',year:'numeric'}).toUpperCase();
var expDate=new Date(EXPORTED);
document.getElementById('hExp').textContent='Datos: '+expDate.toLocaleDateString('es-CL',{day:'numeric',month:'short'}).toUpperCase()+' '+expDate.toLocaleTimeString('es-CL',{hour:'2-digit',minute:'2-digit'});
(function(){
  var s=new Date(PROJECT_START),e=new Date(PROJECT_END),now=new Date();
  var total=e-s,elapsed=Math.min(Math.max(now-s,0),total);
  var pct=Math.round(elapsed/total*100);
  document.getElementById('tlineFill').style.width=pct+'%';
  document.getElementById('tlineMarker').style.left=pct+'%';
  document.getElementById('tlineStart').textContent=fmtDate(PROJECT_START);
  document.getElementById('tlineEnd').textContent=fmtDate(PROJECT_END);
  document.getElementById('tlinePct').textContent=pct+'% del plazo transcurrido';
})();
(function(){
  var allLeaf=[];
  GROUPS.forEach(function(g){
    var sc=g.tasks.map(function(t){return t.scheme;});
    g.tasks.forEach(function(t){if(!sc.some(function(s){return s!==t.scheme&&s.startsWith(t.scheme+'.');}))allLeaf.push(t);});
  });
  var done=allLeaf.filter(function(t){return t.progress===100;}).length;
  var prog=allLeaf.filter(function(t){return t.progress>0&&t.progress<100;}).length;
  var none=allLeaf.filter(function(t){return t.progress===0;}).length;
  document.getElementById('kAvance').textContent=PROJECT_PCT+'%';
  document.getElementById('kAvanceSub').textContent='Avance global del proyecto';
  document.getElementById('kTareas').innerHTML=GROUPS.length+'<span style="font-size:14px;color:var(--t3)"> proy / '+allLeaf.length+' tar</span>';
  document.getElementById('kTareasSub').textContent=done+' completadas';
  document.getElementById('kProg').textContent=prog;
  document.getElementById('kProgSub').textContent=none+' sin iniciar';
  var s=new Date(PROJECT_START),e=new Date(PROJECT_END),now=new Date();
  var diasTotal=Math.round((e-s)/86400000),diasRestantes=Math.max(0,Math.round((e-now)/86400000));
  document.getElementById('kPlazo').textContent=diasRestantes+'d';
  document.getElementById('kPlazoSub').textContent=diasTotal+' días totales';
})();
var fProg='all',curId=null;
document.querySelectorAll('[data-pf]').forEach(function(btn){btn.addEventListener('click',function(){fProg=btn.dataset.pf;document.querySelectorAll('[data-pf]').forEach(function(b){b.classList.remove('on');});btn.classList.add('on');renderCards();});});
document.getElementById('vbtn-area').addEventListener('click',function(){document.getElementById('vbtn-area').classList.add('on');document.getElementById('vbtn-tareas').classList.remove('on');document.getElementById('taskView').classList.remove('active');document.getElementById('areaView').classList.remove('hidden');});
document.getElementById('vbtn-tareas').addEventListener('click',function(){document.getElementById('vbtn-tareas').classList.add('on');document.getElementById('vbtn-area').classList.remove('on');document.getElementById('taskView').classList.add('active');document.getElementById('areaView').classList.add('hidden');renderTaskTable();});
function renderCards(){
  var list=GROUPS.filter(function(g){return fProg==='all'||progBucket(g.progress)===fProg;});
  var container=document.getElementById('cards');container.innerHTML='';
  if(list.length===0){container.innerHTML='<div style="padding:30px;text-align:center;font-family:var(--mono);font-size:10px;color:var(--t3)">Sin proyectos para el filtro seleccionado</div>';return;}
  list.forEach(function(g){
    var col=colorFor(g),pc=progColor(g.progress);
    var div=document.createElement('div');div.className='card'+(g.id===curId?' active':'');div.setAttribute('data-id',g.id);
    var datesHtml=(g.start||g.end)?fmtDate(g.start)+' → '+fmtDate(g.end):'Sin fechas definidas';
    div.innerHTML='<div class="card-top" style="background:'+col+'"></div><div class="card-hdr"><div style="display:flex;align-items:flex-start;gap:10px;flex:1;min-width:0"><span class="cnum">'+g.scheme+'</span><div style="min-width:0"><div class="cname">'+g.label+'</div><div class="cdates">'+datesHtml+'</div></div></div><div style="font-family:var(--mono);font-size:11px;font-weight:600;color:'+pc+';flex-shrink:0;margin-top:2px">'+g.progress+'%</div></div><div class="ptrack"><div class="pfill" style="width:'+g.progress+'%;background:'+col+'"></div></div><div class="stats"><div class="sb"><div class="sv" style="color:'+col+'">'+g.progress+'%</div><div class="sl">Avance</div></div><div class="sb"><div class="sv">'+g.tasksDone+'/'+g.tasksTotal+'</div><div class="sl">Tareas</div></div><div class="sb"><div class="sv" style="color:var(--blue)">'+g.tasksInProgress+'</div><div class="sl">En progreso</div></div></div><div class="cfoot"><span style="color:var(--t3);font-family:var(--mono);font-size:9px">'+(g.responsible||'—')+'</span><span class="ccta">Ver tareas →</span></div>';
    div.addEventListener('click',function(){curId=g.id;document.querySelectorAll('.card').forEach(function(c){c.classList.remove('active');});div.classList.add('active');renderDetail(g);});
    container.appendChild(div);
  });
}
function renderDetail(g){
  var col=colorFor(g);
  var datesHtml=(g.start||g.end)?fmtDate(g.start)+' → '+fmtDate(g.end):'Sin fechas definidas';
  var taskHtml='<div class="tasklist">';
  if(g.tasks.length===0){taskHtml+='<div class="tno-tasks">Sin subtareas registradas</div>';}
  else{g.tasks.forEach(function(t){var tc=progColor(t.progress);var dc=t.depth>=4?'depth-4':t.depth===3?'depth-3':'';taskHtml+='<div class="trow '+dc+'"><div><div class="tname">'+t.name+'</div>'+(t.responsible?'<div class="tsub">'+t.responsible+'</div>':'')+(t.start?'<div class="tsub-date">'+fmtDate(t.start)+' → '+fmtDate(t.end)+'</div>':'')+'</div><div class="tbar-wrap"><div class="tbar"><div class="tbar-fill" style="width:'+t.progress+'%;background:'+tc+'"></div></div></div><div class="tpct" style="color:'+tc+'">'+t.progress+'%</div></div>';});}
  taskHtml+='</div>';
  document.getElementById('detail').innerHTML='<div class="detail-head"><div><div class="dh-title">'+g.scheme+'. '+g.label+'</div><div class="dh-dates">'+datesHtml+'</div>'+(g.responsible?'<div class="dh-sub">'+g.responsible+'</div>':'')+'</div></div><div class="dkpis"><div class="dk"><div class="dk-val" style="color:'+col+'">'+g.progress+'%</div><div class="dk-lbl">Avance</div></div><div class="dk"><div class="dk-val">'+g.tasksDone+'/'+g.tasksTotal+'</div><div class="dk-lbl">Completadas</div></div><div class="dk"><div class="dk-val" style="color:var(--blue)">'+g.tasksInProgress+'</div><div class="dk-lbl">En Progreso</div></div></div><div class="d-progress-wrap"><div class="d-ptrack"><div style="height:100%;width:'+g.progress+'%;background:'+col+';border-radius:99px"></div></div></div>'+taskHtml;
}
renderCards();
var tvGroup='all',tvProg='all';
document.querySelectorAll('[data-tg]').forEach(function(btn){btn.addEventListener('click',function(){tvGroup=btn.dataset.tg;document.querySelectorAll('[data-tg]').forEach(function(b){b.classList.remove('on');});btn.classList.add('on');renderTaskTable();});});
document.querySelectorAll('[data-tp]').forEach(function(btn){btn.addEventListener('click',function(){tvProg=btn.dataset.tp;document.querySelectorAll('[data-tp]').forEach(function(b){b.classList.remove('on');});btn.classList.add('on');renderTaskTable();});});
(function(){var bar=document.querySelector('[data-tg="all"]').parentNode;GROUPS.forEach(function(g){var b=document.createElement('button');b.className='fb';b.setAttribute('data-tg',g.id);b.textContent=g.scheme+'. '+g.label.slice(0,22)+(g.label.length>22?'…':'');bar.appendChild(b);b.addEventListener('click',function(){tvGroup=g.id;document.querySelectorAll('[data-tg]').forEach(function(x){x.classList.remove('on');});b.classList.add('on');renderTaskTable();});});})();
function renderTaskTable(){
  var body=document.getElementById('taskTableBody');body.innerHTML='';var rows=[];
  GROUPS.forEach(function(g){if(tvGroup!=='all'&&g.id!==tvGroup)return;g.tasks.forEach(function(t){if(t.depth<2)return;if(tvProg!=='all'&&progBucket(t.progress)!==tvProg)return;rows.push({g:g,t:t});});});
  if(rows.length===0){var tr=document.createElement('tr');tr.innerHTML='<td colspan="7" style="text-align:center;padding:30px;font-family:var(--mono);font-size:10px;color:var(--t3)">Sin tareas para el filtro seleccionado</td>';body.appendChild(tr);return;}
  var byGroup={},order=[];rows.forEach(function(r){if(!byGroup[r.g.id]){byGroup[r.g.id]=[];order.push(r.g);}byGroup[r.g.id].push(r.t);});
  var seen={};order=order.filter(function(g){if(seen[g.id])return false;seen[g.id]=true;return true;});
  order.forEach(function(g){
    var hrow=document.createElement('tr');hrow.className='tgroup-row';hrow.innerHTML='<td colspan="7">'+g.scheme+'. '+g.label+' — '+byGroup[g.id].length+' tarea'+(byGroup[g.id].length!==1?'s':'')+'</td>';body.appendChild(hrow);
    byGroup[g.id].forEach(function(t){var tc=progColor(t.progress);var indent=t.depth>2?(t.depth-2)*12:0;var tr=document.createElement('tr');
      var td1=document.createElement('td');td1.innerHTML='<span class="tt-scheme">'+t.scheme+'</span>';
      var td2=document.createElement('td');td2.innerHTML='<span class="tt-group" title="'+g.label+'">'+g.label+'</span>';
      var td3=document.createElement('td');td3.innerHTML='<div class="tt-name" style="padding-left:'+indent+'px">'+t.name+'</div>';
      var td4=document.createElement('td');td4.innerHTML='<div class="tt-bar-wrap"><div class="tt-bar"><div class="tt-fill" style="width:'+t.progress+'%;background:'+tc+'"></div></div><div class="tt-pct" style="color:'+tc+'">'+t.progress+'%</div></div>';
      var td5=document.createElement('td');td5.innerHTML='<span class="tt-resp" title="'+(t.responsible||'')+'">'+((t.responsible||'—').split(',')[0].trim()||'—')+'</span>';
      var td6=document.createElement('td');td6.innerHTML='<span style="font-family:var(--mono);font-size:10px;color:var(--t2)">'+fmtDate(t.start)+'</span>';
      var td7=document.createElement('td');td7.innerHTML='<span style="font-family:var(--mono);font-size:10px;color:var(--t2)">'+fmtDate(t.end)+'</span>';
      [td1,td2,td3,td4,td5,td6,td7].forEach(function(td){tr.appendChild(td);});body.appendChild(tr);});
  });
}
</script>
</body>
</html>
"""

def generar_html(data):
    groups_json = json.dumps(data["groups"], ensure_ascii=False)
    html = HTML_TEMPLATE
    html = html.replace("__GROUPS_JSON__", groups_json)
    html = html.replace("__PROJECT_START__", data["proyecto_inicio"])
    html = html.replace("__PROJECT_END__", data["proyecto_fin"])
    html = html.replace("__PROJECT_PCT__", str(data["proyecto_pct"]))
    html = html.replace("__EXPORTED__", data["exportado"])
    return html

# ── MAIN ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    fuente = cargar_excel()
    print("Parseando Excel...")
    data = parsear_excel(fuente)
    print(f"  → {len(data['groups'])} proyectos, avance global: {data['proyecto_pct']}%")
    print("Generando index.html...")
    html = generar_html(data)
    with open("index.html", "w", encoding="utf-8") as f:
        f.write(html)
    print(f"  → index.html generado ({len(html):,} bytes)")
    print("¡Listo!")
