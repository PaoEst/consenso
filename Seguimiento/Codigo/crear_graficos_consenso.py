#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera, por cuenca, el gráfico interactivo para el consenso de expertos:
  izquierda  = seguimiento (aportes + drivers observados)
  derecha    = pronóstico de drivers y cuadros de votación (S1..S5) del evaluador

Entradas (en la carpeta del script):
    Seguimiento_Aportes.xlsx, Seguimiento_Drivers.xlsx, Pronostico_Drivers.xlsx, Consenso.xlsx

Salida:
    <año>/<mes>/<AAAA-MM-DD>/Consenso_<cuenca>.html     (doble clic para abrir)
    <año>/<mes>/<AAAA-MM-DD>/votos/                     (aquí se dejan los archivos de voto)
    logs/ejecucion.log

La fecha de la carpeta es la "Fecha de realización del consenso" de Consenso.xlsx > Fechas
(se puede forzar con --fecha-corrida AAAA-MM-DD).

Uso:  python crear_graficos_consenso.py
"""
from __future__ import annotations

import argparse
import getpass
import html
import json
import platform
import sys
import time
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from consenso_comun import (ARCHIVO_CONSENSO, ARCHIVO_PRONOSTICO, BASE, CARPETA_CONSOLIDADO,
                            ESTILO_NAV, SALIDAS, agregar_barras_drivers, altura_panel_drivers,
                            barra_navegacion, cargar_drivers, carpeta_corrida, drivers_combinados,
                            escribir_indice, etiqueta_semana_corta, leer_evaluadores, leer_fechas,
                            plotly_js, real_a_la_fecha, registrar_cuenca)
from graficar_aportes import (ARCHIVO_APORTES, ARCHIVO_DRIVERS, C_ABAJO, C_ARRIBA, C_LINEA,
                              configurar_logging, definir_ventana, leer_aportes, log,
                              nombre_archivo, relleno_media, sha256)

C_PRONOSTICO = "#1f5f8b"


# --------------------------------------------------------------------------
# Figura
# --------------------------------------------------------------------------
def construir_figura(cuenca, serie, seg, orden_vars, ini, fechas):
    hor_ini, hor_fin = fechas["hor_ini"], fechas["hor_fin"]
    x_fin = hor_fin + pd.Timedelta(days=1)
    n = max(len(orden_vars), 1)
    alto_drv = altura_panel_drivers(n)
    alto_tot = int(400 + alto_drv + 100)
    frac = min(0.62, max(0.25, alto_drv / alto_tot))
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.04,
                        row_heights=[1 - frac, frac])

    # Aportes observados (dato diario al centro de su día, igual que las barras)
    x = list(serie["Fecha"] + pd.Timedelta(hours=12))
    y = list(serie[cuenca])
    xs, arr, abj = relleno_media(x, y)
    base = [100.0] * len(xs)
    for techo, color in ((arr, C_ARRIBA), (abj, C_ABAJO)):
        fig.add_trace(go.Scatter(x=xs, y=base, mode="lines", line=dict(width=0),
                                 hoverinfo="skip", showlegend=False), row=1, col=1)
        fig.add_trace(go.Scatter(x=xs, y=techo, mode="lines", line=dict(width=0), fill="tonexty",
                                 fillcolor=color, hoverinfo="skip", showlegend=False), row=1, col=1)
    fig.add_trace(go.Scatter(
        x=x, y=y, mode="lines", line=dict(color=C_LINEA, width=2.2), showlegend=False,
        hovertemplate="%{x|%d %b %Y}<br><b>%{y:.1f} %</b> de la media<extra></extra>"),
        row=1, col=1)
    fig.add_hline(y=100, line=dict(color="#333", width=1, dash="dash"), row=1, col=1,
                  annotation_text="Media histórica (100 %)", annotation_position="top left",
                  annotation_font=dict(size=11, color="#444"))

    # Pronóstico del evaluador (se actualiza desde JS): línea escalonada y etiquetas
    idx_linea = len(fig.data)
    fig.add_trace(go.Scatter(x=[None], y=[None], mode="lines", showlegend=False, connectgaps=False,
                             line=dict(color=C_PRONOSTICO, width=3), hoverinfo="skip"), row=1, col=1)
    idx_texto = len(fig.data)
    fig.add_trace(go.Scatter(x=[None], y=[None], mode="text", showlegend=False, hoverinfo="skip",
                             textposition="top center",
                             textfont=dict(color=C_PRONOSTICO, size=12)), row=1, col=1)

    # Barras de drivers (una fila por variable; seguimiento a la izquierda, pronóstico a la derecha)
    agregar_barras_drivers(fig, seg, orden_vars, row=2)

    # Separación seguimiento / pronóstico y semanas
    formas = [dict(type="rect", xref="x", yref="paper", x0=hor_ini, x1=x_fin, y0=0, y1=1,
                   fillcolor="rgba(31,95,139,0.05)", line=dict(width=0), layer="below"),
              dict(type="line", xref="x", yref="paper", x0=hor_ini, x1=hor_ini, y0=0, y1=1.0,
                   line=dict(color="#d62728", width=2, dash="dash"))]
    for s in fechas["semanas"]:
        if s["ini"] > hor_ini:
            formas.append(dict(type="line", xref="x", yref="paper", x0=s["ini"], x1=s["ini"],
                               y0=0, y1=1, line=dict(color="#2e9e4f", width=1, dash="dot")))
    anot = [
        dict(x=ini + (hor_ini - ini) / 2, y=1.0, xref="x", yref="paper", yanchor="bottom",
             text="<b>Seguimiento</b>", showarrow=False, font=dict(size=13)),
        dict(x=hor_ini + (x_fin - hor_ini) / 2, y=1.0, xref="x", yref="paper", yanchor="bottom",
             text=f"<b>Pronóstico</b> (inicia {hor_ini:%d-%m-%y})", showarrow=False,
             font=dict(size=13)),
    ]
    rango = [ini, x_fin]
    fig.update_yaxes(title_text="Aportes (% de la media)", row=1, col=1, gridcolor="#e6e6e6",
                     zeroline=False, rangemode="tozero")
    for fila in (1, 2):
        fig.update_xaxes(range=rango, matches="x" if fila == 2 else None, gridcolor="#e6e6e6",
                         showspikes=True, spikemode="across", spikethickness=1,
                         spikecolor="#888", tickformat="%d %b\n%Y", row=fila, col=1)
    fig.update_layout(
        height=alto_tot, template="plotly_white", hovermode="closest", shapes=formas,
        annotations=list(fig.layout.annotations) + anot, hoverlabel=dict(align="left"),
        margin=dict(l=20, r=20, t=40, b=60), plot_bgcolor="#fbfbf9",
        legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.1))
    return fig, idx_linea, idx_texto


# --------------------------------------------------------------------------
# Página HTML (votación)
# --------------------------------------------------------------------------
PLANTILLA = r"""<!doctype html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITULO_PAGINA__</title>
__ESTILO_NAV__
<style>
:root{--naranja:#e8742a;--azul:#1f5f8b;--txt:#1d2733;--sub:#5b6773;--bg:#ffffff;--card:#f6f7f9;--brd:#d9dde3}
@media (prefers-color-scheme: dark){:root{--txt:#e8ecf1;--sub:#a3adb8;--bg:#161b22;--card:#1f262f;--brd:#34404d}}
*{box-sizing:border-box} body{margin:0;font-family:Segoe UI,Arial,sans-serif;color:var(--txt);background:var(--bg)}
header{display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:12px;padding:12px 20px;border-bottom:1px solid var(--brd)}
h1{font-size:20px;margin:0;font-weight:600;flex:1 1 100%;text-align:center}
.campo{display:flex;align-items:center;gap:8px;font-size:13px;color:var(--sub)}
input[type=text],input[type=number]{border:2px solid var(--naranja);border-radius:6px;padding:6px 8px;font-size:14px;color:var(--txt);background:var(--bg);width:90px;text-align:center}
input:disabled{border-color:var(--brd);background:var(--card);cursor:not-allowed}
#evname{font-size:13px;color:var(--azul);min-width:130px;font-weight:600}
#guardar{background:#d62728;color:#fff;border:0;border-radius:6px;padding:9px 22px;font-weight:700;letter-spacing:.5px;cursor:pointer}
#guardar:disabled{background:var(--brd);color:var(--sub);cursor:not-allowed}
#main{display:flex;gap:16px;padding:10px 20px 20px}
#izq{flex:1 1 auto;min-width:0} #der{flex:0 0 220px;display:flex;flex-direction:column;gap:12px;padding-top:70px}
#semrow{position:relative;height:74px}
.sem{position:absolute;top:0;width:58px;text-align:center;font-size:11px;line-height:1.25;color:var(--sub);transform:translateX(-50%)}
.sem b{display:block;color:var(--txt);font-size:13px}
.sem input{width:58px;margin-top:4px;padding:6px 2px}
.card{background:var(--card);border:1px solid var(--brd);border-radius:8px;padding:10px 12px;font-size:13px;color:var(--sub)}
.card input{width:100%;margin-top:6px;font-size:16px}
.valor{margin-top:6px;border:2px solid var(--brd);border-radius:6px;padding:6px;text-align:center;font-size:16px;color:var(--txt);background:var(--bg);min-height:36px}
#msg{font-size:13px;min-height:20px} .ok{color:#1a7f37}.err{color:#d62728}
</style>
<script>__PLOTLYJS__</script></head><body>
__NAV__
<header>
 <h1>__TITULO__</h1>
 <div class="campo">Iniciales Evaluador Técnico <input id="ini" type="text" maxlength="6" autocomplete="off" placeholder="p. ej. PE"><span id="evname"></span></div>
 <div class="campo">No Evaluación <input id="nev" type="number" min="1" step="1" value="1" disabled></div>
 <button id="guardar" disabled>GUARDAR</button>
</header>
<div id="main">
 <div id="izq"><div id="semrow"></div><div id="grafico">__GRAFICO__</div></div>
 <aside id="der">
  <div class="card">Real a la fecha (%)<div class="valor" id="real">—</div></div>
  <div class="card">Esperado el resto del mes (%)<input id="resto" type="number" min="0" step="any" disabled></div>
  <div class="card">Cierre del mes (%)<div class="valor" id="cierre">—</div></div>
  <div id="msg"></div>
 </aside>
</div>
<script>
const D = __DATOS__;
const gd = document.getElementById('grafico').getElementsByClassName('plotly-graph-div')[0];
const $ = id => document.getElementById(id);
const semInputs = [];
const DAY = 86400000;
const fmt = v => (v===null||v===undefined||isNaN(v)) ? '' : String(Math.round(v*10)/10);
const num = s => { s=(s||'').toString().replace(',','.').trim(); if(s==='') return null; const v=Number(s); return isFinite(v)?v:NaN; };

// ---- cuadros por semana, alineados al eje de fechas del gráfico
D.semanas.forEach(s => {
  const div = document.createElement('div'); div.className='sem'; div.dataset.n=s.n;
  div.innerHTML = '<b>S'+s.n+'</b>'+s.etiqueta+'<input type="number" min="0" step="any" disabled>';
  $('semrow').appendChild(div);
  const inp = div.querySelector('input'); inp.dataset.n=s.n; semInputs.push(inp);
  inp.addEventListener('input', () => recalcular('semana'));
});
function posicionar(){
  const xa = gd._fullLayout && gd._fullLayout.xaxis; if(!xa || !xa._length) return;
  const SEP = 68;               // separación mínima entre centros de cajones (px)
  let previo = -1e9;
  D.semanas.forEach(s => {
    const div = document.querySelector('.sem[data-n="'+s.n+'"]');
    const a = xa._offset + xa.d2p(s.ini_ms), b = xa._offset + xa.d2p(s.fin_ms + DAY);
    const centro = Math.max((a+b)/2, previo + SEP);   // si dos semanas quedan muy juntas, se corre la siguiente
    previo = centro;
    div.style.left = centro + 'px';
    div.style.display = (b<xa._offset || a>xa._offset+xa._length) ? 'none' : '';
  });
}
gd.on('plotly_afterplot', posicionar);
window.addEventListener('resize', () => setTimeout(posicionar, 50));
// el gráfico se dibuja mientras la página aún se está leyendo: se reajusta al ancho real al cargar
window.addEventListener('load', () => setTimeout(() => Plotly.Plots.resize(gd), 50));
(function esperar(){ if(gd._fullLayout && gd._fullLayout.xaxis && gd._fullLayout.xaxis._length) posicionar(); else setTimeout(esperar, 60); })();

// ---- cálculos
$('real').textContent = D.real.mean===null ? '—' : fmt(D.real.mean);
function valoresSemanas(){ return semInputs.map(i => num(i.value)); }
function recalcular(origen){
  if(origen==='resto'){
    const r = num($('resto').value);
    if(r!==null && !isNaN(r)) semInputs.forEach(i => i.value = r);
  }
  const v = valoresSemanas();
  let sw=0, sd=0, completas=true;
  v.forEach((x,k) => { if(x===null||isNaN(x)) completas=false; else { sw += x*D.semanas[k].dias; sd += D.semanas[k].dias; } });
  if(origen!=='resto') $('resto').value = sd>0 ? fmt(sw/sd) : '';
  if(completas && sd>0){
    const cierre = (D.real.suma + sw) / (D.real.n + sd);
    $('cierre').textContent = fmt(cierre);
  } else $('cierre').textContent = '—';
  dibujar(v);
}
function dibujar(v){
  const xs=[], ys=[], tx=[], ty=[], tt=[];
  D.semanas.forEach((s,k) => {
    const x = v[k];
    if(x===null||isNaN(x)) { if(xs.length && xs[xs.length-1]!==null){ xs.push(null); ys.push(null);} return; }
    xs.push(s.ini_ms, s.fin_ms+DAY); ys.push(x, x);
    tx.push((s.ini_ms + s.fin_ms + DAY)/2); ty.push(x); tt.push(fmt(x));
  });
  Plotly.restyle(gd, {x:[xs], y:[ys]}, [D.idx_linea]);
  Plotly.restyle(gd, {x:[tx], y:[ty], text:[tt]}, [D.idx_texto]);
}

// ---- evaluador autorizado
function validarEvaluador(){
  const k = $('ini').value.trim().toUpperCase();
  const nombre = D.evaluadores[k];
  const ok = !!nombre;
  $('evname').textContent = ok ? nombre : (k ? 'Iniciales no autorizadas' : '');
  $('evname').style.color = ok ? '' : '#d62728';
  [$('nev'), $('resto'), $('guardar'), ...semInputs].forEach(e => e.disabled = !ok);
  return ok ? k : null;
}
$('ini').addEventListener('input', validarEvaluador);
$('resto').addEventListener('input', () => recalcular('resto'));

// ---- guardar: descarga un archivo de voto
function mensaje(t, cls){ $('msg').textContent=t; $('msg').className=cls||''; }
function uuid(){ return (window.crypto && crypto.randomUUID) ? crypto.randomUUID() : 'v'+Date.now().toString(36)+Math.random().toString(36).slice(2,10); }
$('guardar').addEventListener('click', () => {
  const k = validarEvaluador(); if(!k){ mensaje('Iniciales no autorizadas.','err'); return; }
  const nev = Number($('nev').value);
  if(!Number.isInteger(nev) || nev<1){ mensaje('El No de evaluación debe ser un entero ≥ 1.','err'); return; }
  const v = valoresSemanas();
  if(v.some(x => x===null || isNaN(x) || x<0)){ mensaje('Completa todas las semanas con valores numéricos ≥ 0 (o escribe el esperado del resto del mes).','err'); return; }
  const sem = {}; for(let i=1;i<=5;i++) sem['S'+i] = null;
  D.semanas.forEach((s,j) => sem['S'+s.n] = v[j]);
  let sw=0, sd=0; D.semanas.forEach((s,j) => { sw+=v[j]*s.dias; sd+=s.dias; });
  const ahora = new Date(); const pad = n => String(n).padStart(2,'0');
  const local = ahora.getFullYear()+'-'+pad(ahora.getMonth()+1)+'-'+pad(ahora.getDate())+'T'+pad(ahora.getHours())+':'+pad(ahora.getMinutes())+':'+pad(ahora.getSeconds());
  const voto = {
    id: uuid(), fecha_hora: local, evaluador: k, nombre: D.evaluadores[k], no_evaluacion: nev,
    cuenca: D.cuenca, semanas: sem, real_a_la_fecha: D.real.mean,
    resto_del_mes: Math.round(sw/sd*100)/100,
    cierre_de_mes: Math.round(((D.real.suma+sw)/(D.real.n+sd))*100)/100,
    mes_pronostico: D.mes_pronostico, horizonte: D.horizonte, carpeta_corrida: D.carpeta
  };
  const nombre = 'voto_'+D.archivo+'_'+k+'_E'+nev+'_'+local.replace(/[-:]/g,'')+'.json';
  const blob = new Blob([JSON.stringify(voto,null,2)], {type:'application/json'});
  const a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombre;
  document.body.appendChild(a); a.click(); a.remove(); setTimeout(()=>URL.revokeObjectURL(a.href), 2000);
  mensaje('Voto guardado en el archivo "'+nombre+'". Déjalo en la carpeta "votos" de la corrida.','ok');
});
validarEvaluador(); dibujar(valoresSemanas());
</script></body></html>
"""


def js_seguro(obj) -> str:
    return json.dumps(obj, ensure_ascii=False).replace("</", "<\\/")


def escribir_html(ruta: Path, cuenca: str, fig, idx_linea, idx_texto, fechas, evaluadores,
                  obs: list[float], carpeta: Path) -> None:
    archivo = nombre_archivo(cuenca)
    registrar_cuenca(carpeta, archivo, cuenca)
    ms = lambda f: int(f.timestamp() * 1000)  # f es medianoche local-naive; plotly lo trata igual
    semanas = [{"n": s["n"], "dias": s["dias"], "ini_ms": ms(s["ini"]), "fin_ms": ms(s["fin"]),
                "etiqueta": etiqueta_semana_corta(s)}
               for s in fechas["semanas"]]
    datos = {
        "cuenca": cuenca, "archivo": nombre_archivo(cuenca), "semanas": semanas,
        "evaluadores": evaluadores, "idx_linea": idx_linea, "idx_texto": idx_texto,
        "real": {"n": len(obs), "suma": float(sum(obs)),
                 "mean": (sum(obs) / len(obs)) if obs else None},
        "mes_pronostico": f"{fechas['hor_ini']:%Y-%m}",
        "horizonte": [f"{fechas['hor_ini']:%Y-%m-%d}", f"{fechas['hor_fin']:%Y-%m-%d}"],
        "carpeta": str(carpeta.name),
    }
    titulo = f"Aportes {cuenca} y moduladores climáticos"
    div = fig.to_html(full_html=False, include_plotlyjs=False,
                      config={"displaylogo": False, "responsive": True})
    pagina = (PLANTILLA.replace("__TITULO_PAGINA__", html.escape(f"Consenso - {cuenca}"))
              .replace("__TITULO__", html.escape(titulo))
              .replace("__ESTILO_NAV__", ESTILO_NAV)
              .replace("__NAV__", "<!--NAV:%s-->" % archivo)
              .replace("__DATOS__", js_seguro(datos))
              .replace("__GRAFICO__", div)
              .replace("__PLOTLYJS__", plotly_js()))
    ruta.write_text(pagina, encoding="utf-8")


# --------------------------------------------------------------------------
# Principal
# --------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--aportes", default=BASE / ARCHIVO_APORTES, type=Path)
    ap.add_argument("--drivers", default=BASE / ARCHIVO_DRIVERS, type=Path)
    ap.add_argument("--pronostico", default=BASE / ARCHIVO_PRONOSTICO, type=Path)
    ap.add_argument("--consenso", default=BASE / ARCHIVO_CONSENSO, type=Path)
    ap.add_argument("--salida", default=SALIDAS, type=Path, help="carpeta donde se crea año/mes/fecha")
    ap.add_argument("--logs", default=BASE / "logs", type=Path)
    ap.add_argument("--fecha-corrida", default=None, help="AAAA-MM-DD (por defecto: fecha de realización del consenso)")
    args = ap.parse_args()

    run_id = uuid.uuid4().hex[:8]
    t0 = time.time()
    ruta_log = configurar_logging(args.logs, run_id)
    log.info("=" * 70)
    log.info("INICIO crear_graficos_consenso run_id=%s", run_id)
    log.info("Usuario=%s host=%s SO=%s python=%s pandas=%s plotly=%s", getpass.getuser(),
             platform.node(), platform.platform(), platform.python_version(),
             pd.__version__, plotly.__version__)
    log.info("Argumentos: %s", {k: str(v) for k, v in vars(args).items()})
    try:
        for p in (args.aportes, args.drivers, args.pronostico, args.consenso):
            if not p.exists():
                raise FileNotFoundError(f"No se encontró el archivo Excel: {p} (debe estar en la carpeta del script)")
            st = p.stat()
            log.info("Entrada: %s | %d bytes | modificado %s | sha256=%s", p.resolve(), st.st_size,
                     datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"), sha256(p))

        fechas = leer_fechas(args.consenso)
        evaluadores = leer_evaluadores(args.consenso)
        if not evaluadores:
            raise ValueError("Consenso.xlsx > Evaluadores_Tecnicos no tiene evaluadores")
        aportes = leer_aportes(args.aportes)
        ini, _ = definir_ventana(args.aportes, aportes)
        if ini >= fechas["hor_ini"]:
            log.warning("Fecha de inicio de gráficos (%s) >= inicio del horizonte (%s); se usa 30 días antes",
                        ini.date(), fechas["hor_ini"].date())
            ini = fechas["hor_ini"] - pd.Timedelta(days=30)

        ctx = cargar_drivers(args.drivers, args.pronostico)

        fecha_corrida = (pd.Timestamp(args.fecha_corrida).normalize() if args.fecha_corrida
                         else fechas["consenso"])
        carpeta = carpeta_corrida(args.salida, fecha_corrida)
        (carpeta / "votos").mkdir(parents=True, exist_ok=True)
        (carpeta / CARPETA_CONSOLIDADO).mkdir(exist_ok=True)
        log.info("Carpeta de corrida: %s", carpeta.resolve())

        generados, omitidos = [], []
        for cuenca in [c for c in aportes.columns if c != "Fecha"]:
            col = aportes[cuenca]
            if col.notna().sum() == 0:
                log.warning("Cuenca '%s': sin ningún dato -> no se genera gráfico", cuenca)
                omitidos.append((cuenca, "sin datos"))
                continue
            serie = aportes.loc[(aportes["Fecha"] >= ini) & (aportes["Fecha"] <= fechas["hor_fin"]),
                                ["Fecha", cuenca]].dropna()
            if serie.empty:
                log.warning("Cuenca '%s': sin datos en la ventana; no se genera gráfico", cuenca)
                omitidos.append((cuenca, "sin datos en la ventana"))
                continue

            seg, orden_vars = drivers_combinados(cuenca, ctx, ini, fechas["hor_fin"])
            obs = real_a_la_fecha(aportes, cuenca, fechas["hor_ini"])

            fig, il, it = construir_figura(cuenca, serie, seg, orden_vars, ini, fechas)
            ruta = carpeta / f"Consenso_{nombre_archivo(cuenca)}.html"
            escribir_html(ruta, cuenca, fig, il, it, fechas, evaluadores, obs, carpeta)
            generados.append((cuenca, ruta.name))
            log.info("Cuenca '%s': %s | aportes=%d pts (%s -> %s) | real a la fecha: %d días %s | "
                     "variables=%d barras=%d (seguimiento=%d, pronóstico=%d)", cuenca, ruta.name,
                     len(serie), serie["Fecha"].min().date(), serie["Fecha"].max().date(), len(obs),
                     f"media={sum(obs) / len(obs):.2f}" if obs else "(sin datos del mes)",
                     len(orden_vars), len(seg),
                     int((seg["origen"] == "Seguimiento").sum()) if len(seg) else 0,
                     int((seg["origen"] == "Pronóstico").sum()) if len(seg) else 0)
            for _, r in seg.iterrows():
                log.debug("  barra | %s | %s | %s -> %s | nota=%r | impacto=%s | lugar=%s",
                          r["origen"], r["Variable"], r["ini"].date(), r["fin"].date(), r["nota"],
                          r["impacto"], r["lugar"])

        # barra de navegación (índice + selector de cuenca) ahora que se conocen todas las cuencas
        for archivo in [nombre_archivo(c) for c, _ in generados]:
            ruta = carpeta / f"Consenso_{archivo}.html"
            txt = ruta.read_text(encoding="utf-8")
            ruta.write_text(txt.replace("<!--NAV:%s-->" % archivo,
                                        barra_navegacion(carpeta, archivo, "votacion")), encoding="utf-8")
        escribir_indice(carpeta, fechas)

        log.info("RESUMEN: generados=%d %s | omitidos=%d %s", len(generados),
                 [c for c, _ in generados], len(omitidos), omitidos)
        log.info("FIN OK run_id=%s duración=%.2fs carpeta=%s log=%s", run_id, time.time() - t0,
                 carpeta.resolve(), ruta_log.resolve())
        print(f"\nGráficos en: {carpeta.resolve()}")
        return 0
    except Exception:
        log.exception("FIN CON ERROR run_id=%s tras %.2fs", run_id, time.time() - t0)
        return 1


if __name__ == "__main__":
    sys.exit(main())
