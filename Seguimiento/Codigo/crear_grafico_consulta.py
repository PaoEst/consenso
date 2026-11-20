#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gráfico de SOLO CONSULTA (sin votación): seguimiento de aportes y de drivers climáticos, drivers
pronosticados y otros modelos de pronóstico de aportes, con un selector de cuenca.

Es un programa independiente de la votación: no usa votos, Registros, Consolidado ni Evaluación y no
enlaza con las páginas de votación. De Consenso.xlsx lee únicamente la hoja `Fechas` (inicio y fin del
horizonte del pronóstico y las semanas S1..S5 para los pronósticos semanales de Otros_pronosticos).

Entradas (en la carpeta Insumos, junto a Codigo/):
    Seguimiento_Aportes.xlsx, Seguimiento_Drivers.xlsx, Pronostico_Drivers.xlsx,
    Consenso.xlsx (solo hoja Fechas), Otros_pronosticos.xlsx (opcional)

Salida (un único archivo, se abre con doble clic):
    Salidas/Consulta/Consulta_Seguimiento_Pronostico_AAAA-MM-DD.html     (fecha de generación)
    logs/ejecucion.log

Uso:  python crear_grafico_consulta.py [--fecha AAAA-MM-DD] [--fecha-consenso AAAA-MM-DD]
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
import plotly.io as pio
from plotly.subplots import make_subplots

from consenso_comun import (ARCHIVO_CONSENSO, ARCHIVO_OTROS, ARCHIVO_PRONOSTICO, ESTILO_NAV, INSUMOS,
                            SALIDAS, agregar_barras_drivers, agregar_otros_pronosticos,
                            altura_panel_drivers, cargar_drivers, drivers_combinados, leer_fechas,
                            leer_otros_pronosticos, pie_version, plotly_js)
from graficar_aportes import (ARCHIVO_APORTES, ARCHIVO_DRIVERS, BASE, C_ABAJO, C_ARRIBA, C_LINEA,
                              VERSION_CODIGO, configurar_logging, definir_ventana, leer_aportes, log,
                              relleno_media, sha256)


# --------------------------------------------------------------------------
# Figura de una cuenca
# --------------------------------------------------------------------------
def construir_figura(cuenca, serie, seg, orden_vars, ini, hor_ini, hor_fin, otros=()):
    x_fin = hor_fin + pd.Timedelta(days=1)
    alto_drv = altura_panel_drivers(max(len(orden_vars), 1))
    alto_tot = int(400 + alto_drv + 100)
    frac = min(0.62, max(0.25, alto_drv / alto_tot))
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.04,
                        row_heights=[1 - frac, frac])

    # Aportes observados (dato diario al centro de su día, igual que las barras de drivers)
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
        hovertemplate="%{x|%d %b %Y}<br><b>%{y:.1f} %</b> de la media<extra></extra>"), row=1, col=1)
    fig.add_hline(y=100, line=dict(color="#333", width=1, dash="dash"), row=1, col=1,
                  annotation_text="Media histórica (100 %)", annotation_position="top left",
                  annotation_font=dict(size=11, color="#444"))

    agregar_otros_pronosticos(fig, list(otros), ini, x_fin, row=1)     # modelos externos + leyenda
    agregar_barras_drivers(fig, seg, orden_vars, row=2)                # drivers: una fila por variable

    formas = [dict(type="rect", xref="x", yref="paper", x0=hor_ini, x1=x_fin, y0=0, y1=1,
                   fillcolor="rgba(31,95,139,0.05)", line=dict(width=0), layer="below"),
              dict(type="line", xref="x", yref="paper", x0=hor_ini, x1=hor_ini, y0=0, y1=1.0,
                   line=dict(color="#d62728", width=2, dash="dash"))]
    anot = [
        dict(x=ini + (hor_ini - ini) / 2, y=1.0, xref="x", yref="paper", yanchor="bottom",
             text="<b>Seguimiento</b>", showarrow=False, font=dict(size=13)),
        dict(x=hor_ini + (x_fin - hor_ini) / 2, y=1.0, xref="x", yref="paper", yanchor="bottom",
             text=f"<b>Pronóstico</b> (inicia {hor_ini:%d-%m-%y})", showarrow=False, font=dict(size=13)),
    ]
    fig.update_yaxes(title_text="Aportes (% de la media)", row=1, col=1, gridcolor="#e6e6e6",
                     zeroline=False, rangemode="tozero")
    for fila in (1, 2):
        fig.update_xaxes(range=[ini, x_fin], matches="x" if fila == 2 else None, gridcolor="#e6e6e6",
                         showspikes=True, spikemode="across", spikethickness=1, spikecolor="#888",
                         tickformat="%d %b\n%Y", row=fila, col=1)
    fig.update_layout(
        height=alto_tot, template="plotly_white", hovermode="closest", shapes=formas,
        annotations=list(fig.layout.annotations) + anot, hoverlabel=dict(align="left"),
        margin=dict(l=20, r=20, t=40, b=60), plot_bgcolor="#fbfbf9",
        legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.1))
    return fig


# --------------------------------------------------------------------------
# Página HTML: un archivo, selector de cuenca interno
# --------------------------------------------------------------------------
PLANTILLA = r"""<!doctype html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>__TITULO_PAGINA__</title>
__ESTILO_NAV__
<style>
:root{--txt:#1d2733;--sub:#5b6773;--bg:#fff;--brd:#d9dde3;--azul:#1f5f8b}
@media (prefers-color-scheme: dark){:root{--txt:#e8ecf1;--sub:#a3adb8;--bg:#161b22;--brd:#34404d;--azul:#7fb3e6}}
body{margin:0;font-family:Segoe UI,Arial,sans-serif;color:var(--txt);background:var(--bg)}
header{padding:12px 20px;border-bottom:1px solid var(--brd);text-align:center}
h1{font-size:20px;margin:0;font-weight:600} .sub{color:var(--sub);font-size:13px;margin-top:4px}
#grafico{padding:10px 20px 20px}
</style>
<script>__PLOTLYJS__</script></head><body>
<nav class="nav"><label>Cuenca <select id="sel">__OPCIONES__</select></label></nav>
<header><h1 id="titulo"></h1><div class="sub" id="sub"></div></header>
<div id="grafico"></div>
<script>
const FIGS = __FIGURAS__;
const gd = document.getElementById('grafico');
function dibujar(nombre){
  const f = FIGS[nombre];
  document.getElementById('titulo').textContent = 'Seguimiento y pronóstico · ' + nombre;
  document.getElementById('sub').textContent = f.sub;
  document.title = 'Consulta · ' + nombre;
  Plotly.react(gd, f.fig.data, f.fig.layout, {displaylogo:false, responsive:true});
}
const sel = document.getElementById('sel');
sel.addEventListener('change', () => dibujar(sel.value));
dibujar(sel.value);
window.addEventListener('load', () => setTimeout(() => Plotly.Plots.resize(gd), 50));
</script>__PIE__</body></html>
"""


def escribir_pagina(ruta: Path, figuras: dict) -> None:
    opciones = "".join(f'<option value="{html.escape(c)}">{html.escape(c)}</option>' for c in figuras)
    pagina = (PLANTILLA.replace("__TITULO_PAGINA__", "Consulta · seguimiento y pronóstico")
              .replace("__ESTILO_NAV__", ESTILO_NAV)
              .replace("__OPCIONES__", opciones)
              .replace("__PIE__", pie_version())
              .replace("__PLOTLYJS__", plotly_js())
              .replace("__FIGURAS__", json.dumps(figuras, ensure_ascii=False).replace("</", "<\\/")))
    ruta.write_text(pagina, encoding="utf-8")


# --------------------------------------------------------------------------
# Principal
# --------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--aportes", default=INSUMOS / ARCHIVO_APORTES, type=Path)
    ap.add_argument("--drivers", default=INSUMOS / ARCHIVO_DRIVERS, type=Path)
    ap.add_argument("--pronostico", default=INSUMOS / ARCHIVO_PRONOSTICO, type=Path)
    ap.add_argument("--consenso", default=INSUMOS / ARCHIVO_CONSENSO, type=Path,
                    help="solo se lee la hoja Fechas")
    ap.add_argument("--otros", default=INSUMOS / ARCHIVO_OTROS, type=Path,
                    help="Otros_pronosticos.xlsx (opcional)")
    ap.add_argument("--salida", default=SALIDAS / "Consulta", type=Path)
    ap.add_argument("--logs", default=BASE / "logs", type=Path)
    ap.add_argument("--fecha", default=None, help="AAAA-MM-DD para el nombre del archivo (por defecto: hoy)")
    ap.add_argument("--fecha-consenso", default=None,
                    help="AAAA-MM-DD: fila de la hoja Fechas a usar (por defecto: la última)")
    args = ap.parse_args()

    run_id = uuid.uuid4().hex[:8]
    t0 = time.time()
    ruta_log = configurar_logging(args.logs, run_id)
    log.info("=" * 70)
    log.info("INICIO crear_grafico_consulta run_id=%s", run_id)
    log.info("Usuario=%s host=%s SO=%s python=%s pandas=%s plotly=%s código=%s", getpass.getuser(),
             platform.node(), platform.platform(), platform.python_version(), pd.__version__,
             plotly.__version__, VERSION_CODIGO)
    log.info("Argumentos: %s", {k: str(v) for k, v in vars(args).items()})
    try:
        for p in (args.aportes, args.drivers, args.pronostico, args.consenso, args.otros):
            if not p.exists():
                if p == args.otros:
                    log.info("Otros pronósticos: no existe %s (opcional); se continúa sin ellos", p)
                    continue
                raise FileNotFoundError(f"No se encontró el archivo Excel: {p} (los Excel deben estar en la carpeta Insumos)")
            st = p.stat()
            log.info("Entrada: %s | %d bytes | modificado %s | sha256=%s", p.resolve(), st.st_size,
                     datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"), sha256(p))

        fechas = leer_fechas(args.consenso, args.fecha_consenso)   # solo hoja Fechas
        hor_ini, hor_fin = fechas["hor_ini"], fechas["hor_fin"]
        aportes = leer_aportes(args.aportes)
        ini, _ = definir_ventana(args.aportes, aportes)
        if ini >= hor_ini:
            log.warning("Inicio de gráficos (%s) >= inicio del pronóstico (%s); se usan 30 días antes",
                        ini.date(), hor_ini.date())
            ini = hor_ini - pd.Timedelta(days=30)
        cuencas = [c for c in aportes.columns if c != "Fecha"]
        ctx = cargar_drivers(args.drivers, args.pronostico)
        otros = leer_otros_pronosticos(args.otros, fechas, cuencas)

        figuras, omitidas = {}, []
        for cuenca in cuencas:
            if aportes[cuenca].notna().sum() == 0:
                log.warning("Cuenca '%s': sin ningún dato -> no se incluye", cuenca)
                omitidas.append((cuenca, "sin datos"))
                continue
            serie = aportes.loc[(aportes["Fecha"] >= ini) & (aportes["Fecha"] <= hor_fin),
                                ["Fecha", cuenca]].dropna()
            if serie.empty:
                log.warning("Cuenca '%s': sin datos en la ventana -> no se incluye", cuenca)
                omitidas.append((cuenca, "sin datos en la ventana"))
                continue
            seg, orden_vars = drivers_combinados(cuenca, ctx, ini, hor_fin)
            fig = construir_figura(cuenca, serie, seg, orden_vars, ini, hor_ini, hor_fin,
                                   otros.get(cuenca, []))
            ultimo = serie["Fecha"].max()
            figuras[cuenca] = {
                "fig": json.loads(pio.to_json(fig)),
                "sub": (f"Pronóstico {hor_ini:%d/%m/%Y} – {hor_fin:%d/%m/%Y} · aportes observados hasta "
                        f"{ultimo:%d/%m/%Y} · solo consulta")}
            log.info("Cuenca '%s': aportes=%d pts (%s -> %s) | variables=%d barras=%d | otros modelos=%s",
                     cuenca, len(serie), serie["Fecha"].min().date(), ultimo.date(), len(orden_vars), len(seg),
                     [s["nombre"] for s in otros.get(cuenca, [])])
        if not figuras:
            raise ValueError("Ninguna cuenca tiene datos para graficar")

        fecha = pd.Timestamp(args.fecha) if args.fecha else pd.Timestamp(datetime.now().date())
        args.salida.mkdir(parents=True, exist_ok=True)
        ruta = args.salida / f"Consulta_Seguimiento_Pronostico_{fecha:%Y-%m-%d}.html"
        escribir_pagina(ruta, figuras)
        log.info("RESUMEN: cuencas incluidas=%d %s | omitidas=%d %s", len(figuras), list(figuras),
                 len(omitidas), omitidas)
        log.info("FIN OK run_id=%s duración=%.2fs archivo=%s (%.1f MB) log=%s", run_id, time.time() - t0,
                 ruta.resolve(), ruta.stat().st_size / 1e6, ruta_log.resolve())
        print(f"\nGráfico de consulta: {ruta.resolve()}")
        return 0
    except Exception:
        log.exception("FIN CON ERROR run_id=%s tras %.2fs", run_id, time.time() - t0)
        return 1


if __name__ == "__main__":
    sys.exit(main())
