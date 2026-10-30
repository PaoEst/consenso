#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gráficos interactivos de aportes (% de la media histórica) por cuenca y
moduladores climáticos activos.

Entradas
    Seguimiento_Aportes.xlsx   hojas: "Aportes %MH", "parámetros"
    Seguimiento_Drivers.xlsx   hojas: "Prioridad", "Driver_Clim"

Salidas
    salida/Aportes_<cuenca>.html   un gráfico interactivo por cuenca con datos
    salida/index.html              índice con enlaces a todos los gráficos
    logs/ejecucion.log             bitácora de auditoría (se acumula entre corridas)

Los archivos Excel se leen SIEMPRE de la misma carpeta donde está este script
(sin importar desde dónde se ejecute). Salidas y logs también quedan en esa carpeta.

Uso
    python graficar_aportes.py
"""
from __future__ import annotations

import argparse
import getpass
import hashlib
import html
import logging
import platform
import re
import sys
import textwrap
import time
import unicodedata
import uuid
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# --------------------------------------------------------------------------
# Configuración
# --------------------------------------------------------------------------
BASE = Path(__file__).resolve().parent  # carpeta del proyecto (Excel, salida y logs)
ARCHIVO_APORTES = "Seguimiento_Aportes.xlsx"
ARCHIVO_DRIVERS = "Seguimiento_Drivers.xlsx"
HOJA_APORTES = "Aportes %MH"
HOJA_PARAMETROS = "parámetros"
HOJA_PRIORIDAD = "Prioridad"
HOJA_DRIVERS = "Driver_Clim"
LUGAR_GENERAL = "Colombia"
DIAS_ATRAS_POR_DEFECTO = 30  # "un mes atrás" cuando los parámetros no están activos

# Colores. Tonalidades: claro (impacto débil) -> oscuro (impacto fuerte).
AZUL = ("#bcd4f2", "#1f5fbf")
OCRE = ("#f0d9a0", "#a8620f")
NEUTRO = "#b5b5b0"
UMBRAL_SATURACION = 0.5  # |impacto-1| a partir del cual el color es el más intenso
UMBRAL_ATIPICO = 2.0     # impactos mayores se registran como advertencia en el log

C_LINEA = "#1a1a1a"
C_ARRIBA = "rgba(60,120,200,0.18)"
C_ABAJO = "rgba(205,90,60,0.18)"

log = logging.getLogger("aportes")


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def norm(txt) -> str:
    """Normaliza texto para comparar nombres: sin tildes, minúsculas, sin espacios extra."""
    if txt is None or (isinstance(txt, float) and pd.isna(txt)):
        return ""
    s = unicodedata.normalize("NFKD", str(txt))
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"\s+", " ", s).strip().casefold()


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for bloque in iter(lambda: f.read(1 << 20), b""):
            h.update(bloque)
    return h.hexdigest()


def nombre_archivo(txt: str) -> str:
    s = unicodedata.normalize("NFKD", txt)
    s = "".join(c for c in s if not unicodedata.combining(c))
    return re.sub(r"[^A-Za-z0-9_-]+", "_", s).strip("_") or "cuenca"


def hex_a_rgb(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def mezcla(c_claro: str, c_oscuro: str, t: float) -> str:
    a, b = hex_a_rgb(c_claro), hex_a_rgb(c_oscuro)
    return "#%02x%02x%02x" % tuple(round(a[i] + (b[i] - a[i]) * t) for i in range(3))


def color_impacto(impacto) -> tuple[str, str]:
    """(color, categoría) según 'Impacto Estimado'. >1 azul, <1 ocre, =1/vacío neutro."""
    if pd.isna(impacto) or abs(impacto - 1) < 1e-9:
        return NEUTRO, "neutro"
    t = min(abs(impacto - 1) / UMBRAL_SATURACION, 1.0)
    t = 0.25 + 0.75 * t  # evita tonos casi blancos
    if impacto > 1:
        return mezcla(*AZUL, t), "humedad"
    return mezcla(*OCRE, t), "seco"


def configurar_logging(dir_logs: Path, run_id: str) -> Path:
    dir_logs.mkdir(parents=True, exist_ok=True)
    ruta = dir_logs / "ejecucion.log"
    fmt = logging.Formatter(
        f"%(asctime)s | {run_id} | %(levelname)-8s | %(message)s", "%Y-%m-%d %H:%M:%S"
    )
    log.setLevel(logging.DEBUG)
    log.handlers.clear()
    fh = logging.FileHandler(ruta, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    log.addHandler(fh)
    log.addHandler(ch)
    return ruta


# --------------------------------------------------------------------------
# Lectura de datos
# --------------------------------------------------------------------------
def leer_hoja(ruta: Path, hoja: str, **kw) -> pd.DataFrame:
    df = pd.read_excel(ruta, sheet_name=hoja, **kw)
    log.info("Hoja leída: %s [%s] -> %d filas x %d columnas", ruta.name, hoja, *df.shape)
    log.debug("Columnas de '%s': %s", hoja, list(df.columns))
    return df


def leer_aportes(ruta: Path) -> pd.DataFrame:
    df = leer_hoja(ruta, HOJA_APORTES)
    df = df.rename(columns={df.columns[0]: "Fecha"})
    df.columns = [str(c).strip() for c in df.columns]
    n0 = len(df)
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    invalidas = df["Fecha"].isna().sum()
    if invalidas:
        log.warning("Aportes: %d filas con fecha inválida/vacía descartadas", invalidas)
    df = df.dropna(subset=["Fecha"]).sort_values("Fecha").reset_index(drop=True)
    duplicadas = df["Fecha"].duplicated().sum()
    if duplicadas:
        log.warning("Aportes: %d fechas duplicadas; se conserva la última de cada una", duplicadas)
        df = df.drop_duplicates("Fecha", keep="last")
    df["Fecha"] = df["Fecha"].dt.normalize()
    for c in df.columns[1:]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    log.info("Aportes: %d filas válidas de %d; rango %s -> %s",
             len(df), n0, df["Fecha"].min().date(), df["Fecha"].max().date())
    return df


def _leer_fecha(valor):
    """Fecha de Excel o texto dd/mm/aaaa -> Timestamp normalizado, o None si no se puede leer."""
    if valor is None or (not isinstance(valor, str) and pd.isna(valor)):
        return None
    f = pd.to_datetime(valor, errors="coerce", dayfirst=isinstance(valor, str))
    return None if pd.isna(f) else f.normalize()


def definir_ventana(ruta: Path, aportes: pd.DataFrame) -> tuple[pd.Timestamp, pd.Timestamp]:
    """Fechas de la hoja 'parámetros' (col B) si su columna C = 'si'; si no, último mes de datos.

    Si hay fechas activas pero son inconsistentes (ilegibles, inicio > fin o sin datos de
    aportes en el rango) se usa el último mes de información y se deja una advertencia.
    """
    datos = aportes.dropna(subset=list(aportes.columns[1:]), how="all")
    ultima = (datos if len(datos) else aportes)["Fecha"].max()
    ini_def = ultima - pd.Timedelta(days=DIAS_ATRAS_POR_DEFECTO)
    fin_def = ultima
    log.info("Ventana por defecto: %s -> %s (última fecha con datos, un mes atrás)",
             ini_def.date(), fin_def.date())

    try:
        p = pd.read_excel(ruta, sheet_name=HOJA_PARAMETROS, header=None)
    except Exception as e:  # hoja ausente o ilegible
        log.warning("Parámetros: no se pudo leer la hoja '%s' (%s); se usa el último mes",
                    HOJA_PARAMETROS, e)
        return ini_def, fin_def
    log.debug("Hoja parámetros (crudo):\n%s", p.dropna(how="all").to_string())

    # Fila 2 = inicio, fila 3 = fin (formato: A etiqueta, B fecha, C activa)
    fechas, motivos = {}, []
    for clave, i in {"inicio": 1, "fin": 2}.items():
        try:
            valor, activa = p.iloc[i, 1], p.iloc[i, 2]
        except IndexError:
            log.warning("Parámetros: no se encontró la fila de fecha de %s", clave)
            continue
        activa_ok = norm(activa) in ("si", "s", "yes", "true", "1")
        if not activa_ok:
            log.info("Parámetro fecha %s: valor=%r activa=%r -> ignorada (no activa)", clave, valor, activa)
            continue
        fecha = _leer_fecha(valor)
        if fecha is None:
            motivos.append(f"fecha de {clave} activa pero ilegible (valor={valor!r})")
            log.info("Parámetro fecha %s: valor=%r activa=%r -> ILEGIBLE", clave, valor, activa)
        else:
            fechas[clave] = fecha
            log.info("Parámetro fecha %s: valor=%s activa=%r -> USADA", clave, fecha.date(), activa)

    ini = fechas.get("inicio", ini_def)
    fin = fechas.get("fin", fin_def)
    if not motivos and fechas:
        if ini > fin:
            motivos.append(f"inicio ({ini.date()}) es posterior a fin ({fin.date()})")
        elif not ((aportes["Fecha"] >= ini) & (aportes["Fecha"] <= fin)).any():
            motivos.append(f"no hay datos de aportes entre {ini.date()} y {fin.date()}")
    if motivos:
        log.warning("Parámetros de fechas inconsistentes (%s); se usa el último mes de información: %s -> %s",
                    "; ".join(motivos), ini_def.date(), fin_def.date())
        ini, fin = ini_def, fin_def
    log.info("Ventana final de gráficos: %s -> %s", ini.date(), fin.date())
    return ini, fin


def leer_prioridad(ruta: Path, hoja: str = HOJA_PRIORIDAD) -> pd.DataFrame:
    df = leer_hoja(ruta, hoja)
    df = df.iloc[:, :2].copy()
    df.columns = ["Variable", "Remover"]
    df = df.dropna(subset=["Variable"]).copy()
    df["clave"] = df["Variable"].map(norm)
    df["Variable"] = df["Variable"].astype(str).str.strip()
    df["remover"] = df["Remover"].map(norm)  # 'si' | 'no' | ''
    log.info("Prioridad: %d variables; remover=si: %s", len(df),
             df.loc[df.remover == "si", "Variable"].tolist())
    return df


def leer_drivers(ruta: Path, prioridad: pd.DataFrame) -> pd.DataFrame:
    df = leer_hoja(ruta, HOJA_DRIVERS)
    df.columns = [str(c).strip() for c in df.columns]
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce").dt.normalize()
    sin_fecha = df["Fecha"].isna().sum()
    if sin_fecha:
        log.warning("Driver_Clim: %d filas sin fecha válida descartadas", sin_fecha)
    df = df.dropna(subset=["Fecha"]).copy()
    df["clave"] = df["Variable"].map(norm)
    df["lugar_n"] = df["Lugar"].map(norm)
    df["incluir_n"] = df["Incluir"].map(norm)
    df["presente_n"] = df["Presente"].map(norm)
    df["Impacto Estimado"] = pd.to_numeric(df["Impacto Estimado"], errors="coerce")
    df["Nota"] = df["Nota"].where(df["Nota"].notna(), "").astype(str).str.strip()
    df["comentarios"] = df["comentarios"].where(df["comentarios"].notna(), "").astype(str).str.strip()

    desconocidas = sorted(set(df["clave"]) - set(prioridad["clave"]) - {""})
    if desconocidas:
        log.warning("Variables en Driver_Clim ausentes de la hoja Prioridad (van al final): %s",
                    desconocidas)
    ns = (df["presente_n"] == "ns").sum()
    if ns:
        log.warning("Driver_Clim: %d filas con Presente='ns' (se tratan como no presentes)", ns)
    for _, r in df[df["Impacto Estimado"] > UMBRAL_ATIPICO].iterrows():
        log.warning("Impacto atípico (>%s): %s | %s | %s -> %s",
                    UMBRAL_ATIPICO, r["Fecha"].date(), r["Variable"], r["Lugar"], r["Impacto Estimado"])
    log.info("Driver_Clim: rango %s -> %s; lugares: %s", df["Fecha"].min().date(),
             df["Fecha"].max().date(), sorted(df["Lugar"].dropna().astype(str).unique()))
    return df


# --------------------------------------------------------------------------
# Lógica de drivers por cuenca
# --------------------------------------------------------------------------
def drivers_de_cuenca(drv: pd.DataFrame, prioridad: pd.DataFrame, cuenca: str,
                      ini: pd.Timestamp, fin: pd.Timestamp) -> pd.DataFrame:
    """Filas diarias a graficar para una cuenca (una por variable y día)."""
    d = drv[(drv["Fecha"] >= ini) & (drv["Fecha"] <= fin)]
    esp = d[d["lugar_n"] == norm(cuenca)]
    gen = d[d["lugar_n"] == norm(LUGAR_GENERAL)]

    # La definición específica de la cuenca tiene prioridad sobre Colombia, por (variable, día).
    claves_esp = set(zip(esp["clave"], esp["Fecha"]))
    mask = [(k, f) not in claves_esp for k, f in zip(gen["clave"], gen["Fecha"])]
    gen_ok = gen[mask]
    sel = pd.concat([esp, gen_ok], ignore_index=True)
    log.debug("%s: filas específicas=%d, Colombia=%d (reemplazadas por la cuenca=%d)",
              cuenca, len(esp), len(gen), len(gen) - len(gen_ok))

    # Regla 'Remover' de la hoja Prioridad: si -> fuera; no -> dentro; vacío -> decide 'Incluir' diario
    rem = dict(zip(prioridad["clave"], prioridad["remover"]))
    antes = len(sel)
    if len(sel):
        def conservar(r):
            regla = rem.get(r["clave"], "")
            if regla == "si":
                return False
            if regla == "no":
                return True
            return r["incluir_n"] == "si"
        sel = sel[sel.apply(conservar, axis=1)]
    log.debug("%s: filas tras regla Remover/Incluir: %d de %d", cuenca, len(sel), antes)

    sel = sel[sel["presente_n"] == "si"].copy()
    log.debug("%s: filas con Presente='si': %d", cuenca, len(sel))
    return sel


def segmentos(sel: pd.DataFrame) -> pd.DataFrame:
    """Unifica días consecutivos con misma variable, nota e impacto en una sola barra."""
    cols = ["clave", "Variable", "ini", "fin", "nota", "impacto", "lugar", "coment"]
    if sel.empty:
        return pd.DataFrame(columns=cols)
    sel = sel.sort_values(["clave", "Fecha"])
    out = []
    for clave, g in sel.groupby("clave", sort=False):
        actual = None
        for _, r in g.iterrows():
            imp = r["Impacto Estimado"]
            llave = (r["Nota"], None if pd.isna(imp) else float(imp), r["lugar_n"])
            if (actual and actual["llave"] == llave
                    and r["Fecha"] - actual["fin"] == pd.Timedelta(days=1)):
                actual["fin"] = r["Fecha"]
                if r["comentarios"]:
                    actual["coment"].append((r["Fecha"], r["comentarios"]))
            else:
                if actual:
                    out.append(actual)
                actual = {"llave": llave, "clave": clave, "Variable": str(r["Variable"]).strip(),
                          "ini": r["Fecha"], "fin": r["Fecha"], "nota": r["Nota"],
                          "impacto": imp, "lugar": r["Lugar"],
                          "coment": [(r["Fecha"], r["comentarios"])] if r["comentarios"] else []}
        out.append(actual)
    return pd.DataFrame(out).drop(columns="llave")


# --------------------------------------------------------------------------
# Gráfico
# --------------------------------------------------------------------------
def relleno_media(x, y, ref=100.0):
    """Serie con puntos de cruce interpolados para sombrear por encima/debajo de la media."""
    xs, ys = [], []
    for i in range(len(x)):
        if i > 0 and (y[i - 1] - ref) * (y[i] - ref) < 0:
            f = (ref - y[i - 1]) / (y[i] - y[i - 1])
            xs.append(x[i - 1] + (x[i] - x[i - 1]) * f)
            ys.append(ref)
        xs.append(x[i])
        ys.append(y[i])
    return xs, [max(v, ref) for v in ys], [min(v, ref) for v in ys]


def texto_comentarios(coment: list) -> str:
    """Comentarios de una barra para el hover: uno solo si es igual todos los días; si no, por día."""
    if not coment:
        return ""
    unicos = list(dict.fromkeys(c for _, c in coment))
    envolver = lambda t: "<br>".join(textwrap.wrap(html.escape(t), 60))
    if len(unicos) == 1:
        return "<br>Comentarios: " + envolver(unicos[0])
    return "<br>Comentarios:<br>" + "<br>".join(envolver(f"{f:%d %b}: {c}") for f, c in coment)


def construir_figura(cuenca: str, serie: pd.DataFrame, seg: pd.DataFrame,
                     orden_vars: list[str], ini: pd.Timestamp, fin: pd.Timestamp) -> go.Figure:
    n = max(len(orden_vars), 1)
    alto_drv = 28 * n + 40                   # píxeles del panel de drivers
    alto_tot = int(380 + alto_drv + 110)
    frac = min(0.62, max(0.25, alto_drv / alto_tot))
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.04,
                        row_heights=[1 - frac, frac])

    # --- Serie de aportes ---
    # Cada dato diario se dibuja a mediodía: queda centrado en la columna de su día,
    # que es la que ocupan las barras de drivers ([día 00:00, día+1 00:00)).
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
                  annotation_text="Media histórica (100 %)", annotation_position="top right",
                  annotation_font=dict(size=11, color="#444"))

    # --- Barras de drivers (polígonos rellenos sobre eje de fechas compartido) ---
    pos = {v: i for i, v in enumerate(orden_vars)}
    h = 0.36
    cx, cy, ct = [], [], []
    for _, r in seg.iterrows():
        color = color_impacto(r["impacto"])[0]
        imp = "s/d" if pd.isna(r["impacto"]) else f"{r['impacto']:g}"
        dias = (r["fin"] - r["ini"]).days + 1
        fin_b = r["fin"] + pd.Timedelta(days=1)
        yc = pos[r["Variable"]]
        hover = (f"<b>{html.escape(r['Variable'])}</b><br>"
                 f"{r['ini']:%d %b} – {r['fin']:%d %b %Y} ({dias} d)<br>"
                 f"Nota: {html.escape(r['nota']) or '—'}<br>"
                 f"Impacto estimado: {imp}<br>Lugar: {html.escape(str(r['lugar']))}"
                 + texto_comentarios(r["coment"]))
        fig.add_trace(go.Scatter(
            x=[r["ini"], fin_b, fin_b, r["ini"], r["ini"]],
            y=[yc - h, yc - h, yc + h, yc + h, yc - h], mode="lines", fill="toself",
            fillcolor=color, line=dict(color="white", width=1.5), hoveron="fills",
            text=hover, hoverinfo="text", showlegend=False), row=2, col=1)
        if r["nota"]:
            cx.append(r["ini"] + (fin_b - r["ini"]) / 2)
            cy.append(yc)
            ct.append(r["nota"])
    if ct:
        fig.add_trace(go.Scatter(x=cx, y=cy, mode="text", text=ct, hoverinfo="skip",
                                 textfont=dict(color="white", size=11), showlegend=False),
                      row=2, col=1)

    # Leyenda manual
    for nombre, col in (("Favorece humedad", mezcla(*AZUL, 0.8)),
                        ("Favorece condiciones secas", mezcla(*OCRE, 0.8)),
                        ("Presente, efecto neutro/incierto", NEUTRO)):
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers", name=nombre,
                                 marker=dict(color=col, size=12, symbol="square")),
                      row=2, col=1)

    rango = [ini, fin + pd.Timedelta(days=1)]
    fig.update_yaxes(title_text="Aportes (% de la media)", row=1, col=1,
                     gridcolor="#e6e6e6", zeroline=False)
    fig.update_yaxes(tickmode="array", tickvals=list(range(len(orden_vars))),
                     ticktext=orden_vars, range=[len(orden_vars) - 0.4, -0.6],
                     row=2, col=1, gridcolor="#efefef", tickfont=dict(size=11), automargin=True)
    for fila in (1, 2):
        fig.update_xaxes(range=rango, matches="x" if fila == 2 else None, gridcolor="#e6e6e6", showspikes=True, spikemode="across",
                         spikethickness=1, spikecolor="#888", tickformat="%d %b\n%Y",
                         row=fila, col=1)
    fig.update_layout(
        title=dict(text=f"Seguimiento de Aportes {cuenca} y moduladores climáticos",
                   x=0.5, xanchor="center", font=dict(size=18)),
        hoverlabel=dict(align="left"),
        height=alto_tot, template="plotly_white",
        hovermode="closest", margin=dict(l=220, r=30, t=70, b=60),
        legend=dict(orientation="h", x=0.5, xanchor="center", y=-0.1),
        plot_bgcolor="#fbfbf9")
    return fig


# --------------------------------------------------------------------------
# Programa principal
# --------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--aportes", default=BASE / ARCHIVO_APORTES, type=Path)
    ap.add_argument("--drivers", default=BASE / ARCHIVO_DRIVERS, type=Path)
    ap.add_argument("--salida", default=BASE / "salida", type=Path)
    ap.add_argument("--logs", default=BASE / "logs", type=Path)
    args = ap.parse_args()

    run_id = uuid.uuid4().hex[:8]
    t0 = time.time()
    ruta_log = configurar_logging(args.logs, run_id)
    log.info("=" * 70)
    log.info("INICIO ejecución run_id=%s", run_id)
    log.info("Usuario=%s host=%s SO=%s python=%s pandas=%s plotly=%s", getpass.getuser(),
             platform.node(), platform.platform(), platform.python_version(),
             pd.__version__, plotly.__version__)
    log.info("Argumentos: %s", {k: str(v) for k, v in vars(args).items()})

    try:
        for p in (args.aportes, args.drivers):
            if not p.exists():
                raise FileNotFoundError(f"No se encontró el archivo Excel: {p} (debe estar en la carpeta del script)")
            st = p.stat()
            log.info("Entrada: %s | %d bytes | modificado %s | sha256=%s", p.resolve(), st.st_size,
                     datetime.fromtimestamp(st.st_mtime).isoformat(timespec="seconds"), sha256(p))

        aportes = leer_aportes(args.aportes)
        ini, fin = definir_ventana(args.aportes, aportes)
        prioridad = leer_prioridad(args.drivers)
        drivers = leer_drivers(args.drivers, prioridad)

        cuencas = [c for c in aportes.columns if c != "Fecha"]
        log.info("Columnas de cuenca detectadas (%d): %s", len(cuencas), cuencas)

        # Variables en orden de prioridad; las no listadas van al final
        conocidas = set(prioridad["clave"])
        extra = [str(v).strip() for k, v in drivers.drop_duplicates("clave")[["clave", "Variable"]].values
                 if k and k not in conocidas]
        orden_global = list(prioridad["Variable"]) + extra
        clave_a_nombre = {norm(v): v for v in orden_global}

        args.salida.mkdir(parents=True, exist_ok=True)
        generados, omitidos = [], []

        for cuenca in cuencas:
            if aportes[cuenca].notna().sum() == 0:
                log.warning("Cuenca '%s': sin ningún dato -> no se genera gráfico", cuenca)
                omitidos.append((cuenca, "sin datos"))
                continue
            serie = aportes.loc[(aportes["Fecha"] >= ini) & (aportes["Fecha"] <= fin),
                                ["Fecha", cuenca]].dropna()
            if serie.empty:
                log.warning("Cuenca '%s': sin datos en la ventana %s -> %s; no se genera gráfico",
                            cuenca, ini.date(), fin.date())
                omitidos.append((cuenca, "sin datos en la ventana"))
                continue

            sel = drivers_de_cuenca(drivers, prioridad, cuenca, ini, fin)
            seg = segmentos(sel)
            if len(seg):
                seg["Variable"] = seg["clave"].map(lambda k: clave_a_nombre.get(k, k))
            presentes = set(seg["clave"])
            orden_vars = [v for v in orden_global if norm(v) in presentes]
            vacias = [v for v in orden_global if norm(v) not in presentes]
            n_esp = int((sel["lugar_n"] == norm(cuenca)).sum()) if len(sel) else 0

            fig = construir_figura(cuenca, serie, seg, orden_vars, ini, fin)
            ruta = args.salida / f"Aportes_{nombre_archivo(cuenca)}.html"
            fig.write_html(ruta, include_plotlyjs="cdn", full_html=True,
                           config={"displaylogo": False})
            generados.append((cuenca, ruta.name))
            log.info("Cuenca '%s': gráfico -> %s | puntos=%d (%s -> %s) media=%.1f min=%.1f max=%.1f",
                     cuenca, ruta, len(serie), serie["Fecha"].min().date(), serie["Fecha"].max().date(),
                     serie[cuenca].mean(), serie[cuenca].min(), serie[cuenca].max())
            log.info("Cuenca '%s': variables graficadas=%d filas-día=%d (específicas de cuenca=%d) barras=%d",
                     cuenca, len(orden_vars), len(sel), n_esp, len(seg))
            log.debug("Cuenca '%s': variables mostradas=%s; sin barras en ventana=%s",
                      cuenca, orden_vars, vacias)
            for _, r in seg.iterrows():
                log.debug("  barra | %s | %s -> %s | nota=%r | impacto=%s | lugar=%s", r["Variable"],
                          r["ini"].date(), r["fin"].date(), r["nota"], r["impacto"], r["lugar"])

        items = "\n".join(f'<li><a href="{html.escape(f)}">{html.escape(c)}</a></li>' for c, f in generados)
        (args.salida / "index.html").write_text(
            "<!doctype html><meta charset='utf-8'><title>Aportes por cuenca</title>"
            "<body style='font-family:sans-serif;margin:2rem'>"
            f"<h2>Aportes por cuenca y moduladores climáticos</h2><p>{ini:%d/%m/%Y} – {fin:%d/%m/%Y}</p>"
            f"<ul>{items}</ul></body>", encoding="utf-8")

        log.info("RESUMEN: generados=%d %s | omitidos=%d %s", len(generados),
                 [c for c, _ in generados], len(omitidos), omitidos)
        log.info("FIN OK run_id=%s duración=%.2fs log=%s", run_id, time.time() - t0, ruta_log.resolve())
        return 0
    except Exception:
        log.exception("FIN CON ERROR run_id=%s tras %.2fs", run_id, time.time() - t0)
        return 1


if __name__ == "__main__":
    sys.exit(main())
