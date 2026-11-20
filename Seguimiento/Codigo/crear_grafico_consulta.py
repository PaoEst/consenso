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

ARCHIVO AUTOCONTENIDO: no necesita ningún otro programa de la carpeta (solo pandas, plotly y openpyxl).
Es una copia independiente de la lógica que comparte con los demás programas del proyecto; si esa
lógica cambia allá, actualizarla aquí.

Uso:  python crear_grafico_consulta.py [--fecha AAAA-MM-DD] [--fecha-consenso AAAA-MM-DD]
"""
from __future__ import annotations

import argparse
import difflib
import getpass
import hashlib
import html
import json
import logging
import platform
import re
import sys
import textwrap
import time
import unicodedata
import uuid
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
import plotly
import plotly.graph_objects as go
import plotly.io as pio
import plotly.offline
from openpyxl.utils import get_column_letter
from plotly.subplots import make_subplots

# --------------------------------------------------------------------------
# Rutas (relativas a este archivo): Codigo/ (este archivo y logs), ../Insumos y ../Salidas
# --------------------------------------------------------------------------
BASE = Path(__file__).resolve().parent
INSUMOS = BASE.parent / "Insumos"
SALIDAS = BASE.parent / "Salidas"
MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]

# --------------------------------------------------------------------------
# Constantes (copiadas de graficar_aportes.py y consenso_comun.py)
# --------------------------------------------------------------------------

VERSION_CODIGO = "2026-11-20.3"                 # se actualiza en cada publicación (aparece en el pie de las páginas)
ARCHIVO_APORTES = "Seguimiento_Aportes.xlsx"
ARCHIVO_DRIVERS = "Seguimiento_Drivers.xlsx"
HOJA_APORTES = "Aportes %MH"
HOJA_PARAMETROS = "parámetros"
HOJA_PRIORIDAD = "Prioridad"
HOJA_DRIVERS = "Driver_Clim"
LUGAR_GENERAL = "Colombia"
DIAS_ATRAS_POR_DEFECTO = 30  # "un mes atrás" cuando los parámetros no están activos
AZUL = ("#bcd4f2", "#1f5fbf")
OCRE = ("#f0d9a0", "#a8620f")
NEUTRO = "#b5b5b0"
UMBRAL_SATURACION = 0.5  # |impacto-1| a partir del cual el color es el más intenso
UMBRAL_ATIPICO = 2.0     # impactos mayores se registran como advertencia en el log
C_LINEA = "#1a1a1a"
C_ARRIBA = "rgba(60,120,200,0.18)"
C_ABAJO = "rgba(205,90,60,0.18)"
log = logging.getLogger("aportes")
ARCHIVO_CONSENSO = "Consenso.xlsx"
ARCHIVO_PRONOSTICO = "Pronostico_Drivers.xlsx"
HOJA_FECHAS = "Fechas"
N_SEMANAS = 5
_TERMINO = re.compile(r"([+-]?)([A-Za-z]{1,2}[0-9]+|[0-9]+(?:\.[0-9]+)?)")
ESTILO_NAV = ("<style>.nav{display:flex;flex-wrap:wrap;gap:18px;align-items:center;padding:8px 20px;"
              "font-size:13px;border-bottom:1px solid var(--brd)} .nav a{color:var(--azul);text-decoration:none;"
              "font-weight:600} .nav select{font-size:13px;padding:3px 6px;margin-left:6px;border-radius:5px;"
              "border:1px solid var(--brd);background:var(--bg);color:var(--txt)}</style>")
ARCHIVO_OTROS = "Otros_pronosticos.xlsx"
PALETA_OTROS = ["#c2185b", "#00897b", "#7b1fa2", "#6d4c41", "#827717", "#d32f2f", "#5d4037",
                "#00838f", "#ef6c00", "#455a64"]
_ESTILO_FREQ = {"mensual": dict(dash="dash", width=3), "semanal": dict(dash="dot", width=3),
                "diario": dict(dash="solid", width=2)}

# --------------------------------------------------------------------------
# Funciones (copiadas de graficar_aportes.py)
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


def texto_comentarios(coment: list, ini=None, fin=None) -> str:
    """Comentarios de una barra para el hover. Los días consecutivos con el mismo comentario
    se unifican: 'dd Mmm – dd Mmm: texto'. Si un solo comentario cubre toda la barra se
    muestra sin fechas."""
    if not coment:
        return ""
    grupos = []  # [fecha_ini, fecha_fin, texto]
    for f, c in sorted(coment, key=lambda x: x[0]):
        if grupos and grupos[-1][2] == c and f - grupos[-1][1] <= pd.Timedelta(days=1):
            grupos[-1][1] = f
        else:
            grupos.append([f, f, c])
    envolver = lambda t: "<br>".join(textwrap.wrap(html.escape(t), 55))
    if len(grupos) == 1 and (ini is None or (grupos[0][0] <= ini and grupos[0][1] >= fin)):
        return "<br>Comentarios: " + envolver(grupos[0][2])
    lineas = []
    for gi, gf, t in grupos:
        rango = f"{gi:%d %b}" if gi == gf else f"{gi:%d %b} – {gf:%d %b}"
        lineas.append(envolver(f"{rango}: {t}"))
    return "<br>Comentarios:<br>" + "<br>".join(lineas)



# --------------------------------------------------------------------------
# Funciones (copiadas de consenso_comun.py)
# --------------------------------------------------------------------------

def _fecha(v):
    f = pd.to_datetime(v, errors="coerce", dayfirst=isinstance(v, str))
    return None if pd.isna(f) else f.normalize()


def _valor_celda(ws_f, ws_v, coord: str, profundidad: int = 0):
    """Valor de una celda; si no tiene valor guardado (Excel no recalculado, p. ej. tras guardar
    con openpyxl) evalúa fórmulas simples del tipo =+C2, =+E2+6, =F2-1."""
    v = ws_v[coord].value
    if v is not None:
        return v
    f = ws_f[coord].value
    if not (isinstance(f, str) and f.startswith("=")) or profundidad > 20:
        return f
    expr = f[1:].replace(" ", "")
    total, pos, fecha = 0.0, 0, False
    while pos < len(expr):
        m = _TERMINO.match(expr, pos)
        if not m:
            raise ValueError(f"Fórmula no soportada en {coord}: {f}")
        signo = -1 if m.group(1) == "-" else 1
        tok = m.group(2)
        if tok[0].isalpha():
            ref = _valor_celda(ws_f, ws_v, tok.upper(), profundidad + 1)
            if isinstance(ref, (datetime, date)):
                fecha = True
                ref = (ref.date() if isinstance(ref, datetime) else ref).toordinal()
            if ref is None:
                return None
            total += signo * float(ref)
        else:
            total += signo * float(tok)
        pos = m.end()
    return datetime.fromordinal(int(round(total))) if fecha else total


def _filas_fechas(ws_f, ws_v) -> list[int]:
    """Filas registradas de la hoja Fechas (cualquier celda de A a N con valor o fórmula), en orden.
    La última de esta lista es siempre 'la última fila registrada'."""
    filas = []
    for r in range(2, ws_f.max_row + 1):
        if any(ws_v.cell(r, c).value not in (None, "") or ws_f.cell(r, c).value not in (None, "")
               for c in range(1, 5 + 2 * N_SEMANAS)):
            filas.append(r)
    return filas


def _parsear_fila_fechas(ws_f, ws_v, fila: int, silencioso: bool = False, estricto: bool = True) -> dict:
    """Fecha del consenso, horizonte y semanas S1..S5 (inicio/fin) de una fila de la hoja Fechas.

    `semanas`: semanas votables (dentro del horizonte; las usa la votación). `semanas_dibujo`: TODAS las
    semanas con inicio y fin llenos (inicio <= fin), tal como están en la fila (las usa el gráfico de consulta).
    Con estricto=False, si faltan las fechas del horizonte o del consenso se derivan de las semanas."""
    aviso = log.debug if silencioso else log.warning
    info = log.debug if silencioso else log.info
    v = [_valor_celda(ws_f, ws_v, f"{get_column_letter(c)}{fila}") for c in range(1, 5 + 2 * N_SEMANAS)]
    out = {"fila": fila, "seguimiento_ini": _fecha(v[0]), "consenso": _fecha(v[1]),
           "hor_ini": _fecha(v[2]), "hor_fin": _fecha(v[3]), "semanas": [], "semanas_dibujo": []}
    for k in range(N_SEMANAS):
        ini, fin = _fecha(v[4 + 2 * k]), _fecha(v[5 + 2 * k])
        if ini is not None and fin is not None and ini <= fin:
            out["semanas_dibujo"].append({"n": k + 1, "ini": ini, "fin": fin, "dias": (fin - ini).days + 1})
        elif ini is not None or fin is not None:
            aviso("Fechas fila %d: S%d con inicio/fin incompletos o invertidos (%s, %s); no se dibuja", fila,
                  k + 1, ini.date() if ini is not None else None, fin.date() if fin is not None else None)
    sd = out["semanas_dibujo"]
    if not estricto:
        if out["hor_ini"] is None and sd:
            out["hor_ini"] = min(w["ini"] for w in sd)
            aviso("Fechas fila %d: sin inicio del horizonte; se usa el inicio de la primera semana (%s)", fila,
                  out["hor_ini"].date())
        if out["hor_fin"] is None and sd:
            out["hor_fin"] = max(w["fin"] for w in sd)
            aviso("Fechas fila %d: sin fin del horizonte; se usa el fin de la última semana (%s)", fila,
                  out["hor_fin"].date())
        if out["consenso"] is None and out["hor_ini"] is not None:
            out["consenso"] = out["hor_ini"]
            aviso("Fechas fila %d: sin fecha de consenso; se usa el inicio del horizonte", fila)
    for faltante in ("consenso", "hor_ini", "hor_fin"):
        if out[faltante] is None:
            raise ValueError(f"Hoja '{HOJA_FECHAS}' fila {fila}: falta o es ilegible la fecha '{faltante}'")
    if out["hor_ini"] > out["hor_fin"]:
        raise ValueError(f"Hoja Fechas fila {fila}: el inicio del horizonte es posterior a su fin")
    for k in range(N_SEMANAS):
        ini, fin = _fecha(v[4 + 2 * k]), _fecha(v[5 + 2 * k])
        if ini is None or fin is None:
            info("Fechas fila %d: S%d sin fechas completas; no se usa", fila, k + 1)
            continue
        if ini > fin:   # p. ej. cadena de fórmulas que se pasa del fin del horizonte: la semana no existe
            aviso("Fechas fila %d: S%d tiene inicio (%s) posterior al fin (%s); no se usa", fila, k + 1,
                  ini.date(), fin.date())
            continue
        if fin < out["hor_ini"] or ini > out["hor_fin"]:
            aviso("Fechas fila %d: S%d (%s -> %s) queda totalmente fuera del horizonte (%s -> %s); no se vota",
                  fila, k + 1, ini.date(), fin.date(), out["hor_ini"].date(), out["hor_fin"].date())
            continue
        if ini < out["hor_ini"] or fin > out["hor_fin"]:
            aviso("Fechas fila %d: S%d (%s -> %s) sobresale del horizonte (%s -> %s)", fila, k + 1,
                  ini.date(), fin.date(), out["hor_ini"].date(), out["hor_fin"].date())
        # días de la semana dentro del MES del inicio del horizonte (resto del mes y cierre no mezclan meses)
        m_ini = out["hor_ini"].replace(day=1)
        m_fin = m_ini + pd.offsets.MonthEnd(0)
        ini_m, fin_m = max(ini, m_ini), min(fin, m_fin)
        dias_mes = max((fin_m - ini_m).days + 1, 0)
        out["semanas"].append({"n": k + 1, "ini": ini, "fin": fin, "dias": (fin - ini).days + 1,
                               "dias_mes": dias_mes})
    if not out["semanas"] and estricto:
        raise ValueError(f"Hoja Fechas fila {fila}: no hay ninguna semana dentro del horizonte")
    return out


def _abrir_fechas(ruta: Path):
    wb_f = openpyxl.load_workbook(ruta)
    wb_v = openpyxl.load_workbook(ruta, data_only=True)
    if HOJA_FECHAS not in wb_f.sheetnames:
        raise ValueError(f"{ruta.name} no tiene la hoja '{HOJA_FECHAS}'")
    return wb_f[HOJA_FECHAS], wb_v[HOJA_FECHAS]


def leer_fechas(ruta: Path, fecha=None, estricto: bool = True) -> dict:
    """Datos del consenso en curso. Cada consenso es una FILA de la hoja Fechas (historial);
    por defecto se usa la última fila con datos. Con `fecha` (AAAA-MM-DD) se usa la fila cuya
    'Fecha de realización del consenso' coincide (la última si hay repetidas). `estricto=False` (gráfico de
    consulta) tolera que falten las fechas del horizonte o del consenso derivándolas de las semanas."""
    ws_f, ws_v = _abrir_fechas(ruta)
    filas = _filas_fechas(ws_f, ws_v)
    if not filas:
        raise ValueError(f"La hoja '{HOJA_FECHAS}' no tiene ninguna fila de consenso")
    if fecha is None:
        fila = filas[-1]
        motivo = "última fila con datos"
    else:
        objetivo = pd.Timestamp(fecha).normalize()
        cand = []
        for r in filas:
            try:
                if _fecha(_valor_celda(ws_f, ws_v, f"B{r}")) == objetivo:
                    cand.append(r)
            except ValueError:
                continue
        if not cand:
            raise ValueError(f"La hoja '{HOJA_FECHAS}' no tiene un consenso con fecha {objetivo.date()}")
        fila = cand[-1]
        motivo = f"fila con fecha de consenso {objetivo.date()}"
    out = _parsear_fila_fechas(ws_f, ws_v, fila, estricto=estricto)
    log.info("Fechas: usando el consenso %d de %d (fila %d de la hoja; %s): consenso=%s horizonte=%s -> %s semanas=%s",
             filas.index(fila) + 1, len(filas), fila, motivo, out["consenso"].date(), out["hor_ini"].date(), out["hor_fin"].date(),
             [(w["n"], str(w["ini"].date()), str(w["fin"].date()), w["dias"], w["dias_mes"]) for w in out["semanas"]])
    log.info("Fechas: semanas a dibujar (inicio y fin llenos): %s",
             [(w["n"], str(w["ini"].date()), str(w["fin"].date())) for w in out["semanas_dibujo"]])
    return out


def plotly_js() -> str:
    """Código de plotly.js para incrustarlo en el HTML (funciona sin internet)."""
    return plotly.offline.get_plotlyjs()


def hoja_prioridad(ruta: Path) -> str:
    """Nombre de la hoja de prioridad (tolera 'Prioridad' y el typo 'Priodidad')."""
    for h in pd.ExcelFile(ruta).sheet_names:
        if norm(h).startswith("prio"):
            return h
    raise ValueError(f"{ruta.name}: no se encontró una hoja de prioridad")


def cargar_drivers(ruta_seg: Path, ruta_pro: Path) -> dict:
    """Lee seguimiento y pronóstico y define el orden de las filas del panel de drivers.

    Orden: prioridad de Seguimiento_Drivers; después las variables que solo existen en el
    pronóstico (prioridad de Pronostico_Drivers); al final cualquier otra no listada."""
    pri_seg = leer_prioridad(ruta_seg)
    drv_seg = leer_drivers(ruta_seg, pri_seg)
    pri_pro = leer_prioridad(ruta_pro, hoja_prioridad(ruta_pro))
    drv_pro = leer_drivers(ruta_pro, pri_pro)
    orden = list(pri_seg["Variable"])
    conocidas = set(pri_seg["clave"])
    for k, v in zip(pri_pro["clave"], pri_pro["Variable"]):
        if k not in conocidas:
            orden.append(v)
            conocidas.add(k)
    for df in (drv_seg, drv_pro):
        for k, v in df.drop_duplicates("clave")[["clave", "Variable"]].values:
            if k and k not in conocidas:
                orden.append(str(v).strip())
                conocidas.add(k)
    nombre_de = {}
    for v in orden:
        nombre_de.setdefault(norm(v), v)
    log.info("Variables solo del pronóstico (al final): %s",
             [v for v in orden if norm(v) not in set(pri_seg["clave"])])
    return {"pri_seg": pri_seg, "drv_seg": drv_seg, "pri_pro": pri_pro, "drv_pro": drv_pro,
            "orden": orden, "nombre_de": nombre_de}


def drivers_combinados(cuenca, ctx: dict, ini, fin):
    """Barras de seguimiento y pronóstico de una cuenca + lista ordenada de variables con barras."""
    sel_seg = drivers_de_cuenca(ctx["drv_seg"], ctx["pri_seg"], cuenca, ini, fin)
    sel_pro = drivers_de_cuenca(ctx["drv_pro"], ctx["pri_pro"], cuenca, ini, fin)
    if len(sel_seg) and len(sel_pro):  # si un día de una variable está en ambos, gana el seguimiento
        ya = set(zip(sel_seg["clave"], sel_seg["Fecha"]))
        dup = [(k, f) in ya for k, f in zip(sel_pro["clave"], sel_pro["Fecha"])]
        if any(dup):
            log.warning("%s: %d filas de pronóstico coinciden con días de seguimiento; se omiten",
                        cuenca, sum(dup))
        sel_pro = sel_pro[[not d for d in dup]]
    partes = []
    for sel, origen in ((sel_seg, "Seguimiento"), (sel_pro, "Pronóstico")):
        s = segmentos(sel)
        s["origen"] = origen
        partes.append(s)
    seg = pd.concat(partes, ignore_index=True)
    if len(seg):
        seg["Variable"] = seg["clave"].map(lambda k: ctx["nombre_de"].get(k, k))
    presentes = set(seg["clave"])
    return seg, [v for v in ctx["orden"] if norm(v) in presentes]


def hover_barra(r) -> str:
    imp = "s/d" if pd.isna(r["impacto"]) else f"{r['impacto']:g}"
    dias = (r["fin"] - r["ini"]).days + 1
    return (f"<b>{html.escape(r['Variable'])}</b> ({r['origen']})<br>"
            f"{r['ini']:%d %b} – {r['fin']:%d %b %Y} ({dias} d)<br>"
            f"Nota: {html.escape(r['nota']) or '—'}<br>"
            f"Impacto estimado: {imp}<br>Lugar: {html.escape(str(r['lugar']))}"
            + texto_comentarios(r["coment"], r["ini"], r["fin"]))


def etiqueta_y(v: str, ancho: int = 24) -> str:
    """Parte los nombres largos de variable en dos renglones."""
    return "<br>".join(textwrap.wrap(v, ancho)[:2]) if len(v) > ancho else v


def altura_panel_drivers(n_vars: int) -> int:
    return 26 * max(n_vars, 1) + 50


def agregar_barras_drivers(fig, seg: pd.DataFrame, orden_vars: list[str], row: int, col: int = 1):
    """Dibuja las barras de drivers (una fila por variable) y configura el eje Y de esa fila."""
    pos = {v: i for i, v in enumerate(orden_vars)}
    h = 0.36
    cx, cy, ct = [], [], []
    for _, r in seg.iterrows():
        yc = pos[r["Variable"]]
        fin_b = r["fin"] + pd.Timedelta(days=1)
        fig.add_trace(go.Scatter(
            x=[r["ini"], fin_b, fin_b, r["ini"], r["ini"]],
            y=[yc - h, yc - h, yc + h, yc + h, yc - h], mode="lines", fill="toself",
            fillcolor=color_impacto(r["impacto"])[0], line=dict(color="white", width=1.5),
            hoveron="fills", text=hover_barra(r), hoverinfo="text", showlegend=False),
            row=row, col=col)
        if r["nota"]:
            cx.append(r["ini"] + (fin_b - r["ini"]) / 2)
            cy.append(yc)
            ct.append(r["nota"])
    if ct:
        fig.add_trace(go.Scatter(x=cx, y=cy, mode="text", text=ct, hoverinfo="skip",
                                 textfont=dict(color="white", size=11), showlegend=False),
                      row=row, col=col)
    for nombre, color in (("Favorece humedad", mezcla(*AZUL, 0.8)),
                          ("Favorece condiciones secas", mezcla(*OCRE, 0.8)),
                          ("Presente, efecto neutro/incierto", NEUTRO)):
        fig.add_trace(go.Scatter(x=[None], y=[None], mode="markers", name=nombre,
                                 marker=dict(color=color, size=12, symbol="square")),
                      row=row, col=col)
    fig.update_yaxes(tickmode="array", tickvals=list(range(len(orden_vars))),
                     ticktext=[etiqueta_y(v) for v in orden_vars],
                     range=[len(orden_vars) - 0.4, -0.6], row=row, col=col, gridcolor="#efefef",
                     tickfont=dict(size=10), automargin=True)


def pie_version() -> str:
    """Pie de página con la fecha de generación y la versión del código (para detectar páginas viejas)."""
    return (f'<div style="padding:6px 20px 14px;font-size:11px;color:#8a949e">Generado el '
            f'{datetime.now():%Y-%m-%d %H:%M} · versión del código {VERSION_CODIGO}</div>')


def _cuenca_de_columna(col: str, cuencas: list[str], hoja: str) -> str | None:
    """Cuenca de aportes que corresponde a una columna: nombre normalizado igual o, si no,
    el más parecido (>= 0.85, p. ej. 'Ceunca 2' -> 'Cuenca 2') con una advertencia en el log."""
    mapa = {norm(c): c for c in cuencas}
    n = norm(col)
    if n in mapa:
        return mapa[n]
    cerca = difflib.get_close_matches(n, list(mapa), n=1, cutoff=0.85)
    if cerca:
        log.warning("Otros pronósticos [%s]: la columna '%s' no coincide exacto; se asocia a la cuenca '%s'",
                    hoja, col, mapa[cerca[0]])
        return mapa[cerca[0]]
    log.warning("Otros pronósticos [%s]: la columna '%s' no corresponde a ninguna cuenca de Seguimiento_Aportes; se ignora",
                hoja, col)
    return None


def leer_otros_pronosticos(ruta: Path, fechas: dict, cuencas: list[str]) -> dict[str, list[dict]]:
    """Lee las hojas mensuales / semanales / diarias y devuelve {cuenca: [serie, ...]}.

    Cada serie: {modelo, frecuencia, nombre, color, dias: {Timestamp: valor}}.
    mensual -> el valor aplica a todos los días del mes; semanal (S1..S5 del consenso) -> a todos los
    días de esa semana según la hoja Fechas; diario -> un valor por día."""
    if not ruta.exists():
        log.info("Otros pronósticos: no existe %s; se continúa sin ellos", ruta.name)
        return {}
    xl = pd.ExcelFile(ruta)
    acum: dict[tuple, dict] = {}   # (modelo, freq, cuenca) -> {día: valor}
    for hoja in xl.sheet_names:
        n = norm(hoja)
        freq = "mensual" if "mensual" in n else "semanal" if "semanal" in n else "diario" if "diari" in n else None
        if freq is None:
            log.warning("Otros pronósticos: la hoja '%s' no es mensual/semanal/diaria; se ignora", hoja)
            continue
        df = pd.read_excel(ruta, sheet_name=hoja)
        df.columns = [str(c).strip() for c in df.columns]
        if df.shape[1] < 3:
            log.warning("Otros pronósticos [%s]: faltan columnas; se ignora", hoja)
            continue
        col_mod, col_fecha = df.columns[0], df.columns[1]
        mapa_cols = {c: _cuenca_de_columna(c, cuencas, hoja) for c in df.columns[2:]}
        df = df.dropna(subset=[col_mod, col_fecha], how="any")
        log.info("Otros pronósticos [%s]: %d filas de modelo; columnas->cuenca: %s", hoja, len(df),
                 {k: v for k, v in mapa_cols.items() if v})
        for _, r in df.iterrows():
            modelo = str(r[col_mod]).strip()
            clave = r[col_fecha]
            if not modelo:
                continue
            if freq == "mensual":
                t = pd.to_datetime(clave, errors="coerce")
                if pd.isna(t):
                    log.warning("Otros pronósticos [%s]: mes ilegible %r (%s)", hoja, clave, modelo)
                    continue
                ini = t.normalize().replace(day=1)
                dias = list(pd.date_range(ini, ini + pd.offsets.MonthEnd(0)))
            elif freq == "semanal":
                m = re.search(r"(\d+)", str(clave))
                sem = next((s for s in fechas.get("semanas_dibujo", fechas["semanas"])
                            if m and s["n"] == int(m.group(1))), None)
                if sem is None:
                    log.warning("Otros pronósticos [%s]: la semana %r (%s) no existe en la hoja Fechas del consenso",
                                hoja, clave, modelo)
                    continue
                dias = list(pd.date_range(sem["ini"], sem["fin"]))
            else:
                t = pd.to_datetime(clave, errors="coerce")
                if pd.isna(t):
                    log.warning("Otros pronósticos [%s]: día ilegible %r (%s)", hoja, clave, modelo)
                    continue
                dias = [t.normalize()]
            for col, cuenca in mapa_cols.items():
                if cuenca is None:
                    continue
                v = pd.to_numeric(r[col], errors="coerce")
                if pd.isna(v):
                    continue
                d = acum.setdefault((modelo, freq, cuenca), {})
                for dia in dias:
                    if dia in d and d[dia] != float(v):
                        log.warning("Otros pronósticos: %s/%s/%s repetido el %s; gana el último valor",
                                    modelo, freq, cuenca, dia.date())
                    d[dia] = float(v)
    modelos = sorted({k[0] for k in acum})
    freqs_por_modelo = {m: {k[1] for k in acum if k[0] == m} for m in modelos}
    color = {m: PALETA_OTROS[i % len(PALETA_OTROS)] for i, m in enumerate(modelos)}
    out: dict[str, list[dict]] = {}
    for (modelo, freq, cuenca), dias in sorted(acum.items()):
        nombre = modelo if len(freqs_por_modelo[modelo]) == 1 else f"{modelo} ({freq})"
        out.setdefault(cuenca, []).append({"modelo": modelo, "frecuencia": freq, "nombre": nombre,
                                           "color": color[modelo], "dias": dias})
    log.info("Otros pronósticos: modelos=%s | cuencas con datos=%s", modelos, sorted(out))
    return out


def agregar_otros_pronosticos(fig, series: list[dict], ini, x_fin, row: int = 1, col: int = 1) -> int:
    """Una traza por modelo (y frecuencia) con leyenda propia 'legend2'. Devuelve cuántas dibujó."""
    n = 0
    for s in series:
        dias = sorted(d for d in s["dias"] if ini <= d < x_fin)
        if not dias:
            log.debug("Otros pronósticos: %s sin días en la ventana del gráfico", s["nombre"])
            continue
        xs, ys, hv = [], [], []
        paso = s["frecuencia"] in ("mensual", "semanal")
        bloques, actual = [], [dias[0]]
        for d in dias[1:]:  # tramos de días consecutivos (un hueco corta la línea)
            if (d - actual[-1]).days == 1:
                actual.append(d)
            else:
                bloques.append(actual)
                actual = [d]
        bloques.append(actual)
        for b in bloques:
            if xs:
                xs.append(None), ys.append(None), hv.append("")
            if paso:   # meseta: de 00:00 del primer día a 24:00 del último
                xs.append(b[0]), ys.append(s["dias"][b[0]]), hv.append(f"{b[0]:%d %b %Y}")
            for d in b:
                xs.append(d + pd.Timedelta(hours=12))
                ys.append(s["dias"][d])
                hv.append(f"{d:%d %b %Y}")
            if paso:
                xs.append(b[-1] + pd.Timedelta(days=1)), ys.append(s["dias"][b[-1]]), hv.append(f"{b[-1]:%d %b %Y}")
        est = _ESTILO_FREQ[s["frecuencia"]]
        fig.add_trace(go.Scatter(
            x=xs, y=ys, mode="lines" if paso else "lines+markers", name=s["nombre"], legend="legend2",
            line=dict(color=s["color"], **est), marker=dict(size=6, color=s["color"]),
            connectgaps=False, text=hv,
            hovertemplate="<b>%{fullData.name}</b> (" + s["frecuencia"] + ")<br>%{text}<br>%{y:.1f} %<extra></extra>"),
            row=row, col=col)
        n += 1
    if n:
        fig.update_layout(legend2=dict(title=dict(text="Otros pronósticos"), orientation="v", x=0.995,
                                       xanchor="right", y=0.99, yanchor="top",
                                       bgcolor="rgba(255,255,255,0.75)", bordercolor="#d9dde3",
                                       borderwidth=1, font=dict(size=11)))
    return n




# --------------------------------------------------------------------------
# Figura de una cuenca
# --------------------------------------------------------------------------
def construir_figura(cuenca, serie, seg, orden_vars, ini, hor_ini, hor_fin, otros=(), semanas=()):
    # el eje llega hasta el fin de la última semana llena (si es posterior al fin del horizonte)
    x_fin = max([hor_fin] + [w["fin"] for w in semanas]) + pd.Timedelta(days=1)
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
                   fillcolor="rgba(31,95,139,0.05)", line=dict(width=0), layer="below")]
    # Semanas (hoja Fechas, solo las que tienen inicio y fin llenos): límites con una línea suave
    limites = sorted({x for w in semanas for x in (w["ini"], w["fin"] + pd.Timedelta(days=1))
                      if ini <= x <= x_fin and x != hor_ini})
    for x in limites:
        formas.append(dict(type="line", xref="x", yref="paper", x0=x, x1=x, y0=0, y1=1, layer="below",
                           line=dict(color="rgba(90,110,130,0.40)", width=1, dash="dot")))
    formas.append(dict(type="line", xref="x", yref="paper", x0=hor_ini, x1=hor_ini, y0=0, y1=1.0,
                       line=dict(color="#d62728", width=2, dash="dash")))
    # Rótulos S1, S2... centrados en la parte visible de cada semana, en la franja entre los dos paneles
    y_franja = (fig.layout.yaxis.domain[0] + fig.layout.yaxis2.domain[1]) / 2
    rotulos = []
    for w in semanas:
        a, b = max(w["ini"], ini), min(w["fin"] + pd.Timedelta(days=1), x_fin)
        if a < b:
            rotulos.append(dict(x=a + (b - a) / 2, y=y_franja, xref="x", yref="paper", yanchor="middle",
                                text=f"S{w['n']}", showarrow=False, font=dict(size=11, color="#5b6773")))
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
        annotations=list(fig.layout.annotations) + anot + rotulos, hoverlabel=dict(align="left"),
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

        fechas = leer_fechas(args.consenso, args.fecha_consenso, estricto=False)   # solo hoja Fechas (última fila)
        hor_ini, hor_fin = fechas["hor_ini"], fechas["hor_fin"]
        semanas = fechas["semanas_dibujo"]
        fin_graf = max([hor_fin] + [w["fin"] for w in semanas])   # hasta donde llega el gráfico
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
            serie = aportes.loc[(aportes["Fecha"] >= ini) & (aportes["Fecha"] <= fin_graf),
                                ["Fecha", cuenca]].dropna()
            if serie.empty:
                log.warning("Cuenca '%s': sin datos en la ventana -> no se incluye", cuenca)
                omitidas.append((cuenca, "sin datos en la ventana"))
                continue
            seg, orden_vars = drivers_combinados(cuenca, ctx, ini, fin_graf)
            fig = construir_figura(cuenca, serie, seg, orden_vars, ini, hor_ini, hor_fin,
                                   otros.get(cuenca, []), semanas)
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
