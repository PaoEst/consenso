# -*- coding: utf-8 -*-
"""Lectura de Consenso.xlsx y utilidades compartidas por los scripts de consenso."""
from __future__ import annotations

import re
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter
import plotly.offline

from graficar_aportes import BASE, INSUMOS, SALIDAS, VERSION_CODIGO, log, norm

ARCHIVO_CONSENSO = "Consenso.xlsx"
ARCHIVO_PRONOSTICO = "Pronostico_Drivers.xlsx"
HOJA_FECHAS = "Fechas"
HOJA_EVALUADORES = "Evaluadores_Tecnicos"
HOJA_REGISTROS = "Registros"
N_SEMANAS = 5

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto",
         "septiembre", "octubre", "noviembre", "diciembre"]


def fecha_corta(f: pd.Timestamp) -> str:
    return f"{f.day:02d} {MESES[f.month - 1][:3]}"


def etiqueta_semana_corta(s: dict) -> str:
    """'01-07 nov' o, si cruza de mes, '26 nov–2 dic' (corto para que quepa en una sola línea)."""
    if s["ini"].month == s["fin"].month:
        return f"{s['ini'].day:02d}-{fecha_corta(s['fin'])}"
    return f"{s['ini'].day} {MESES[s['ini'].month - 1][:3]}–{s['fin'].day} {MESES[s['fin'].month - 1][:3]}"


def _fecha(v):
    f = pd.to_datetime(v, errors="coerce", dayfirst=isinstance(v, str))
    return None if pd.isna(f) else f.normalize()


_TERMINO = re.compile(r"([+-]?)([A-Za-z]{1,2}[0-9]+|[0-9]+(?:\.[0-9]+)?)")


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
    """Filas de la hoja Fechas con un consenso (columna B con fecha o fórmula), en orden."""
    filas = []
    for r in range(2, ws_f.max_row + 1):
        if ws_v.cell(r, 2).value not in (None, "") or ws_f.cell(r, 2).value not in (None, ""):
            filas.append(r)
    return filas


def _parsear_fila_fechas(ws_f, ws_v, fila: int, silencioso: bool = False) -> dict:
    """Fecha del consenso, horizonte y semanas S1..S5 (inicio/fin) de una fila de la hoja Fechas."""
    aviso = log.debug if silencioso else log.warning
    info = log.debug if silencioso else log.info
    v = [_valor_celda(ws_f, ws_v, f"{get_column_letter(c)}{fila}") for c in range(1, 5 + 2 * N_SEMANAS)]
    out = {"fila": fila, "seguimiento_ini": _fecha(v[0]), "consenso": _fecha(v[1]),
           "hor_ini": _fecha(v[2]), "hor_fin": _fecha(v[3]), "semanas": []}
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
    if not out["semanas"]:
        raise ValueError(f"Hoja Fechas fila {fila}: no hay ninguna semana dentro del horizonte")
    return out


def _abrir_fechas(ruta: Path):
    wb_f = openpyxl.load_workbook(ruta)
    wb_v = openpyxl.load_workbook(ruta, data_only=True)
    if HOJA_FECHAS not in wb_f.sheetnames:
        raise ValueError(f"{ruta.name} no tiene la hoja '{HOJA_FECHAS}'")
    return wb_f[HOJA_FECHAS], wb_v[HOJA_FECHAS]


def leer_fechas(ruta: Path, fecha=None) -> dict:
    """Datos del consenso en curso. Cada consenso es una FILA de la hoja Fechas (historial);
    por defecto se usa la última fila con datos. Con `fecha` (AAAA-MM-DD) se usa la fila cuya
    'Fecha de realización del consenso' coincide (la última si hay repetidas)."""
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
    out = _parsear_fila_fechas(ws_f, ws_v, fila)
    log.info("Fechas: usando el consenso %d de %d (fila %d de la hoja; %s): consenso=%s horizonte=%s -> %s semanas=%s",
             filas.index(fila) + 1, len(filas), fila, motivo, out["consenso"].date(), out["hor_ini"].date(), out["hor_fin"].date(),
             [(s["n"], str(s["ini"].date()), str(s["fin"].date()), s["dias"], s["dias_mes"]) for s in out["semanas"]])
    return out


def leer_todas_fechas(ruta: Path) -> dict:
    """{Timestamp fecha_consenso: datos de esa fila} con todo el historial de la hoja Fechas.
    Filas ilegibles se saltan con una advertencia; si una fecha se repite, gana la última fila."""
    ws_f, ws_v = _abrir_fechas(ruta)
    out = {}
    for r in _filas_fechas(ws_f, ws_v):
        try:
            d = _parsear_fila_fechas(ws_f, ws_v, r, silencioso=True)
        except ValueError as e:
            log.warning("Fechas: fila %d ignorada en el historial: %s", r, e)
            continue
        if d["consenso"] in out:
            log.warning("Fechas: la fecha de consenso %s está repetida (filas %d y %d); se usa la última",
                        d["consenso"].date(), out[d["consenso"]]["fila"], r)
        out[d["consenso"]] = d
    log.info("Fechas: historial de %d consensos: %s", len(out), [str(k.date()) for k in sorted(out)])
    return out


def guardar_excel(wb, ruta: Path, dir_logs: Path) -> None:
    """Copia de seguridad (dir_logs/backups) + guardado atómico del libro."""
    import os
    import shutil
    dir_bk = dir_logs / "backups"
    dir_bk.mkdir(parents=True, exist_ok=True)
    bk = dir_bk / f"{ruta.stem}_{datetime.now():%Y%m%d_%H%M%S}{ruta.suffix}"
    shutil.copy2(ruta, bk)
    log.info("Copia de seguridad: %s", bk.resolve())
    tmp = ruta.with_name(ruta.stem + ".tmp" + ruta.suffix)
    wb.save(tmp)
    try:
        os.replace(tmp, ruta)
    except PermissionError:
        tmp.unlink(missing_ok=True)
        raise PermissionError(f"No se pudo escribir {ruta.name}: ciérralo en Excel y vuelve a correr "
                              f"(no se guardó nada; la copia {bk.name} no cambió nada).")
    log.info("Guardado %s", ruta.resolve())


def leer_evaluadores(ruta: Path) -> dict[str, str]:
    """{INICIALES_MAYUSCULA: nombre} de la hoja Evaluadores_Tecnicos."""
    df = pd.read_excel(ruta, sheet_name=HOJA_EVALUADORES)
    df = df.iloc[:, :2].dropna(how="all")
    df.columns = ["nombre", "iniciales"]
    out = {}
    for _, r in df.dropna(subset=["iniciales"]).iterrows():
        ini = str(r["iniciales"]).strip().upper()
        if ini in out:
            log.warning("Evaluadores: iniciales repetidas '%s'; se conserva la primera", ini)
            continue
        out[ini] = str(r["nombre"]).strip()
    log.info("Evaluadores técnicos (%d): %s", len(out), out)
    return out


def carpeta_corrida(base: Path, fecha: pd.Timestamp) -> Path:
    """<base>/<año>/<mes>/<AAAA-MM-DD>"""
    return base / f"{fecha.year}" / f"{fecha.month:02d}" / f"{fecha:%Y-%m-%d}"


def plotly_js() -> str:
    """Código de plotly.js para incrustarlo en el HTML (funciona sin internet)."""
    return plotly.offline.get_plotlyjs()


# ==========================================================================
# Drivers climáticos (seguimiento + pronóstico) compartidos por los dos gráficos
# ==========================================================================
import html  # noqa: E402
import json  # noqa: E402
import textwrap  # noqa: E402

import plotly.graph_objects as go  # noqa: E402

from graficar_aportes import (ARCHIVO_APORTES, ARCHIVO_DRIVERS, AZUL, NEUTRO, OCRE,  # noqa: E402,F401
                              color_impacto, drivers_de_cuenca, leer_drivers, leer_prioridad,
                              mezcla, segmentos, texto_comentarios)

ARCHIVO_INDICE = "indice_cuencas.json"   # archivo -> nombre de cuenca (para el índice/selector)
CARPETA_CONSOLIDADO = "consolidado"
CARPETA_INFORMATIVO = "informativo"       # páginas sin votación (solo consulta)
PREFIJO_INFORMATIVO = "Informativo"
PREFIJO_CONSOLIDADO = "Consolidado"


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


def real_a_la_fecha(aportes: pd.DataFrame, cuenca: str, hor_ini: pd.Timestamp) -> list[float]:
    """Aportes con dato del mes del inicio del horizonte, anteriores al horizonte (no mezcla meses)."""
    f = aportes["Fecha"]
    m = aportes[(f.dt.year == hor_ini.year) & (f.dt.month == hor_ini.month) & (f < hor_ini)][cuenca]
    return [float(v) for v in m.dropna()]


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


# ==========================================================================
# Navegación entre cuencas (índice + selector) con enlaces relativos
# ==========================================================================
def registrar_cuenca(carpeta: Path, archivo: str, cuenca: str) -> None:
    ruta = carpeta / ARCHIVO_INDICE
    d = json.loads(ruta.read_text(encoding="utf-8")) if ruta.exists() else {}
    d[archivo] = cuenca
    ruta.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")


def cuencas_registradas(carpeta: Path) -> dict[str, str]:
    """{archivo_base: nombre} (archivo_base = nombre_archivo(cuenca))."""
    ruta = carpeta / ARCHIVO_INDICE
    return json.loads(ruta.read_text(encoding="utf-8")) if ruta.exists() else {}


def barra_navegacion(carpeta: Path, actual: str, donde: str, ronda: int | None = None) -> str:
    """HTML de la barra superior: enlace al índice, selector de cuenca (y de ronda) y enlaces cruzados.

    donde = 'votacion' (carpeta de corrida) | 'consolidado' (carpeta/consolidado/) |
            'informativo' (carpeta/informativo/, sin votación)."""
    cuencas = cuencas_registradas(carpeta)
    en_cons = donde == "consolidado"
    en_info = donde == "informativo"
    ruta_cons = carpeta / CARPETA_CONSOLIDADO
    prefijo = "../" if (en_cons or en_info) else ""
    opciones = []
    for archivo, nombre in cuencas.items():
        if en_cons:
            if ronda is None:
                continue
            # misma ronda si existe; si no, la más cercana anterior (o la primera disponible)
            disp = sorted(int(m.group(1)) for f in ruta_cons.glob(f"{PREFIJO_CONSOLIDADO}_{archivo}_Ronda*.html")
                          if (m := re.search(r"Ronda(\d+)\.html$", f.name)))
            if not disp:
                continue
            r_dest = ronda if ronda in disp else max([r for r in disp if r < ronda] or [disp[0]])
            destino = f"{PREFIJO_CONSOLIDADO}_{archivo}_Ronda{r_dest}.html"
        elif en_info:
            destino = f"{PREFIJO_INFORMATIVO}_{archivo}.html"
            if not (carpeta / CARPETA_INFORMATIVO / destino).exists():
                continue
        else:
            destino = f"Consenso_{archivo}.html"
            if not (carpeta / destino).exists():
                continue
        sel = " selected" if archivo == actual else ""
        opciones.append(f'<option value="{html.escape(destino)}"{sel}>{html.escape(nombre)}</option>')
    partes = [f'<a href="{prefijo}index.html">‹ Índice</a>']
    if opciones:
        partes.append('<label>Cuenca <select onchange="location.href=this.value">'
                      + "".join(opciones) + "</select></label>")
    if en_cons and ronda is not None:
        rondas = sorted({int(m.group(1)) for f in ruta_cons.glob(f"{PREFIJO_CONSOLIDADO}_{actual}_Ronda*.html")
                         if (m := re.search(r"Ronda(\d+)\.html$", f.name))})
        if len(rondas) > 1:
            ops = "".join(f'<option value="{PREFIJO_CONSOLIDADO}_{actual}_Ronda{r}.html"'
                          f'{" selected" if r == ronda else ""}>Ronda {r}</option>' for r in rondas)
            partes.append(f'<label>Ronda <select onchange="location.href=this.value">{ops}</select></label>')
    if en_info and (carpeta / f"Consenso_{actual}.html").exists():
        partes.append(f'<a href="../Consenso_{html.escape(actual)}.html">Ir a la votación ›</a>')
    if donde == "votacion" and (carpeta / CARPETA_INFORMATIVO / f"{PREFIJO_INFORMATIVO}_{actual}.html").exists():
        partes.append(f'<a href="{CARPETA_INFORMATIVO}/{PREFIJO_INFORMATIVO}_{html.escape(actual)}.html">Ver informativo (sin votación) ›</a>')
    return '<nav class="nav">' + "".join(partes) + "</nav>"


def pie_version() -> str:
    """Pie de página con la fecha de generación y la versión del código (para detectar páginas viejas)."""
    return (f'<div style="padding:6px 20px 14px;font-size:11px;color:#8a949e">Generado el '
            f'{datetime.now():%Y-%m-%d %H:%M} · versión del código {VERSION_CODIGO}</div>')


ESTILO_NAV = ("<style>.nav{display:flex;flex-wrap:wrap;gap:18px;align-items:center;padding:8px 20px;"
              "font-size:13px;border-bottom:1px solid var(--brd)} .nav a{color:var(--azul);text-decoration:none;"
              "font-weight:600} .nav select{font-size:13px;padding:3px 6px;margin-left:6px;border-radius:5px;"
              "border:1px solid var(--brd);background:var(--bg);color:var(--txt)}</style>")


def escribir_indice(carpeta: Path, fechas: dict) -> Path:
    """index.html de la corrida: una fila por cuenca con enlace a votación y a cada ronda del Consolidado."""
    cuencas = cuencas_registradas(carpeta)
    filas = []
    for archivo, nombre in cuencas.items():
        if not (carpeta / f"Consenso_{archivo}.html").exists():
            continue
        enlaces = [f'<a href="Consenso_{html.escape(archivo)}.html">Votación</a>']
        if (carpeta / CARPETA_INFORMATIVO / f"{PREFIJO_INFORMATIVO}_{archivo}.html").exists():
            enlaces.append(f'<a href="{CARPETA_INFORMATIVO}/{PREFIJO_INFORMATIVO}_{html.escape(archivo)}.html">Informativo</a>')
        for f in sorted((carpeta / CARPETA_CONSOLIDADO).glob(f"{PREFIJO_CONSOLIDADO}_{archivo}_Ronda*.html")):
            m = re.search(r"Ronda(\d+)\.html$", f.name)
            enlaces.append(f'<a href="{CARPETA_CONSOLIDADO}/{html.escape(f.name)}">Consolidado ronda {m.group(1)}</a>')
        filas.append(f"<tr><td>{html.escape(nombre)}</td><td>{' · '.join(enlaces)}</td></tr>")
    pagina = f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Consenso de aportes</title><style>
:root{{--txt:#1d2733;--sub:#5b6773;--bg:#fff;--brd:#d9dde3;--azul:#1f5f8b}}
@media (prefers-color-scheme: dark){{:root{{--txt:#e8ecf1;--sub:#a3adb8;--bg:#161b22;--brd:#34404d;--azul:#7fb3e6}}}}
body{{font-family:Segoe UI,Arial,sans-serif;color:var(--txt);background:var(--bg);margin:0;padding:28px 36px}}
h1{{font-size:22px;margin:0 0 4px}} p{{color:var(--sub);margin:0 0 18px;font-size:14px}}
table{{border-collapse:collapse;min-width:420px}} td,th{{border-bottom:1px solid var(--brd);padding:9px 16px;text-align:left;font-size:14px}}
th{{color:var(--sub);font-weight:600}} a{{color:var(--azul);font-weight:600;text-decoration:none}} a:hover{{text-decoration:underline}}
</style></head><body><h1>Consenso de aportes por cuenca</h1>
<p>Pronóstico {MESES[fechas['hor_ini'].month - 1]} {fechas['hor_ini'].year} · horizonte {fechas['hor_ini']:%d/%m/%Y} – {fechas['hor_fin']:%d/%m/%Y}
· consenso del {fechas['consenso']:%d/%m/%Y}</p>
<table><tr><th>Cuenca</th><th>Abrir</th></tr>{''.join(filas)}</table></body></html>"""
    ruta = carpeta / "index.html"
    ruta.write_text(pagina, encoding="utf-8")
    return ruta


# ==========================================================================
# Otros pronósticos (modelos externos) - Otros_pronosticos.xlsx
# ==========================================================================
import difflib  # noqa: E402

ARCHIVO_OTROS = "Otros_pronosticos.xlsx"
PALETA_OTROS = ["#c2185b", "#00897b", "#7b1fa2", "#6d4c41", "#827717", "#d32f2f", "#5d4037",
                "#00838f", "#ef6c00", "#455a64"]
_ESTILO_FREQ = {"mensual": dict(dash="dash", width=3), "semanal": dict(dash="dot", width=3),
                "diario": dict(dash="solid", width=2)}


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
                sem = next((s for s in fechas["semanas"] if m and s["n"] == int(m.group(1))), None)
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
