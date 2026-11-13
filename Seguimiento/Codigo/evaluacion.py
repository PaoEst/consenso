#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Evaluación semanal del consenso: compara el pronóstico de cada semana (S1..S5) del consenso
marcado `Consenso = SI` en la hoja Consolidado con los aportes observados de Seguimiento_Aportes.

Resultado: hoja `Evaluación` de Consenso.xlsx, en formato largo (una fila por semana y métrica):
    Fecha_consenso, Cuenca, No Evaluación, Evaluador, horizonte, fecha_inicio_horizonte,
    fecha_fin_horizonte, Valor_pronos, Valor_real, Metrica_desempeño, Valor_Metrica,
    Estadístico, Dias_evaluados, Dias_semana, Estado, Fecha_corte, Fecha_evaluacion

Métricas (P = valor pronosticado de la semana, igual para todos sus días; Od = observado del día):
    MAPE = 100 · media(|P − Od| / Od)         días con dato y Od > 0 (los días con 0 se excluyen)
    MSE  = media((P − Od)²)                   sobre los días con dato
    FB   = 2 · (Ō − P) / (Ō + P)              sesgo fraccional; positivo = el pronóstico subestima
Para juntar semanas (p. ej. todo un mes): MSE y MAPE se ponderan por Dias_evaluados
(RMSE = √MSE ponderado).

Estado: `Pendiente` si la semana aún no tiene ningún día observado (Valor_real y métricas en blanco),
`Parcial` mientras falten días por observar (el último dato es anterior al fin de la semana) y `Evaluada`
cuando la semana ya cerró; las filas `Evaluada` se congelan (usar --reevaluar para recalcular).
La hoja `Evaluación` se crea siempre en Consenso.xlsx, aunque no haya nada que evaluar.

Se ejecuta sola al final de crear_graficos_consenso.py, o a mano:  python evaluacion.py [--reevaluar]
"""
from __future__ import annotations

import argparse
import getpass
import platform
import sys
import time
import uuid
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd

from consenso_comun import (ARCHIVO_CONSENSO, INSUMOS, N_SEMANAS, _fecha, guardar_excel,
                            leer_todas_fechas)
from graficar_aportes import (ARCHIVO_APORTES, BASE, configurar_logging, leer_aportes, log, norm,
                              sha256)

HOJA_CONSOLIDADO = "Consolidado"
HOJA_EVALUACION = "Evaluación"
COLUMNAS_EVAL = ["Fecha_consenso", "Cuenca", "No Evaluación", "Evaluador", "horizonte",
                 "fecha_inicio_horizonte", "fecha_fin_horizonte", "Valor_pronos", "Valor_real",
                 "Metrica_desempeño", "Valor_Metrica", "Estadístico", "Dias_evaluados", "Dias_semana",
                 "Estado", "Fecha_corte", "Fecha_evaluacion"]
COLUMNAS_FECHA = ("Fecha_consenso", "fecha_inicio_horizonte", "fecha_fin_horizonte", "Fecha_corte",
                  "Fecha_evaluacion")
METRICAS = ("MAPE", "MSE", "FB")
CLAVE = ("Fecha_consenso", "Cuenca", "No Evaluación", "Evaluador", "horizonte", "Metrica_desempeño")


# --------------------------------------------------------------------------
# Métricas
# --------------------------------------------------------------------------
def calcular_metricas(p: float, obs: list[float]) -> dict[str, float | None]:
    """MAPE, MSE y FB de una semana: `p` pronosticado (constante) y `obs` observados diarios."""
    n = len(obs)
    mse = sum((p - o) ** 2 for o in obs) / n
    validos = [o for o in obs if o > 0]
    mape = 100 * sum(abs(p - o) / o for o in validos) / len(validos) if validos else None
    media_o = sum(obs) / n
    fb = 2 * (media_o - p) / (media_o + p) if (media_o + p) != 0 else None
    return {"MAPE": mape, "MSE": mse, "FB": fb}


# --------------------------------------------------------------------------
# Lectura del Consolidado
# --------------------------------------------------------------------------
def _cols(ws) -> dict[str, int]:
    return {norm(ws.cell(1, c).value): c for c in range(1, ws.max_column + 1) if ws.cell(1, c).value}


def consensos_seleccionados(wb) -> list[dict]:
    """Filas de Consolidado con Consenso = SI (una por mes + fecha de consenso + cuenca)."""
    if HOJA_CONSOLIDADO not in wb.sheetnames:
        log.info("Evaluación: el libro no tiene la hoja '%s'; no hay nada que evaluar", HOJA_CONSOLIDADO)
        return []
    ws = wb[HOJA_CONSOLIDADO]
    cols = _cols(ws)
    faltan = [c for c in ("Fecha_consenso", "Cuenca", "No Evaluación", "Evaluador", "Estadístico", "Consenso")
              if norm(c) not in cols]
    if faltan:
        log.warning("Evaluación: la hoja Consolidado no tiene las columnas %s (¿falta correr consolidar_votos.py "
                    "para actualizar su formato?); no se evalúa", faltan)
        return []
    g = lambda r, n: ws.cell(r, cols[norm(n)]).value
    elegidos: dict[tuple, dict] = {}
    for r in range(2, ws.max_row + 1):
        if norm(g(r, "Consenso")) != "si":
            continue
        fecha = _fecha(g(r, "Fecha_consenso"))
        cuenca = str(g(r, "Cuenca") or "").strip()
        if fecha is None or not cuenca:
            log.warning("Evaluación: fila %d de Consolidado marcada SI sin fecha de consenso o cuenca; se omite", r)
            continue
        est = norm(g(r, "Estadístico"))
        if est not in ("media", "mediana"):
            log.warning("Evaluación: fila %d (%s, %s) marcada SI con estadístico '%s': no es un valor puntual; se omite",
                        r, cuenca, fecha.date(), g(r, "Estadístico"))
            continue
        try:
            ronda = int(g(r, "No Evaluación"))
        except (TypeError, ValueError):
            log.warning("Evaluación: fila %d con No Evaluación ilegible; se omite", r)
            continue
        fila = {"fila": r, "fecha": fecha, "cuenca": cuenca, "ronda": ronda,
                "evaluador": str(g(r, "Evaluador") or "").strip(), "estadistico": str(g(r, "Estadístico")).strip(),
                "sem": {f"S{k}": g(r, f"Aportes S{k}") for k in range(1, N_SEMANAS + 1)}}
        clave = (fecha, cuenca)
        if clave in elegidos:
            log.warning("Evaluación: más de un SI para %s el %s; se usa la ronda %d (la mayor)", cuenca,
                        fecha.date(), max(ronda, elegidos[clave]["ronda"]))
            if ronda < elegidos[clave]["ronda"]:
                continue
        elegidos[clave] = fila
    log.info("Evaluación: %d consensos marcados SI", len(elegidos))
    return list(elegidos.values())


# --------------------------------------------------------------------------
# Hoja Evaluación
# --------------------------------------------------------------------------
def _valor_ini(v):
    """Normaliza una celda a un valor comparable (fechas -> 'AAAA-MM-DD')."""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    return v


def _hoja_evaluacion(wb):
    """Devuelve (hoja, columnas, cambiada). Crea la hoja Evaluación si no existe (con encabezados)."""
    cambiada = False
    if HOJA_EVALUACION in wb.sheetnames:
        ws = wb[HOJA_EVALUACION]
    else:
        ws = wb.create_sheet(HOJA_EVALUACION)
        ref = wb[HOJA_CONSOLIDADO] if HOJA_CONSOLIDADO in wb.sheetnames else None
        for c, nombre in enumerate(COLUMNAS_EVAL, start=1):
            celda = ws.cell(1, c, nombre)
            if ref is not None and ref.cell(1, 1).has_style:
                celda._style = ref.cell(1, 1)._style
            ws.column_dimensions[celda.column_letter].width = (
                22 if nombre in ("Metrica_desempeño", "fecha_inicio_horizonte", "fecha_fin_horizonte",
                                 "Fecha_evaluacion") else 16)
        ws.freeze_panes = "A2"
        cambiada = True
        log.info("Se creó la hoja '%s'", HOJA_EVALUACION)
    cols = _cols(ws)
    for nombre in COLUMNAS_EVAL:   # agrega columnas que falten (por si se editó la hoja)
        if norm(nombre) not in cols:
            c = max(cols.values()) + 1
            ws.cell(1, c, nombre)
            cols[norm(nombre)] = c
            cambiada = True
    return ws, cols, cambiada


def _num(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return v


def _igual(a, b) -> bool:
    a, b = _num(_valor_ini(a)), _num(_valor_ini(b))
    if a is None or b is None:
        return a is None and b is None
    if isinstance(a, float) and isinstance(b, float):
        return abs(a - b) < 1e-6
    return a == b


def resumen_texto(cont: dict, ruta: Path) -> str:
    """Mensaje para pantalla: dónde quedó la evaluación y qué contiene."""
    if cont.get("seleccionados", 0) == 0:
        return (f"Evaluación: no hay consensos marcados SI en la hoja Consolidado; la hoja '{HOJA_EVALUACION}' de "
                f"{ruta.name} queda vacía.")
    return (f"Evaluación → hoja '{HOJA_EVALUACION}' de {ruta.name}: {cont['nuevas']} filas nuevas, "
            f"{cont['actualizadas']} actualizadas, {cont['congeladas']} congeladas; "
            f"{cont['semanas_pendientes']} semanas sin datos observados (celdas en blanco, estado Pendiente).")


def evaluar(ruta_consenso: Path, ruta_aportes: Path, dir_logs: Path, reevaluar: bool = False) -> dict:
    """Evalúa por semana los consensos marcados SI y escribe la hoja Evaluación (que siempre se crea).
    Semanas sin días observados quedan con Valor_real y métricas en blanco (Estado = Pendiente).
    Devuelve contadores {seleccionados, nuevas, actualizadas, congeladas, sin_cambio, semanas_pendientes}."""
    hoy = datetime.now().date()
    cont = {"seleccionados": 0, "nuevas": 0, "actualizadas": 0, "congeladas": 0, "sin_cambio": 0,
            "semanas_pendientes": 0}
    wb = openpyxl.load_workbook(ruta_consenso)
    ws, cols, cambiada = _hoja_evaluacion(wb)
    elegidos = consensos_seleccionados(wb)
    cont["seleccionados"] = len(elegidos)
    if not elegidos:
        log.info("Evaluación: sin consensos marcados SI; la hoja '%s' queda vacía", HOJA_EVALUACION)
        if cambiada:
            guardar_excel(wb, ruta_consenso, dir_logs)
        return cont
    fechas = leer_todas_fechas(ruta_consenso)
    aportes = leer_aportes(ruta_aportes)
    existentes = {}
    for r in range(2, ws.max_row + 1):
        k = tuple(_valor_ini(ws.cell(r, cols[norm(n)]).value) for n in CLAVE)
        if k[0] is not None:
            existentes[k] = r
    sig = max(existentes.values(), default=1) + 1

    for ch in elegidos:
        cuenca = ch["cuenca"]
        if cuenca not in aportes.columns:
            log.warning("Evaluación: la cuenca '%s' no existe en Seguimiento_Aportes; se omite", cuenca)
            continue
        info = fechas.get(ch["fecha"])
        if info is None:
            log.warning("Evaluación: no hay fila de Fechas con el consenso %s; se omite %s", ch["fecha"].date(), cuenca)
            continue
        serie = aportes.set_index("Fecha")[cuenca].dropna()
        corte = serie.index.max() if len(serie) else None
        for sem in info["semanas"]:
            clave_s = f"S{sem['n']}"
            p = ch["sem"].get(clave_s)
            if p is None or p == "":
                continue
            try:
                p = float(p)
            except (TypeError, ValueError):
                log.warning("Evaluación: %s %s %s tiene un valor no numérico (%r); se omite", cuenca,
                            ch["fecha"].date(), clave_s, p)
                continue
            dias = pd.date_range(sem["ini"], sem["fin"])
            obs = [float(v) for v in serie.reindex(dias).dropna()]
            if obs:
                estado = "Evaluada" if corte is not None and corte >= sem["fin"] else "Parcial"
                ceros = sum(1 for o in obs if o <= 0)
                if ceros:
                    log.warning("Evaluación: %s %s %s: %d días con aporte <= 0 se excluyen del MAPE", cuenca,
                                ch["fecha"].date(), clave_s, ceros)
                mets = calcular_metricas(p, obs)
                valor_real = round(sum(obs) / len(obs), 4)
            else:   # sin datos observados: todo en blanco hasta que lleguen
                estado = "Pendiente"
                mets = {m: None for m in METRICAS}
                valor_real = None
                cont["semanas_pendientes"] += 1
                log.info("Evaluación: %s %s %s sin días observados todavía (Pendiente)", cuenca,
                         ch["fecha"].date(), clave_s)
            base = {"Fecha_consenso": ch["fecha"].date(), "Cuenca": cuenca, "No Evaluación": ch["ronda"],
                    "Evaluador": ch["evaluador"], "horizonte": clave_s,
                    "fecha_inicio_horizonte": sem["ini"].date(), "fecha_fin_horizonte": sem["fin"].date(),
                    "Valor_pronos": round(p, 4), "Valor_real": valor_real,
                    "Estadístico": ch["estadistico"], "Dias_evaluados": len(obs), "Dias_semana": sem["dias"],
                    "Estado": estado, "Fecha_corte": corte.date() if corte is not None else None,
                    "Fecha_evaluacion": hoy}
            for metrica in METRICAS:
                valor = mets[metrica]
                if valor is None and obs:
                    log.info("Evaluación: %s %s %s: %s no calculable (sin días válidos); queda en blanco", cuenca,
                             ch["fecha"].date(), clave_s, metrica)
                fila = dict(base, Metrica_desempeño=metrica,
                            Valor_Metrica=None if valor is None else round(valor, 6))
                k = tuple(_valor_ini(fila[n]) for n in CLAVE)
                if k in existentes:
                    r = existentes[k]
                    previo = ws.cell(r, cols[norm("Estado")]).value
                    if norm(previo) == "evaluada" and not reevaluar:
                        cont["congeladas"] += 1
                        continue
                    if all(_igual(ws.cell(r, cols[norm(n)]).value, fila[n])
                           for n in COLUMNAS_EVAL if n != "Fecha_evaluacion"):
                        cont["sin_cambio"] += 1
                        continue
                    cont["actualizadas"] += 1
                    accion = "actualizada"
                else:
                    r = sig
                    sig += 1
                    cont["nuevas"] += 1
                    accion = "nueva"
                for n in COLUMNAS_EVAL:
                    celda = ws.cell(r, cols[norm(n)], fila[n])
                    if n in COLUMNAS_FECHA:
                        celda.number_format = "yyyy-mm-dd"
                log.info("Evaluación %s | %s | %s | %s %s | pron=%.2f real=%s (%d/%d días, %s) | %s=%s", accion,
                         cuenca, ch["fecha"].date(), clave_s, ch["estadistico"], p,
                         "" if valor_real is None else f"{valor_real:.2f}", len(obs), sem["dias"], estado,
                         metrica, "" if valor is None else f"{valor:.4f}")
    log.info("Evaluación: nuevas=%d actualizadas=%d congeladas=%d sin cambio=%d semanas pendientes=%d",
             cont["nuevas"], cont["actualizadas"], cont["congeladas"], cont["sin_cambio"],
             cont["semanas_pendientes"])
    if cambiada or cont["nuevas"] or cont["actualizadas"]:
        guardar_excel(wb, ruta_consenso, dir_logs)
    return cont


# --------------------------------------------------------------------------
# Programa principal (uso independiente)
# --------------------------------------------------------------------------
def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--aportes", default=INSUMOS / ARCHIVO_APORTES, type=Path)
    ap.add_argument("--consenso", default=INSUMOS / ARCHIVO_CONSENSO, type=Path)
    ap.add_argument("--logs", default=BASE / "logs", type=Path)
    ap.add_argument("--reevaluar", action="store_true", help="recalcula también las filas ya 'Evaluada'")
    args = ap.parse_args()
    run_id = uuid.uuid4().hex[:8]
    t0 = time.time()
    ruta_log = configurar_logging(args.logs, run_id)
    log.info("=" * 70)
    log.info("INICIO evaluacion run_id=%s | usuario=%s host=%s python=%s", run_id, getpass.getuser(),
             platform.node(), platform.python_version())
    log.info("Argumentos: %s", {k: str(v) for k, v in vars(args).items()})
    try:
        for p in (args.aportes, args.consenso):
            if not p.exists():
                raise FileNotFoundError(f"No se encontró el archivo Excel: {p} (los Excel deben estar en la carpeta Insumos)")
            log.info("Entrada: %s | sha256=%s", p.resolve(), sha256(p))
        cont = evaluar(args.consenso, args.aportes, args.logs, args.reevaluar)
        log.info("FIN OK run_id=%s duración=%.2fs log=%s", run_id, time.time() - t0, ruta_log.resolve())
        print("\n" + resumen_texto(cont, args.consenso))
        return 0
    except Exception:
        log.exception("FIN CON ERROR run_id=%s tras %.2fs", run_id, time.time() - t0)
        return 1


if __name__ == "__main__":
    sys.exit(main())
