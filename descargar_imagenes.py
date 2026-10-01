#!/usr/bin/env python3
"""Descarga diaria de imágenes listadas en Lista_de_Insumos.xlsx.

Estructura de salida:  <salida>/<variable>/<dd_mmm_aaaa>/<filename>_<YYYYMMDD>.<ext>
Ejemplo:               descargas/MJO/29_sep_2026/vel_pote_200hpa_hovm_20260929.gif

Columnas usadas del Excel: Variable, filename_type, filename, Link de imagen.
  - filename_type = "fijo":     la URL es siempre la misma.
  - filename_type = "variable": la URL contiene una fecha YYYYMMDD (p. ej.
    wk1.wk2_20260831.wnd850.png) que se reemplaza por la fecha del día.

Logs (carpeta logs/, siempre en modo append, nunca se sobrescriben):
  - ejecucion.log    : log legible de cada ejecución.
  - auditoria.jsonl  : un registro JSON por intento de descarga (URL, HTTP,
                       tamaño, SHA-256, ruta, duración, resultado, etc.).
"""
import argparse
import getpass
import hashlib
import json
import logging
import os
import platform
import re
import socket
import sys
import time
import urllib.error
import urllib.request
import uuid
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from urllib.parse import urlparse

BASE = Path(__file__).resolve().parent
MESES = ["ene", "feb", "mar", "abr", "may", "jun",
         "jul", "ago", "sep", "oct", "nov", "dic"]
FECHA_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")
EXT_POR_TIPO = {"image/png": ".png", "image/gif": ".gif", "image/jpeg": ".jpg",
                "image/webp": ".webp", "image/svg+xml": ".svg"}
USER_AGENT = "Mozilla/5.0 (compatible; descarga-imagenes-diaria/1.0)"

log = logging.getLogger("descargas")


# ----------------------------------------------------------------- utilidades
def carpeta_dia(d: date) -> str:
    return f"{d.day:02d}_{MESES[d.month - 1]}_{d.year}"


def limpiar(nombre: str) -> str:
    """Evita separadores de ruta / caracteres inválidos en nombres de carpeta."""
    nombre = re.sub(r'[\\/:*?"<>|]', "_", str(nombre)).strip().rstrip(".")
    return nombre or "sin_nombre"


def norm(s) -> str:
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()


def leer_tabla(ruta: Path):
    """Lee el Excel y devuelve filas con variable (heredada de celdas vacías)."""
    import openpyxl
    ws = openpyxl.load_workbook(ruta, data_only=True).worksheets[0]
    filas = ws.iter_rows(values_only=True)
    cab = [norm(c) for c in next(filas)]
    req = ["variable", "filename_type", "filename", "link de imagen"]
    faltan = [c for c in req if c not in cab]
    if faltan:
        raise SystemExit(f"Faltan columnas en {ruta.name}: {faltan}")
    idx = {c: cab.index(c) for c in req + (["insumo"] if "insumo" in cab else [])}
    out, variable = [], None
    for n, fila in enumerate(filas, start=2):
        g = lambda k: (str(fila[idx[k]]).strip() if k in idx and fila[idx[k]] else "")
        if g("variable"):
            variable = g("variable")  # las celdas vacías heredan la variable anterior
        if not (g("filename") and g("link de imagen")):
            continue
        out.append(dict(fila=n, variable=variable or "sin_variable",
                        insumo=g("insumo"), tipo=norm(g("filename_type")),
                        filename=g("filename"), url=g("link de imagen")))
    return out


def url_para_fecha(url: str, d: date) -> str:
    return FECHA_RE.sub(d.strftime("%Y%m%d"), url)


def extension(url: str, content_type: str) -> str:
    ext = os.path.splitext(urlparse(url).path)[1].lower()
    return ext if ext else EXT_POR_TIPO.get(content_type, ".bin")


def descargar(url: str, timeout: int, reintentos: int):
    """GET con reintentos. Devuelve (bytes, info). Lanza HTTPError si 4xx."""
    ultimo = None
    for intento in range(1, reintentos + 1):
        t0 = time.monotonic()
        try:
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
            with urllib.request.urlopen(req, timeout=timeout) as r:
                datos = r.read()
                return datos, dict(
                    http=r.status, url_final=r.geturl(), intentos=intento,
                    content_type=(r.headers.get_content_type() or "").lower(),
                    last_modified=r.headers.get("Last-Modified"),
                    etag=r.headers.get("ETag"),
                    ms=int((time.monotonic() - t0) * 1000))
        except urllib.error.HTTPError as e:
            if e.code < 500:      # 404 etc.: no tiene sentido reintentar
                e.intentos = intento
                raise
            ultimo = e
        except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as e:
            ultimo = e
        log.warning("Intento %d/%d falló para %s: %s", intento, reintentos, url, ultimo)
        if intento < reintentos:
            time.sleep(2 ** intento)
    raise ultimo


def guardar_atomico(destino: Path, datos: bytes):
    destino.parent.mkdir(parents=True, exist_ok=True)
    tmp = destino.with_name(destino.name + ".part")
    tmp.write_bytes(datos)
    os.replace(tmp, destino)


# ----------------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--tabla", type=Path, default=BASE / "Lista_de_Insumos.xlsx")
    ap.add_argument("--salida", type=Path, default=BASE / "descargas")
    ap.add_argument("--logs", type=Path, default=BASE / "logs")
    ap.add_argument("--fecha", help="AAAA-MM-DD (por defecto: hoy)")
    ap.add_argument("--retroceso", type=int, default=3,
                    help="filas 'variable': si la URL de la fecha no existe (404), "
                         "probar hasta N días anteriores (defecto 3)")
    ap.add_argument("--timeout", type=int, default=60)
    ap.add_argument("--reintentos", type=int, default=3)
    ap.add_argument("--forzar", action="store_true",
                    help="sobrescribir si el archivo del día ya existe")
    a = ap.parse_args()

    hoy = datetime.strptime(a.fecha, "%Y-%m-%d").date() if a.fecha else date.today()
    a.logs.mkdir(parents=True, exist_ok=True)
    run_id = uuid.uuid4().hex[:8]

    fmt = logging.Formatter(f"%(asctime)s [%(levelname)s] [{run_id}] %(message)s")
    log.setLevel(logging.DEBUG)
    for h in (logging.FileHandler(a.logs / "ejecucion.log", encoding="utf-8"),
              logging.StreamHandler(sys.stdout)):
        h.setFormatter(fmt)
        log.addHandler(h)
    auditoria = open(a.logs / "auditoria.jsonl", "a", encoding="utf-8")

    def audit(**kw):
        kw = dict(run_id=run_id, ts=datetime.now(timezone.utc).isoformat(), **kw)
        auditoria.write(json.dumps(kw, ensure_ascii=False) + "\n")
        auditoria.flush()

    log.info("=== INICIO run=%s fecha=%s host=%s usuario=%s python=%s so=%s",
             run_id, hoy, socket.gethostname(), getpass.getuser(),
             platform.python_version(), platform.platform())
    log.info("tabla=%s salida=%s", a.tabla, a.salida)
    try:
        filas = leer_tabla(a.tabla)
    except Exception:
        log.exception("No se pudo leer la tabla")
        return 2
    log.info("%d insumos a procesar", len(filas))

    cuentas = {"ok": 0, "existente": 0, "error": 0}
    for f in filas:
        variable, dia = limpiar(f["variable"]), carpeta_dia(hoy)
        base = {"fecha_ejecucion": str(hoy), "variable": f["variable"],
                "insumo": f["insumo"], "fila_excel": f["fila"],
                "filename_type": f["tipo"], "filename": f["filename"]}
        # Candidatas: fijo -> URL tal cual; variable -> fecha de hoy y días previos.
        if f["tipo"] == "variable":
            if not FECHA_RE.search(f["url"]):
                log.warning("[%s] fila %d es 'variable' pero la URL no contiene YYYYMMDD",
                            f["filename"], f["fila"])
            cand = [(url_para_fecha(f["url"], hoy - timedelta(days=k)), hoy - timedelta(days=k))
                    for k in range(a.retroceso + 1)]
        else:
            cand = [(f["url"], None)]

        resultado = None
        for url, fecha_url in cand:
            log.info("[%s/%s] GET %s", f["variable"], f["filename"], url)
            try:
                datos, info = descargar(url, a.timeout, a.reintentos)
            except urllib.error.HTTPError as e:
                log.warning("[%s] HTTP %s en %s", f["filename"], e.code, url)
                audit(**base, url_solicitada=url, resultado="http_error", http=e.code)
                if e.code == 404 and len(cand) > 1:
                    continue      # probar día anterior
                resultado = "error"
                break
            except Exception as e:
                log.error("[%s] fallo de red en %s: %s", f["filename"], url, e)
                audit(**base, url_solicitada=url, resultado="error_red", error=repr(e))
                resultado = "error"
                break

            if not info["content_type"].startswith("image/") or not datos:
                log.error("[%s] respuesta no es imagen (content-type=%s, %d bytes)",
                          f["filename"], info["content_type"], len(datos))
                audit(**base, url_solicitada=url, resultado="no_es_imagen",
                      **{k: info[k] for k in ("http", "content_type")}, bytes=len(datos))
                resultado = "error"
                break

            ext = extension(url, info["content_type"])
            destino = a.salida / variable / dia / (
                f"{limpiar(f['filename'])}_{hoy:%Y%m%d}{ext}")
            sha = hashlib.sha256(datos).hexdigest()
            if destino.exists() and not a.forzar:
                log.info("[%s] ya existe %s (se omite; use --forzar)", f["filename"], destino)
                audit(**base, url_solicitada=url, resultado="existente",
                      ruta=str(destino), sha256=sha, bytes=len(datos))
                resultado = "existente"
                break
            guardar_atomico(destino, datos)
            if fecha_url and fecha_url != hoy:
                log.warning("[%s] la imagen de hoy no existía; se usó la del %s",
                            f["filename"], fecha_url)
            log.info("[%s] OK %s (%d bytes, sha256=%s…)", f["filename"], destino,
                     len(datos), sha[:12])
            audit(**base, url_solicitada=url, resultado="ok", ruta=str(destino),
                  bytes=len(datos), sha256=sha,
                  fecha_en_url=str(fecha_url) if fecha_url else None, **info)
            resultado = "ok"
            break
        else:
            log.error("[%s] sin imagen disponible tras probar %d fecha(s)",
                      f["filename"], len(cand))
            resultado = "error"
        cuentas[resultado] += 1

    log.info("=== FIN run=%s ok=%d existentes=%d errores=%d",
             run_id, cuentas["ok"], cuentas["existente"], cuentas["error"])
    audit(evento="resumen", **cuentas)
    auditoria.close()
    return 1 if cuentas["error"] else 0


if __name__ == "__main__":
    sys.exit(main())
