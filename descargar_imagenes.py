#!/usr/bin/env python3  # Línea "shebang": permite ejecutar el archivo directamente con Python 3 en Linux/macOS
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
import argparse  # Para leer los argumentos de la línea de comandos (--fecha, --salida, ...)
import getpass  # Para obtener el nombre del usuario que ejecuta el script (auditoría)
import hashlib  # Para calcular el hash SHA-256 de cada imagen (integridad/auditoría)
import json  # Para escribir los registros de auditoría en formato JSON
import logging  # Para el log de ejecución (archivo + pantalla)
import os  # Para operaciones de archivos (reemplazo atómico, extensión)
import platform  # Para registrar versión de Python y sistema operativo
import re  # Expresiones regulares (fechas en URLs, limpieza de nombres)
import socket  # Para obtener el nombre del equipo (auditoría)
import sys  # Para stdout y el código de salida del programa
import time  # Para medir duración y esperar entre reintentos
import urllib.error  # Excepciones HTTP/red de urllib
import urllib.request  # Cliente HTTP de la librería estándar (sin dependencias)
import uuid  # Para generar un identificador único por ejecución
from datetime import date, datetime, timedelta, timezone  # Manejo de fechas y horas
from pathlib import Path  # Manejo de rutas de forma portable (Windows/Linux/macOS)
from urllib.parse import urlparse  # Para separar las partes de una URL

BASE = Path(__file__).resolve().parent  # Carpeta donde vive este script (rutas por defecto)
MESES = ["ene", "feb", "mar", "abr", "may", "jun",  # Abreviaturas de meses en español,
         "jul", "ago", "sep", "oct", "nov", "dic"]  # independientes del idioma del sistema
FECHA_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")  # Busca exactamente 8 dígitos seguidos (YYYYMMDD) en una URL
EXT_POR_TIPO = {"image/png": ".png", "image/gif": ".gif", "image/jpeg": ".jpg",  # Extensión según el Content-Type,
                "image/webp": ".webp", "image/svg+xml": ".svg"}  # usada si la URL no trae extensión
USER_AGENT = "Mozilla/5.0 (compatible; descarga-imagenes-diaria/1.0)"  # Identificación enviada al servidor

log = logging.getLogger("descargas")  # Logger principal de la aplicación


# ----------------------------------------------------------------- utilidades
def carpeta_dia(d: date) -> str:  # Convierte una fecha en el nombre de carpeta del día
    return f"{d.day:02d}_{MESES[d.month - 1]}_{d.year}"  # Ej.: 2026-09-29 -> "29_sep_2026"


def limpiar(nombre: str) -> str:  # Limpia un texto para usarlo como nombre de carpeta/archivo
    """Evita separadores de ruta / caracteres inválidos en nombres de carpeta."""
    nombre = re.sub(r'[\\/:*?"<>|]', "_", str(nombre)).strip().rstrip(".")  # Reemplaza caracteres prohibidos por "_" y quita espacios/puntos finales
    return nombre or "sin_nombre"  # Si quedó vacío, usa un nombre por defecto


def norm(s) -> str:  # Normaliza un texto para comparar encabezados
    return re.sub(r"\s+", " ", str(s or "")).strip().lower()  # Minúsculas, sin espacios dobles ni extremos


def leer_tabla(ruta: Path):  # Lee el Excel y devuelve la lista de imágenes a descargar
    """Lee el Excel y devuelve filas con variable (heredada de celdas vacías)."""
    import openpyxl  # Se importa aquí para que --help funcione aunque no esté instalado
    ws = openpyxl.load_workbook(ruta, data_only=True).worksheets[0]  # Abre la primera hoja del libro
    filas = ws.iter_rows(values_only=True)  # Iterador de filas con solo los valores de las celdas
    cab = [norm(c) for c in next(filas)]  # Primera fila = encabezados, normalizados
    req = ["variable", "filename_type", "filename", "link de imagen"]  # Columnas obligatorias
    faltan = [c for c in req if c not in cab]  # Detecta las columnas obligatorias ausentes
    if faltan:  # Si falta alguna...
        raise SystemExit(f"Faltan columnas en {ruta.name}: {faltan}")  # ...termina con un mensaje claro
    idx = {c: cab.index(c) for c in req + (["insumo"] if "insumo" in cab else [])}  # Posición de cada columna (y "insumo" si existe)
    out, variable = [], None  # Lista de resultados y variable vigente (para celdas vacías)
    for n, fila in enumerate(filas, start=2):  # Recorre los datos; n = número de fila en Excel
        g = lambda k: (str(fila[idx[k]]).strip() if k in idx and fila[idx[k]] else "")  # Obtiene el texto de una columna ("" si vacía)
        if g("variable"):  # Si la celda Variable trae valor...
            variable = g("variable")  # ...pasa a ser la variable vigente (las vacías heredan la anterior)
        if not (g("filename") and g("link de imagen")):  # Si falta nombre o enlace...
            continue  # ...se omite la fila
        out.append(dict(fila=n, variable=variable or "sin_variable",  # Guarda la fila y su variable
                        insumo=g("insumo"), tipo=norm(g("filename_type")),  # Descripción y tipo (fijo/variable)
                        filename=g("filename"), url=g("link de imagen")))  # Nombre base y URL de la imagen
    return out  # Devuelve todas las imágenes a procesar


def url_para_fecha(url: str, d: date) -> str:  # Ajusta la fecha dentro de una URL "variable"
    return FECHA_RE.sub(d.strftime("%Y%m%d"), url)  # Reemplaza el YYYYMMDD de la URL por la fecha indicada


def extension(url: str, content_type: str) -> str:  # Decide la extensión del archivo final
    ext = os.path.splitext(urlparse(url).path)[1].lower()  # Extensión que trae la URL (.gif, .png, ...)
    return ext if ext else EXT_POR_TIPO.get(content_type, ".bin")  # Si no hay, se deduce del Content-Type


def descargar(url: str, timeout: int, reintentos: int):  # Descarga una URL con reintentos
    """GET con reintentos. Devuelve (bytes, info). Lanza HTTPError si 4xx."""
    ultimo = None  # Guarda el último error para relanzarlo si se agotan los intentos
    for intento in range(1, reintentos + 1):  # Repite hasta el número de reintentos
        t0 = time.monotonic()  # Marca de tiempo para medir la duración
        try:  # Intenta la petición
            req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})  # Construye la petición con User-Agent
            with urllib.request.urlopen(req, timeout=timeout) as r:  # Abre la conexión con tiempo límite
                datos = r.read()  # Lee todo el contenido (los bytes de la imagen)
                return datos, dict(  # Devuelve los bytes y metadatos para auditoría
                    http=r.status, url_final=r.geturl(), intentos=intento,  # Código HTTP, URL final tras redirecciones, nº de intento
                    content_type=(r.headers.get_content_type() or "").lower(),  # Tipo de contenido devuelto
                    last_modified=r.headers.get("Last-Modified"),  # Fecha de modificación según el servidor
                    etag=r.headers.get("ETag"),  # Identificador de versión del servidor
                    ms=int((time.monotonic() - t0) * 1000))  # Duración de la descarga en milisegundos
        except urllib.error.HTTPError as e:  # El servidor respondió con un error HTTP
            if e.code < 500:  # Errores 4xx (p. ej. 404): reintentar no sirve
                e.intentos = intento  # Anota cuántos intentos se hicieron
                raise  # Se propaga al llamador
            ultimo = e  # Errores 5xx: se recuerda y se reintenta
        except (urllib.error.URLError, TimeoutError, ConnectionError, OSError) as e:  # Fallos de red/timeout
            ultimo = e  # Se recuerda el error y se reintenta
        log.warning("Intento %d/%d falló para %s: %s", intento, reintentos, url, ultimo)  # Deja constancia del fallo en el log
        if intento < reintentos:  # Si quedan intentos...
            time.sleep(2 ** intento)  # ...espera 2, 4, 8... segundos (espera exponencial)
    raise ultimo  # Se agotaron los intentos: relanza el último error


def guardar_atomico(destino: Path, datos: bytes):  # Guarda el archivo sin dejar archivos a medias
    destino.parent.mkdir(parents=True, exist_ok=True)  # Crea las carpetas variable/día si no existen
    tmp = destino.with_name(destino.name + ".part")  # Archivo temporal junto al destino
    tmp.write_bytes(datos)  # Escribe primero en el temporal
    os.replace(tmp, destino)  # Lo renombra al nombre final de forma atómica


# ----------------------------------------------------------------------- main
def main():  # Función principal del programa
    ap = argparse.ArgumentParser(description=__doc__,  # Define los argumentos de línea de comandos
                                 formatter_class=argparse.RawDescriptionHelpFormatter)  # Respeta el formato del texto de ayuda
    ap.add_argument("--tabla", type=Path, default=BASE / "Lista_de_Insumos.xlsx")  # Ruta del Excel con los insumos
    ap.add_argument("--salida", type=Path, default=BASE / "descargas")  # Carpeta raíz donde se guardan las imágenes
    ap.add_argument("--logs", type=Path, default=BASE / "logs")  # Carpeta donde se guardan los logs
    ap.add_argument("--fecha", help="AAAA-MM-DD (por defecto: hoy)")  # Permite simular otra fecha
    ap.add_argument("--retroceso", type=int, default=3,  # Días hacia atrás a probar en filas "variable"
                    help="filas 'variable': si la URL de la fecha no existe (404), "
                         "probar hasta N días anteriores (defecto 3)")
    ap.add_argument("--timeout", type=int, default=60)  # Segundos máximos de espera por descarga
    ap.add_argument("--reintentos", type=int, default=3)  # Intentos por URL ante fallos de red/5xx
    ap.add_argument("--forzar", action="store_true",  # Bandera: si se pasa, vale True
                    help="sobrescribir si el archivo del día ya existe")
    a = ap.parse_args()  # Lee y valida los argumentos recibidos

    hoy = datetime.strptime(a.fecha, "%Y-%m-%d").date() if a.fecha else date.today()  # Fecha de trabajo: la indicada o la de hoy
    a.logs.mkdir(parents=True, exist_ok=True)  # Crea la carpeta de logs si no existe
    run_id = uuid.uuid4().hex[:8]  # Identificador corto único de esta ejecución

    fmt = logging.Formatter(f"%(asctime)s [%(levelname)s] [{run_id}] %(message)s")  # Formato de cada línea del log
    log.setLevel(logging.DEBUG)  # Registra todos los niveles
    for h in (logging.FileHandler(a.logs / "ejecucion.log", encoding="utf-8"),  # Escribe el log en archivo (modo append)
              logging.StreamHandler(sys.stdout)):  # y también lo muestra en pantalla
        h.setFormatter(fmt)  # Aplica el formato al manejador
        log.addHandler(h)  # Registra el manejador en el logger
    auditoria = open(a.logs / "auditoria.jsonl", "a", encoding="utf-8")  # Abre el archivo de auditoría en modo append

    def audit(**kw):  # Escribe un registro de auditoría
        kw = dict(run_id=run_id, ts=datetime.now(timezone.utc).isoformat(), **kw)  # Agrega id de ejecución y marca de tiempo UTC
        auditoria.write(json.dumps(kw, ensure_ascii=False) + "\n")  # Una línea JSON por registro
        auditoria.flush()  # Fuerza el guardado en disco por si el programa se interrumpe

    log.info("=== INICIO run=%s fecha=%s host=%s usuario=%s python=%s so=%s",  # Registra el contexto de la ejecución
             run_id, hoy, socket.gethostname(), getpass.getuser(),  # Id, fecha, equipo y usuario
             platform.python_version(), platform.platform())  # Versión de Python y sistema operativo
    log.info("tabla=%s salida=%s", a.tabla, a.salida)  # Registra la tabla de entrada y la carpeta de salida
    try:  # Intenta leer el Excel
        filas = leer_tabla(a.tabla)  # Carga la lista de imágenes
    except Exception:  # Si falla (archivo inexistente, dañado, etc.)
        log.exception("No se pudo leer la tabla")  # Registra el error con su traza completa
        return 2  # Termina con código 2 (error de configuración)
    log.info("%d insumos a procesar", len(filas))  # Registra cuántas imágenes se procesarán

    cuentas = {"ok": 0, "existente": 0, "error": 0}  # Contadores para el resumen final
    for f in filas:  # Procesa cada imagen de la tabla
        variable, dia = limpiar(f["variable"]), carpeta_dia(hoy)  # Nombres de carpeta de variable y de día
        base = {"fecha_ejecucion": str(hoy), "variable": f["variable"],  # Datos comunes a todos los registros de auditoría
                "insumo": f["insumo"], "fila_excel": f["fila"],  # de esta imagen: descripción y fila de origen
                "filename_type": f["tipo"], "filename": f["filename"]}  # tipo (fijo/variable) y nombre base
        # Candidatas: fijo -> URL tal cual; variable -> fecha de hoy y días previos.
        if f["tipo"] == "variable":  # URL que cambia cada día
            if not FECHA_RE.search(f["url"]):  # Si la URL no contiene una fecha YYYYMMDD...
                log.warning("[%s] fila %d es 'variable' pero la URL no contiene YYYYMMDD",  # ...se avisa en el log
                            f["filename"], f["fila"])
            cand = [(url_para_fecha(f["url"], hoy - timedelta(days=k)), hoy - timedelta(days=k))  # Lista de (URL, fecha) para hoy y días previos
                    for k in range(a.retroceso + 1)]
        else:  # URL fija
            cand = [(f["url"], None)]  # Una sola candidata: la URL tal cual

        resultado = None  # Resultado final de esta imagen (ok/existente/error)
        for url, fecha_url in cand:  # Prueba cada URL candidata en orden
            log.info("[%s/%s] GET %s", f["variable"], f["filename"], url)  # Registra la descarga que se intenta
            try:  # Intenta descargar
                datos, info = descargar(url, a.timeout, a.reintentos)  # Descarga con reintentos
            except urllib.error.HTTPError as e:  # El servidor devolvió un error HTTP
                log.warning("[%s] HTTP %s en %s", f["filename"], e.code, url)  # Lo registra en el log
                audit(**base, url_solicitada=url, resultado="http_error", http=e.code)  # Y en la auditoría
                if e.code == 404 and len(cand) > 1:  # 404 en fila variable: la imagen del día aún no existe
                    continue  # Prueba con el día anterior
                resultado = "error"  # Otro error HTTP: se da por fallida
                break  # No se prueban más candidatas
            except Exception as e:  # Fallo de red u otro error inesperado
                log.error("[%s] fallo de red en %s: %s", f["filename"], url, e)  # Lo registra en el log
                audit(**base, url_solicitada=url, resultado="error_red", error=repr(e))  # Y en la auditoría
                resultado = "error"  # Se da por fallida
                break  # No se prueban más candidatas

            if not info["content_type"].startswith("image/") or not datos:  # Valida que sea realmente una imagen y no esté vacía
                log.error("[%s] respuesta no es imagen (content-type=%s, %d bytes)",  # Registra el problema
                          f["filename"], info["content_type"], len(datos))
                audit(**base, url_solicitada=url, resultado="no_es_imagen",  # Deja constancia en la auditoría
                      **{k: info[k] for k in ("http", "content_type")}, bytes=len(datos))
                resultado = "error"  # Se da por fallida
                break  # No se prueban más candidatas

            ext = extension(url, info["content_type"])  # Extensión del archivo final
            destino = a.salida / variable / dia / (  # Ruta final: salida/variable/día/nombre
                f"{limpiar(f['filename'])}_{hoy:%Y%m%d}{ext}")  # Nombre = filename + "_YYYYMMDD" + extensión
            sha = hashlib.sha256(datos).hexdigest()  # Huella SHA-256 del contenido descargado
            if destino.exists() and not a.forzar:  # Si ya se descargó hoy y no se pidió forzar...
                log.info("[%s] ya existe %s (se omite; use --forzar)", f["filename"], destino)  # ...se informa
                audit(**base, url_solicitada=url, resultado="existente",  # Se registra en la auditoría
                      ruta=str(destino), sha256=sha, bytes=len(datos))
                resultado = "existente"  # Resultado: ya existía
                break  # No se prueban más candidatas
            guardar_atomico(destino, datos)  # Guarda la imagen en disco
            if fecha_url and fecha_url != hoy:  # Si se usó la imagen de un día anterior...
                log.warning("[%s] la imagen de hoy no existía; se usó la del %s",  # ...se deja constancia
                            f["filename"], fecha_url)
            log.info("[%s] OK %s (%d bytes, sha256=%s…)", f["filename"], destino,  # Registra la descarga exitosa
                     len(datos), sha[:12])
            audit(**base, url_solicitada=url, resultado="ok", ruta=str(destino),  # Registro de auditoría completo
                  bytes=len(datos), sha256=sha,
                  fecha_en_url=str(fecha_url) if fecha_url else None, **info)
            resultado = "ok"  # Resultado: descargada correctamente
            break  # Ya no se necesitan más candidatas
        else:  # Se ejecuta solo si el bucle terminó sin "break": ninguna candidata sirvió
            log.error("[%s] sin imagen disponible tras probar %d fecha(s)",  # Registra el fallo definitivo
                      f["filename"], len(cand))
            resultado = "error"  # Resultado: error
        cuentas[resultado] += 1  # Suma al contador correspondiente

    log.info("=== FIN run=%s ok=%d existentes=%d errores=%d",  # Resumen final en el log
             run_id, cuentas["ok"], cuentas["existente"], cuentas["error"])
    audit(evento="resumen", **cuentas)  # Resumen final en la auditoría
    auditoria.close()  # Cierra el archivo de auditoría
    return 1 if cuentas["error"] else 0  # Código de salida: 1 si hubo errores, 0 si todo bien


if __name__ == "__main__":  # Solo se ejecuta si el archivo se corre directamente (no si se importa)
    sys.exit(main())  # Termina el proceso con el código que devuelve main()
