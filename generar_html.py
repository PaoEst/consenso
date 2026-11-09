#!/usr/bin/env python3  # Línea "shebang": permite ejecutar el archivo directamente con Python 3 en Linux/macOS
"""Genera un HTML "Seguimiento Drivers Climáticos" con las imágenes de Lista_de_Insumos.xlsx incrustadas.

Usa las mismas reglas de consulta que descargar_imagenes.py (fecha de hoy en las URL "variable",
retroceso de días con 404, reintentos, validación de imagen y logs de auditoría), pero en lugar de
guardar cada imagen en una carpeta la incrusta (base64) dentro de un único archivo HTML.

Salida:  <salida>/Seguimiento_Drivers_Climaticos_<YYYYMMDD>.html   (se sobrescribe si ya existe)

Distribución de la página:
  - Un título central por cada Variable (en el orden de la tabla).
  - Columna izquierda: filas con Carpeta = Seguimiento.
  - Columna derecha:   filas con Carpeta = Pronostico o Seguim_Pronos.
  - Si una Variable tiene un solo tipo de información, sus mapas van en UNA columna centrada con el
    título que corresponde: Seguimiento, Pronóstico o Seguimiento-Pronóstico.
  - Al inicio hay un índice con un enlace a cada Variable; cada Variable y un botón flotante permiten volver al índice.
  - El título de cada mapa es la columna Insumo. Solo se incluyen filas con Descargar = "si".
"""
import argparse  # Para leer los argumentos de la línea de comandos
import base64  # Para incrustar las imágenes dentro del HTML (data URI)
import getpass  # Para registrar el usuario que ejecuta el script (auditoría)
import hashlib  # Para calcular el SHA-256 de cada imagen (auditoría)
import html  # Para escapar los textos de la tabla antes de ponerlos en el HTML
import json  # Para escribir los registros de auditoría en formato JSON
import logging  # Para el log de ejecución (archivo + pantalla)
import platform  # Para registrar versión de Python y sistema operativo
import re  # Para limpiar los nombres al crear identificadores de enlace
import socket  # Para registrar el nombre del equipo (auditoría)
import sys  # Para stdout y el código de salida del programa
import urllib.error  # Excepciones HTTP/red de urllib
import unicodedata  # Para quitar acentos al crear identificadores de enlace
import uuid  # Para generar un identificador único por ejecución
from datetime import date, datetime, timedelta, timezone  # Manejo de fechas y horas
from pathlib import Path  # Manejo de rutas de forma portable

import descargar_imagenes as di  # Reutiliza las funciones del script de descargas (lectura de tabla, GET con reintentos, etc.)

BASE = Path(__file__).resolve().parent  # Carpeta donde vive este script (rutas por defecto)
TITULO = "Seguimiento Drivers Climáticos"  # Título principal de la página
IZQUIERDA = {"seguimiento"}  # Valores de Carpeta que van en la columna izquierda
ETIQUETAS = {"seguimiento": "Seguimiento", "pronostico": "Pronóstico", "pronóstico": "Pronóstico", "seguim_pronos": "Seguimiento-Pronóstico"}  # Título que corresponde a cada tipo de Carpeta
DERECHA = {"pronostico", "pronóstico", "seguim_pronos"}  # Valores de Carpeta que van en la columna derecha (junto a los pronósticos)

CSS = """
:root{--bg:#f4f6f8;--card:#fff;--txt:#1d2733;--sub:#5b6b7b;--line:#d8dee4;--acc:#0b5cad;--warn:#b45309}
@media (prefers-color-scheme:dark){:root{--bg:#10151b;--card:#1a222b;--txt:#e6ebf0;--sub:#9aa9b8;--line:#2b3744;--acc:#6ab0f3;--warn:#f0b35a}}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--txt);font:16px/1.4 system-ui,-apple-system,"Segoe UI",Roboto,sans-serif}
header{padding:28px 16px 8px;text-align:center}
h1{margin:0;font-size:2rem}
.meta{color:var(--sub);font-size:.9rem;margin-top:6px}
main{max-width:1500px;margin:0 auto;padding:0 16px 40px}
html{scroll-behavior:smooth}
@media (prefers-reduced-motion:reduce){html{scroll-behavior:auto}}
section{scroll-margin-top:12px}
h2{position:relative;text-align:center;margin:40px 0 12px;padding:10px 96px;background:var(--acc);color:#fff;border-radius:8px;font-size:1.4rem}
.cols{display:grid;grid-template-columns:1fr 1fr;gap:16px;align-items:start}
.col h4{margin:0 0 8px;color:var(--sub);font-size:.85rem;text-transform:uppercase;letter-spacing:.06em;text-align:center}
figure{margin:0 0 16px;padding:12px;background:var(--card);border:1px solid var(--line);border-radius:8px}
figure h3{margin:0 0 8px;font-size:1rem;text-align:center}
figure img{display:block;width:80%;max-width:560px;height:auto;margin:0 auto}
.nota{margin-top:6px;font-size:.8rem;color:var(--warn)}
.fallo{padding:24px 12px;text-align:center;color:var(--warn);border:1px dashed var(--warn);border-radius:6px}
#indice{margin:20px 0 0;padding:14px 16px;background:var(--card);border:1px solid var(--line);border-radius:8px}
#indice h2{all:unset;display:block;text-align:center;font-weight:700;font-size:1.1rem;margin-bottom:10px}
#indice ul{list-style:none;margin:0;padding:0;display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:8px}
#indice a{display:block;text-align:center;padding:8px 12px;border:1px solid var(--line);border-radius:6px;color:var(--acc);text-decoration:none;font-weight:600}
#indice a:hover{background:var(--acc);color:#fff}
.volver{position:absolute;right:12px;top:50%;transform:translateY(-50%);color:#fff;font-size:.8rem;font-weight:400;text-decoration:none;opacity:.9}
.volver:hover{text-decoration:underline}
.subir{position:fixed;right:16px;bottom:16px;padding:10px 14px;background:var(--acc);color:#fff;border-radius:999px;text-decoration:none;font-weight:600;box-shadow:0 2px 8px rgba(0,0,0,.3)}
@media print{.subir,.volver{display:none}}
.cols.uno{grid-template-columns:minmax(0,640px);justify-content:center}
@media (max-width:900px){.cols{grid-template-columns:1fr}h2{padding:10px 12px}.volver{position:static;display:block;transform:none;margin-top:4px}}
"""  # Estilos de la página (claro/oscuro; en pantallas angostas las columnas se apilan)


def slug(texto: str, usados: set) -> str:  # Crea un identificador de enlace único a partir del nombre de una Variable
    base = unicodedata.normalize("NFKD", texto).encode("ascii", "ignore").decode("ascii").lower()  # Quita acentos y pasa a minúsculas
    base = "var-" + (re.sub(r"[^a-z0-9]+", "-", base).strip("-") or "sin-nombre")  # Solo letras, números y guiones, con prefijo "var-"
    ident, n = base, 2  # Identificador candidato y contador para repetidos
    while ident in usados:  # Si ya existe uno igual...
        ident, n = f"{base}-{n}", n + 1  # ...se le agrega -2, -3, ...
    usados.add(ident)  # Se registra como usado
    return ident  # Identificador único


def lado_de(carpeta: str) -> str:  # Decide en qué columna va una fila según su Carpeta
    n = di.norm(carpeta)  # Normaliza el texto (minúsculas, sin espacios extra)
    if n in IZQUIERDA:  # Seguimiento...
        return "izq"  # ...van a la izquierda
    if n not in DERECHA:  # Si no es un valor conocido...
        di.log.warning("Carpeta desconocida %r: se coloca en la columna derecha", carpeta)  # ...se avisa en el log
    return "der"  # Pronostico (y cualquier otro valor) va a la derecha


def obtener_imagen(f, hoy, a, audit, base):  # Descarga una imagen con las reglas de siempre y la devuelve en memoria
    if f["tipo"] == "variable":  # URL que cambia cada día
        if not di.FECHA_RE.search(f["url"]):  # Si la URL no contiene una fecha YYYYMMDD...
            di.log.warning("[%s] fila %d es 'variable' pero la URL no contiene YYYYMMDD", f["filename"], f["fila"])  # ...se avisa
        cand = [(di.url_para_fecha(f["url"], hoy - timedelta(days=k)), hoy - timedelta(days=k))  # Candidatas: hoy y días previos
                for k in range(a.retroceso + 1)]
    else:  # URL fija
        cand = [(f["url"], None)]  # Una sola candidata: la URL tal cual
    motivo = "sin imagen disponible"  # Motivo del fallo (si ninguna candidata sirve)
    for url, fecha_url in cand:  # Prueba cada URL candidata en orden
        di.log.info("[%s/%s/%s] GET %s", f["carpeta"], f["variable"], f["filename"], url)  # Registra la descarga que se intenta
        try:  # Intenta descargar
            datos, info = di.descargar(url, a.timeout, a.reintentos)  # GET con reintentos
        except urllib.error.HTTPError as e:  # El servidor devolvió un error HTTP
            di.log.warning("[%s] HTTP %s en %s", f["filename"], e.code, url)  # Lo registra en el log
            audit(**base, url_solicitada=url, resultado="http_error", http=e.code)  # Y en la auditoría
            motivo = f"HTTP {e.code}"  # Guarda el motivo para mostrarlo en la página
            if e.code == 404 and len(cand) > 1:  # 404 en fila variable: la imagen del día aún no existe
                continue  # Prueba con el día anterior
            break  # Otro error HTTP: se da por fallida
        except Exception as e:  # Fallo de red u otro error inesperado
            di.log.error("[%s] fallo de red en %s: %s", f["filename"], url, e)  # Lo registra en el log
            audit(**base, url_solicitada=url, resultado="error_red", error=repr(e))  # Y en la auditoría
            motivo = "error de red o certificado"  # Motivo para mostrar en la página
            break  # Se da por fallida
        if not info["content_type"].startswith("image/") or not datos:  # Valida que sea realmente una imagen y no esté vacía
            di.log.error("[%s] respuesta no es imagen (content-type=%s, %d bytes)", f["filename"], info["content_type"], len(datos))  # Registra el problema
            audit(**base, url_solicitada=url, resultado="no_es_imagen", http=info["http"], content_type=info["content_type"], bytes=len(datos))  # Auditoría
            motivo = "la respuesta no es una imagen"  # Motivo para mostrar en la página
            break  # Se da por fallida
        sha = hashlib.sha256(datos).hexdigest()  # Huella SHA-256 del contenido descargado
        if fecha_url and fecha_url != hoy:  # Si se usó la imagen de un día anterior...
            di.log.warning("[%s] la imagen de hoy no existía; se usó la del %s", f["filename"], fecha_url)  # ...se deja constancia
        di.log.info("[%s] OK (%d bytes, sha256=%s…)", f["filename"], len(datos), sha[:12])  # Registra la descarga exitosa
        audit(**base, url_solicitada=url, resultado="ok", bytes=len(datos), sha256=sha, fecha_en_url=str(fecha_url) if fecha_url else None, **info)  # Auditoría completa
        uri = "data:%s;base64,%s" % (info["content_type"], base64.b64encode(datos).decode("ascii"))  # Imagen incrustada como data URI
        return {"ok": True, "uri": uri, "fecha_url": fecha_url, "bytes": len(datos)}  # Devuelve la imagen lista para el HTML
    di.log.error("[%s] sin imagen disponible tras probar %d fecha(s)", f["filename"], len(cand))  # Registra el fallo definitivo
    return {"ok": False, "motivo": motivo}  # Devuelve el motivo para mostrar el aviso en la página


def figura(f, img, hoy) -> str:  # Construye el bloque HTML de un mapa (título + imagen o aviso)
    titulo = html.escape(f["insumo"] or f["filename"])  # Título del mapa = columna Insumo (escapado)
    if img["ok"]:  # Si la imagen se pudo obtener...
        cuerpo = f'<img src="{img["uri"]}" alt="{titulo}">'  # ...se incrusta
        if img["fecha_url"] and img["fecha_url"] != hoy:  # Si es de un día anterior...
            cuerpo += f'<div class="nota">Imagen del {img["fecha_url"]:%d-%m-%Y} (la de hoy aún no estaba disponible)</div>'  # ...se avisa
    else:  # Si no se pudo obtener...
        cuerpo = f'<div class="fallo">Imagen no disponible ({html.escape(img["motivo"])})</div>'  # ...se muestra un aviso
    return f"<figure><h3>{titulo}</h3>{cuerpo}</figure>"  # Bloque completo del mapa


def columna(nombre: str, figuras: list) -> str:  # Construye una columna (izquierda o derecha) de una Variable
    contenido = "".join(figuras)  # Los mapas de la columna, uno debajo de otro
    return f'<div class="col"><h4>{nombre}</h4>{contenido}</div>'  # Columna con su encabezado pequeño


def construir_html(grupos: dict, hoy: date, generado: datetime) -> str:  # Arma la página completa
    secciones, indice, usados = [], [], set()  # Secciones y entradas del índice (una por Variable) e identificadores ya usados
    for variable, lados in grupos.items():  # Recorre las Variables en el orden de la tabla
        if lados["izq"] and lados["der"]:  # Hay seguimiento y pronóstico: dos columnas
            cols = f'<div class="cols">{columna("Seguimiento", lados["izq"])}{columna("Pronóstico", lados["der"])}</div>'  # Izquierda y derecha
        else:  # Un solo tipo de información: una columna centrada
            lado = "izq" if lados["izq"] else "der"  # Lado que tiene los mapas
            if len(lados["tipos"]) == 1:  # Un único tipo (Seguimiento, Pronostico o Seguim_Pronos)...
                titulo = ETIQUETAS.get(next(iter(lados["tipos"])), "Pronóstico")  # ...su título correspondiente
            else:  # Varios tipos del mismo lado...
                titulo = "Seguimiento" if lado == "izq" else "Pronóstico"  # ...título genérico del lado
            cols = f'<div class="cols uno">{columna(titulo, lados[lado])}</div>'  # Una sola columna centrada
        ident = slug(variable, usados)  # Identificador de enlace de esta Variable
        indice.append(f'<li><a href="#{ident}">{html.escape(variable)}</a></li>')  # Entrada del índice
        secciones.append(f'<section id="{ident}"><h2>{html.escape(variable)}<a class="volver" href="#indice">↑ Índice</a></h2>{cols}</section>')  # Sección con título central y enlace para volver
    return (  # Documento HTML completo
        '<!DOCTYPE html><html lang="es"><head><meta charset="utf-8">'  # Cabecera: codificación UTF-8
        '<meta name="viewport" content="width=device-width,initial-scale=1">'  # Adaptable a celulares
        f"<title>{TITULO}</title><style>{CSS}</style></head><body>"  # Título de la pestaña y estilos
        f"<header><h1>{TITULO}</h1>"  # Título principal
        f'<div class="meta">Generado el {generado:%d-%m-%Y %H:%M}</div></header>'  # Fecha y hora
        f'<main><nav id="indice"><h2>Índice</h2><ul>{"".join(indice)}</ul></nav>'  # Índice con un enlace por Variable
        f"{''.join(secciones)}</main>"  # Secciones de las Variables
        '<a class="subir" href="#indice">↑ Índice</a></body></html>')  # Botón flotante para volver al índice


def main():  # Función principal del programa
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)  # Argumentos de línea de comandos
    ap.add_argument("--tabla", type=Path, default=BASE / "Lista_de_Insumos.xlsx")  # Ruta del Excel con los insumos
    ap.add_argument("--salida", type=Path, default=BASE / "informes")  # Carpeta donde se guarda el HTML
    ap.add_argument("--logs", type=Path, default=BASE / "logs")  # Carpeta donde se guardan los logs
    ap.add_argument("--fecha", help="AAAA-MM-DD (por defecto: hoy)")  # Permite simular otra fecha
    ap.add_argument("--retroceso", type=int, default=3, help="filas 'variable': días anteriores a probar si la URL da 404 (defecto 3)")  # Retroceso
    ap.add_argument("--timeout", type=int, default=60)  # Segundos máximos de espera por descarga
    ap.add_argument("--reintentos", type=int, default=3)  # Intentos por URL ante fallos de red/5xx
    a = ap.parse_args()  # Lee y valida los argumentos recibidos

    hoy = datetime.strptime(a.fecha, "%Y-%m-%d").date() if a.fecha else date.today()  # Fecha de trabajo: la indicada o la de hoy
    a.logs.mkdir(parents=True, exist_ok=True)  # Crea la carpeta de logs si no existe
    run_id = uuid.uuid4().hex[:8]  # Identificador corto único de esta ejecución

    fmt = logging.Formatter(f"%(asctime)s [%(levelname)s] [{run_id}] %(message)s")  # Formato de cada línea del log
    di.log.setLevel(logging.DEBUG)  # Registra todos los niveles
    for h in (logging.FileHandler(a.logs / "ejecucion.log", encoding="utf-8"), logging.StreamHandler(sys.stdout)):  # Archivo (append) y pantalla
        h.setFormatter(fmt)  # Aplica el formato al manejador
        di.log.addHandler(h)  # Registra el manejador en el logger
    auditoria = open(a.logs / "auditoria.jsonl", "a", encoding="utf-8")  # Abre el archivo de auditoría en modo append

    def audit(**kw):  # Escribe un registro de auditoría
        kw = dict(run_id=run_id, ts=datetime.now(timezone.utc).isoformat(), **kw)  # Agrega id de ejecución y marca de tiempo UTC
        auditoria.write(json.dumps(kw, ensure_ascii=False) + "\n")  # Una línea JSON por registro
        auditoria.flush()  # Fuerza el guardado en disco

    di.log.info("=== INICIO HTML run=%s fecha=%s host=%s usuario=%s python=%s so=%s", run_id, hoy, socket.gethostname(), getpass.getuser(), platform.python_version(), platform.platform())  # Contexto
    di.log.info("tabla=%s salida=%s", a.tabla, a.salida)  # Registra la tabla de entrada y la carpeta de salida
    try:  # Intenta leer el Excel
        filas = di.leer_tabla(a.tabla)  # Carga la lista de imágenes
    except Exception:  # Si falla (archivo inexistente, dañado, etc.)
        di.log.exception("No se pudo leer la tabla")  # Registra el error con su traza completa
        return 2  # Termina con código 2 (error de configuración)

    grupos, cuentas = {}, {"ok": 0, "omitido": 0, "error": 0}  # Mapas agrupados por Variable y contadores del resumen
    for f in filas:  # Procesa cada fila de la tabla
        base = {"fecha_ejecucion": str(hoy), "carpeta": f["carpeta"], "variable": f["variable"], "insumo": f["insumo"], "fila_excel": f["fila"], "filename_type": f["tipo"], "filename": f["filename"]}  # Datos comunes de auditoría
        if not f["descargar"]:  # Si la columna Descargar no dice "si"...
            di.log.info("[%s/%s/%s] Descargar=%r: no se incluye", f["carpeta"], f["variable"], f["filename"], f["descargar_valor"])  # ...se informa
            audit(**base, resultado="omitido_descargar_no", descargar=f["descargar_valor"])  # Y se registra en la auditoría
            cuentas["omitido"] += 1  # Se cuenta como omitida (no es un error)
            continue  # Pasa a la siguiente fila
        img = obtener_imagen(f, hoy, a, audit, base)  # Obtiene la imagen (con retroceso, reintentos y validación)
        cuentas["ok" if img["ok"] else "error"] += 1  # Suma al contador correspondiente
        lados = grupos.setdefault(f["variable"], {"izq": [], "der": [], "tipos": set()})  # Crea la Variable si es la primera vez que aparece
        lados[lado_de(f["carpeta"])].append(figura(f, img, hoy))  # Agrega el mapa a la columna que le corresponde
        lados["tipos"].add(di.norm(f["carpeta"]))  # Recuerda qué tipos de información tiene la Variable

    destino = a.salida / f"Seguimiento_Drivers_Climaticos_{hoy:%Y%m%d}.html"  # Ruta del HTML de salida
    pagina = construir_html(grupos, hoy, datetime.now()).encode("utf-8")  # Página completa en bytes UTF-8
    di.guardar_atomico(destino, pagina)  # Guarda el HTML (reemplaza si ya existía)
    sha = hashlib.sha256(pagina).hexdigest()  # Huella SHA-256 del HTML generado
    di.log.info("HTML generado: %s (%.1f MB, sha256=%s…)", destino, len(pagina) / 1e6, sha[:12])  # Registra el archivo generado
    audit(evento="html_generado", ruta=str(destino), bytes=len(pagina), sha256=sha)  # Auditoría del archivo generado
    di.log.info("=== FIN HTML run=%s ok=%d omitidos=%d errores=%d", run_id, cuentas["ok"], cuentas["omitido"], cuentas["error"])  # Resumen final
    audit(evento="resumen", **cuentas)  # Resumen final en la auditoría
    auditoria.close()  # Cierra el archivo de auditoría
    return 1 if cuentas["error"] else 0  # Código de salida: 1 si alguna imagen falló, 0 si todo bien


if __name__ == "__main__":  # Solo se ejecuta si el archivo se corre directamente (no si se importa)
    sys.exit(main())  # Termina el proceso con el código que devuelve main()
