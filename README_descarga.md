# Descarga diaria de imágenes

Lee `Lista_de_Insumos.xlsx` (columnas `Variable`, `filename_type`, `filename`, `Link de imagen`) y guarda:

    descargas/<Variable>/<dd_mmm_aaaa>/<filename>_<YYYYMMDD>.<ext>
    # ej.: descargas/MJO/30_sep_2026/vel_pote_200hpa_hovm_20260930.gif

- `fijo`: se descarga la URL tal cual.
- `variable`: la fecha `YYYYMMDD` de la URL se reemplaza por la de hoy; si aún no existe (404) se prueban hasta 3 días anteriores (`--retroceso`), y se deja constancia en el log.
- La extensión se toma de la URL; la fecha se agrega al final del nombre, antes de la extensión.

## Uso

    pip install -r requirements.txt
    python descargar_imagenes.py                 # hoy
    python descargar_imagenes.py --fecha 2026-09-29 --forzar

Código de salida 0 si todo fue bien, 1 si alguna imagen falló.

## Logs (`logs/`, solo se agrega, nunca se sobrescribe)

- `ejecucion.log`: log legible (host, usuario, versión, cada GET, reintentos, errores, resumen).
- `auditoria.jsonl`: un JSON por intento: URL solicitada/final, HTTP, content-type, bytes, SHA-256, ETag, Last-Modified, ruta, resultado y `run_id`.

## Programación diaria

- Linux/macOS (cron, 07:00): `0 7 * * * cd /ruta/repo && /usr/bin/python3 descargar_imagenes.py`
- Windows: Programador de tareas → acción `python.exe` con argumento `descargar_imagenes.py` y "Iniciar en" la carpeta del repo.
