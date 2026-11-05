# Descarga diaria de imágenes

Lee `Lista_de_Insumos.xlsx` (columnas `Carpeta`, `Variable`, `filename_type`, `filename`, `Link de imagen`, `Descargar`) y guarda:

    descargas/<Carpeta>/<Variable>/<dd_mmm_aaaa>/<filename>_<YYYYMMDD>.<ext>
    # ej.: descargas/Seguimiento/MJO/30_sep_2026/vel_pote_200hpa_hovm_20260930.gif

- `Descargar`: `si` descarga la imagen; `no` (o vacío) la omite y queda registrado en los logs.
- **Sobrescritura:** si el archivo del día ya existe, se reemplaza. El log y la auditoría indican si se sobrescribió y guardan el SHA-256 del archivo anterior.
- `fijo`: se descarga la URL tal cual.
- `variable`: la fecha `YYYYMMDD` de la URL se reemplaza por la de hoy; si aún no existe (404) se prueban hasta 3 días anteriores (`--retroceso`), y se deja constancia en el log.
- La extensión se toma de la URL; la fecha se agrega al final del nombre, antes de la extensión.

## Uso

    pip install -r requirements.txt
    python descargar_imagenes.py                 # hoy
    python descargar_imagenes.py --fecha 2026-09-29

Código de salida 0 si todo fue bien, 1 si alguna imagen falló.

## Logs (`logs/`, solo se agrega, nunca se sobrescribe)

- `ejecucion.log`: log legible (host, usuario, versión, cada GET, reintentos, errores, resumen).
- `auditoria.jsonl`: un JSON por intento: URL solicitada/final, HTTP, content-type, bytes, SHA-256, ETag, Last-Modified, ruta, resultado y `run_id`.

## Programación diaria

- Linux/macOS (cron, 07:00): `0 7 * * * cd /ruta/repo && /usr/bin/python3 descargar_imagenes.py`
- Windows: Programador de tareas → acción `python.exe` con argumento `descargar_imagenes.py` y "Iniciar en" la carpeta del repo.

## HTML "Seguimiento Drivers Climáticos" (`generar_html.py`)

Genera un único archivo HTML con **todas las imágenes incrustadas** (no depende de internet para verse y se puede enviar por correo):

    python generar_html.py                 # hoy
    python generar_html.py --fecha 2026-11-05
    # salida: informes/Seguimiento_Drivers_Climaticos_<YYYYMMDD>.html (se sobrescribe si ya existe)

- Un título central por cada `Variable` (en el orden de la tabla).
- **Izquierda:** filas con `Carpeta` = `Seguimiento` o `Seguim_Pronos`. **Derecha:** filas con `Carpeta` = `Pronostico`.
- El título de cada mapa es la columna `Insumo`. Solo se incluyen filas con `Descargar` = `si`.
- Mismas reglas de consulta que `descargar_imagenes.py` (fecha de hoy en las URL `variable`, retroceso de días con 404, reintentos, validación de imagen) y mismos logs en `logs/`. Si una imagen falla, la página muestra "Imagen no disponible" con el motivo y el script termina con código 1.
- Es independiente de las descargas: no guarda las imágenes en carpetas. El archivo pesa varios MB (las imágenes van dentro).
