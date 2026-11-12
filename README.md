# Moduladores Climáticos

Genera un **único archivo HTML** con los mapas de `Lista_de_Insumos.xlsx` incrustados (no necesita internet para verse y se puede enviar por correo).

## Uso

    pip install -r requirements.txt
    python generar_html.py                 # hoy
    python generar_html.py --fecha 2026-11-05
    # salida: informes/Moduladores_Climaticos_<YYYYMMDD>.html (se sobrescribe si ya existe)

Código de salida 0 si todo fue bien, 1 si alguna imagen falló (el HTML se genera igual) y 2 si no se pudo leer la tabla.

## Tabla `Lista_de_Insumos.xlsx`

Columnas usadas: `Carpeta`, `Variable`, `Insumo`, `filename_type`, `filename`, `Link de imagen`, `Descargar`.

- `Descargar` = `si` incluye el mapa; `no` (o vacío) lo omite y queda registrado en los logs.
- `filename_type` = `fijo`: se usa la URL tal cual. `variable`: la fecha `YYYYMMDD` de la URL se reemplaza por la de hoy; si aún no existe (404) se prueban hasta 3 días anteriores (`--retroceso`) y se avisa debajo del mapa.
- El título de cada mapa es la columna `Insumo`.

## Cómo se ve la página

- Título "Moduladores Climáticos", índice con un enlace a cada `Variable` y botón flotante "↑ Índice" para volver. Cada Variable tiene un título central.
- Si una Variable tiene **un solo tipo** de `Carpeta` (`Seguimiento`, `Pronostico` o `Seguim_Pronos`), sus mapas van en **filas de máximo 4** (la última fila queda centrada) bajo el título `Seguimiento`, `Pronóstico` o `Seguimiento-Pronóstico`.
- Si tiene seguimiento y pronóstico a la vez, se muestran dos columnas: izquierda `Seguimiento`, derecha `Pronostico` y `Seguim_Pronos`.
- **Zoom:** clic en un mapa lo amplía (otro clic: tamaño real; otro clic o `Esc`: cerrar). No aumenta el peso del archivo.
- Si una imagen falla, la página muestra "Imagen no disponible" con el motivo.
- En pantallas pequeñas las filas se reducen a 2 y 1 mapa.

## Logs (`logs/`, solo se agrega, nunca se sobrescribe)

- `ejecucion.log`: log legible (host, usuario, versión, cada GET, reintentos, errores, resumen).
- `auditoria.jsonl`: un JSON por intento con URL solicitada/final, HTTP, content-type, bytes, SHA-256, ETag, Last-Modified, resultado y `run_id`.

## Programación diaria

- Linux/macOS (cron, 07:00): `0 7 * * * cd /ruta/repo && /usr/bin/python3 generar_html.py`
- Windows: Programador de tareas → acción `python.exe` con argumento `generar_html.py` y "Iniciar en" la carpeta del repo.
