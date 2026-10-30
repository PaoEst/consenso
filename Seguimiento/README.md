# Seguimiento y consenso de aportes por cuenca

Proyecto independiente para el consenso de expertos sobre los aportes (% de la media histórica) de cada cuenca.

## Estructura de carpetas
```
Seguimiento/
  README.md
  Codigo/                      <- todo lo técnico (no se comparte con los evaluadores)
    crear_graficos_consenso.py     paso 1: genera los gráficos de votación
    consolidar_votos.py            paso 3: registra votos + resultado del consenso + gráfico Consolidado
    graficar_aportes.py            (solo seguimiento, sin votación)  ·  consenso_comun.py (funciones compartidas)
    Seguimiento_Aportes.xlsx, Seguimiento_Drivers.xlsx, Pronostico_Drivers.xlsx, Consenso.xlsx   <- entradas
    requirements.txt, logs/ (bitácora de auditoría y backups/)
  Salidas/                     <- resultados; esta es la carpeta que se comparte (SharePoint)
    <año>/<mes>/<AAAA-MM-DD>/      fecha = "Fecha de realización del consenso" de Consenso.xlsx > Fechas
      index.html                   índice: enlaces a cada cuenca (votación y Consolidado)
      Consenso_<cuenca>.html       gráfico de votación
      votos/                       aquí se dejan los archivos voto_*.json
      consolidado/                 Consolidado_<cuenca>_Ronda<N>.html
```
Los scripts **siempre leen los Excel de la carpeta `Codigo/`**, sin importar desde dónde se ejecuten. Los Excel se editan ahí mismo.

## Instalación en Windows

1. **Instalar Python 3.10 o superior** desde https://www.python.org/downloads/windows/ .
   En el instalador marca **"Add python.exe to PATH"**. Verifica en una terminal nueva:
   ```
   python --version
   ```
2. **Obtener el proyecto** (elige una opción):
   - Con Git: `git clone https://github.com/PaoEst/consenso.git` y luego `cd consenso` y `git checkout seguimiento`.
   - Sin Git: en GitHub abre la rama `seguimiento`, pulsa **Code > Download ZIP** y descomprime.

   Después entra a la carpeta del código: `cd Seguimiento\Codigo`.
3. **Crear y activar un entorno virtual** (recomendado):
   ```
   python -m venv .venv
   .venv\Scripts\Activate.ps1
   ```
   - En PowerShell, si aparece un error de "ejecución de scripts deshabilitada", ejecuta una vez
     `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned` y vuelve a activar.
   - En CMD usa `.venv\Scripts\activate.bat`.
4. **Instalar las dependencias:**
   ```
   pip install -r requirements.txt
   ```
5. **Ejecutar** (con los Excel cerrados y en esta misma carpeta), ver "Flujo del consenso" más abajo.

**Problemas frecuentes**
- *"python no se reconoce como comando"*: prueba `py` en lugar de `python`, o reinstala Python marcando "Add to PATH".
- *`PermissionError` o no encuentra el Excel*: cierra los archivos en Excel y confirma que `Seguimiento_Aportes.xlsx` y `Seguimiento_Drivers.xlsx` están junto al script.
- Ante cualquier error revisa `logs\ejecucion.log`.

## Flujo del consenso (ejecutar dentro de `Codigo/`, con los Excel cerrados)
1. **Preparar los Excel** (seguimiento, pronóstico, `Consenso.xlsx` con `Fechas` y `Evaluadores_Tecnicos`).
2. **Generar los gráficos de votación:**
   ```
   python crear_graficos_consenso.py
   ```
   Crea `Salidas/<año>/<mes>/<fecha>/` con `index.html` y un `Consenso_<cuenca>.html` por cuenca con datos. Se abren con
   doble clic en el navegador (no necesitan internet); desde el índice o el selector superior se pasa de una cuenca a otra.
   Se comparte la carpeta de la corrida en SharePoint.
3. **Cada evaluador** abre el HTML de la cuenca, escribe sus **iniciales** (deben estar en `Evaluadores_Tecnicos`; si no,
   el resto queda bloqueado), el **No de evaluación** (1 = primera ronda; 2, 3... si vuelve a votar tras la discusión) y
   completa los cuadros por semana **o** el "Esperado el resto del mes" (se copia a todas las semanas; si luego edita una
   semana, el esperado se recalcula). Pulsa **GUARDAR**: se descarga un `voto_*.json` que deja en la carpeta `votos/` de la corrida.
   - *Real a la fecha*: promedio de los aportes con dato del mes del inicio del pronóstico (vacío si aún no hay datos).
   - *Esperado resto del mes*: promedio de las semanas ponderado por sus días (fechas de `Consenso.xlsx > Fechas`).
   - *Cierre del mes*: promedio de todos los días del mes (observados + pronosticados).
4. **Registrar y consolidar** (una sola corrida):
   ```
   python consolidar_votos.py            # todas las rondas existentes
   python consolidar_votos.py --ronda 2  # solo los gráficos de una ronda
   ```
   - Agrega cada voto nuevo a `Consenso.xlsx > Registros` (columnas `Mes_pronostico` e `Id_voto` al final; no duplica).
   - Escribe el resultado en la hoja **`Consolidado`**: por cada cuenca y ronda, 3 filas (`Estadístico` = Media, Mediana, Rango),
     con `No Evaluación` = ronda y `Evaluador` = iniciales de los votos considerados separadas por comas. En Rango las celdas
     son texto "mín - máx". Si se vuelve a correr, las filas se actualizan sin duplicarse.
   - Genera `consolidado/Consolidado_<cuenca>_Ronda<N>.html`: línea por evaluador, drivers climáticos, botones **Media / Mediana /
     Rango** que actualizan gráfico, recuadros (real, esperado resto, cierre), fila de totales y la tabla de trazabilidad por rondas.
   - En la ronda N cada evaluador cuenta con su voto más reciente de ronda <= N (si no volvió a votar, se mantiene el anterior).

Notas: antes de escribir, `Consenso.xlsx` se copia a `logs/backups/`; si está abierto en Excel el script avisa y no escribe.
Al guardar, el archivo pierde los valores calculados guardados de sus fórmulas (Excel los recalcula al abrirlo; los scripts
evalúan las fórmulas de `Fechas` por su cuenta). Los votos con evaluador no autorizado, mes distinto o datos inválidos se
rechazan y quedan anotados en el log.

## Solo seguimiento (`graficar_aportes.py`)
```
python graficar_aportes.py
```
Genera `Salidas/Solo_seguimiento/` con un HTML por cuenca y `index.html`.

## Reglas de los gráficos
- Cada columna de `Aportes %MH` (salvo `Fecha`) es una cuenca; sin datos -> no se grafica. Cuencas nuevas se detectan solas.
- Fechas: hoja `parámetros`, col. B, solo si col. C = "si" (inicio y fin por separado; si solo una está activa, la otra toma el valor por defecto). Si ninguna está activa: última fecha con datos y 30 días atrás. Si están activas pero son inconsistentes (ilegibles, inicio > fin o sin datos en el rango) se usa el último mes de información y se registra una advertencia en el log.
- Drivers: orden de la hoja `Prioridad`; `Remover`=si excluye, `no` incluye, vacío -> decide `Incluir` día a día.
- Solo `Presente`="si". Lugar `Colombia` por defecto; si hay filas con el nombre de la cuenca, prevalecen para esa variable y día.
- Color por `Impacto Estimado`: >1 azul, <1 ocre (más intenso a mayor diferencia), =1 o vacío gris. Texto = `Nota`. El hover de cada barra muestra variable, fechas, nota, impacto, lugar y los `comentarios` de `Driver_Clim`.
- Días consecutivos con misma variable, nota e impacto se unifican en una barra.
- Ambos paneles comparten el eje de fechas: cada dato diario de aportes se dibuja al centro de la columna de su día, igual que las barras.
- Pronóstico: filas de `Pronostico_Drivers.xlsx` con las mismas reglas. Misma variable = misma fila (seguimiento a la izquierda,
  pronóstico a la derecha), en el orden de `Prioridad` de `Seguimiento_Drivers.xlsx`; las variables que solo existan en el pronóstico
  van al final según la prioridad del pronóstico.
