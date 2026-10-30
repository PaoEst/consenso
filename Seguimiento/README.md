# Seguimiento y consenso de aportes por cuenca

Proyecto independiente con dos flujos:

1. **Consenso de expertos** (principal): `crear_graficos_consenso.py` genera por cuenca un gráfico
   interactivo con el seguimiento (aportes + drivers) a la izquierda, el pronóstico de drivers a la
   derecha y los cuadros para que cada evaluador registre su pronóstico de aportes (S1..S5 o resto
   del mes). `consolidar_votos.py` registra los votos en `Consenso.xlsx` y genera la consolidación
   por ronda.
2. **Solo seguimiento**: `graficar_aportes.py` (gráfico de aportes + drivers de seguimiento).

## Archivos (todos en esta carpeta; los scripts siempre los leen de aquí)
- `Seguimiento_Aportes.xlsx` (`Aportes %MH`, `parámetros`), `Seguimiento_Drivers.xlsx` (`Prioridad`, `Driver_Clim`),
  `Pronostico_Drivers.xlsx` (`Priodidad`/`Prioridad`, `Driver_Clim`) y `Consenso.xlsx`
  (`Fechas`, `Evaluadores_Tecnicos`, `Registros`).
- `<año>/<mes>/<AAAA-MM-DD>/` (se genera; fecha = "Fecha de realización del consenso" de `Consenso.xlsx > Fechas`):
  `Consenso_<cuenca>.html`, `votos/` y `consolidacion/`.
- `logs/ejecucion.log`: bitácora de auditoría acumulativa de todos los scripts (run_id, SHA-256 de los Excel,
  fechas usadas, advertencias, votos aceptados/rechazados, filas escritas). `logs/backups/` guarda una copia de
  `Consenso.xlsx` antes de cada escritura.

## Instalación en Windows

1. **Instalar Python 3.10 o superior** desde https://www.python.org/downloads/windows/ .
   En el instalador marca **"Add python.exe to PATH"**. Verifica en una terminal nueva:
   ```
   python --version
   ```
2. **Obtener el proyecto** (elige una opción):
   - Con Git: `git clone https://github.com/PaoEst/consenso.git` y luego `cd consenso` y `git checkout seguimiento`.
   - Sin Git: en GitHub abre la rama `seguimiento`, pulsa **Code > Download ZIP** y descomprime.

   Después entra a la carpeta del proyecto: `cd Seguimiento`.
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

## Flujo del consenso
1. **Preparar los Excel** en esta carpeta (seguimiento, pronóstico, fechas y evaluadores en `Consenso.xlsx`).
2. **Generar los gráficos:**
   ```
   python crear_graficos_consenso.py
   ```
   Crea `<año>/<mes>/<fecha>/Consenso_<cuenca>.html` (un archivo por cuenca con datos). Se abren con doble clic
   en el navegador; no necesitan internet. Se pueden compartir copiando la carpeta de la corrida a SharePoint.
3. **Cada evaluador** abre el HTML de la cuenca, escribe sus **iniciales** (deben estar en `Evaluadores_Tecnicos`; si no,
   el resto de controles queda bloqueado), el **No de evaluación** (1 = primera ronda; 2, 3... si vuelve a votar tras la
   discusión) y completa los cuadros naranja por semana **o** el "Esperado el resto del mes" (se copia a todas las
   semanas; si luego edita una semana, el esperado se recalcula). Pulsa **GUARDAR**: se descarga un archivo
   `voto_*.json` que debe dejar en la carpeta `votos/` de la corrida.
   - *Real a la fecha*: promedio de los aportes con dato del mes del inicio del pronóstico (vacío si aún no hay datos).
   - *Esperado resto del mes*: promedio de las semanas ponderado por sus días (las fechas salen de `Fechas`).
   - *Cierre del mes*: promedio de todos los días del mes (observados + pronosticados).
4. **Registrar y consolidar** (una sola corrida, con `Consenso.xlsx` cerrado):
   ```
   python consolidar_votos.py            # todas las rondas existentes
   python consolidar_votos.py --ronda 2  # solo una ronda
   ```
   Agrega cada voto nuevo a `Consenso.xlsx > Registros` (columnas `Mes_pronostico` e `Id_voto` al final; no duplica si se
   vuelve a correr) y genera `consolidacion/Consolidacion_<cuenca>_Ronda<N>.html`: una línea por evaluador, y botones
   **Media / Mediana / Rango** para "resto del mes" y "cierre". En la ronda N cada evaluador cuenta con su voto más
   reciente de ronda <= N (si no volvió a votar, se mantiene su voto anterior). Todas las votaciones quedan en `Registros`.

Notas: al guardar, `Consenso.xlsx` pierde los valores calculados guardados de sus fórmulas (Excel los recalcula al abrirlo);
los scripts evalúan las fórmulas de `Fechas` por su cuenta. Los votos con evaluador no autorizado, mes distinto o datos
inválidos se rechazan y se anotan en el log.

## Solo seguimiento (`graficar_aportes.py`)
```
python graficar_aportes.py
```
Genera `salida/` con un HTML por cuenca y `index.html`.

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
