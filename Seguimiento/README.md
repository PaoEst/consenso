# Seguimiento: aportes por cuenca y moduladores climáticos

Proyecto independiente. `graficar_aportes.py` genera un HTML interactivo (Plotly) por cada
cuenca de la hoja `Aportes %MH`: arriba la serie de aportes (% de la media histórica) y
abajo los drivers climáticos presentes.

## Archivos (todos en esta carpeta)
- `Seguimiento_Aportes.xlsx` (hojas `Aportes %MH`, `parámetros`) y `Seguimiento_Drivers.xlsx` (hojas `Prioridad`, `Driver_Clim`):
  el script **siempre** lee estos archivos desde su propia carpeta. Se editan aquí mismo.
- `graficar_aportes.py`, `requirements.txt`.
- `salida/` (se genera): un HTML por cuenca + `index.html`.
- `logs/ejecucion.log` (se genera): bitácora de auditoría acumulativa (run_id, SHA-256 de los Excel, ventana de fechas, advertencias, barras por cuenca).

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
5. **Ejecutar** (con los Excel cerrados y en esta misma carpeta):
   ```
   python graficar_aportes.py
   ```
   Abre `salida\index.html` con el navegador para ver los gráficos (requiere internet para cargar Plotly).

**Problemas frecuentes**
- *"python no se reconoce como comando"*: prueba `py` en lugar de `python`, o reinstala Python marcando "Add to PATH".
- *`PermissionError` o no encuentra el Excel*: cierra los archivos en Excel y confirma que `Seguimiento_Aportes.xlsx` y `Seguimiento_Drivers.xlsx` están junto al script.
- Ante cualquier error revisa `logs\ejecucion.log`.

## Uso
```
pip install -r requirements.txt
python graficar_aportes.py
```

## Reglas
- Cada columna de `Aportes %MH` (salvo `Fecha`) es una cuenca; sin datos -> no se grafica. Cuencas nuevas se detectan solas.
- Fechas: hoja `parámetros`, col. B, solo si col. C = "si" (inicio y fin por separado; si solo una está activa, la otra toma el valor por defecto). Si ninguna está activa: última fecha con datos y 30 días atrás. Si están activas pero son inconsistentes (ilegibles, inicio > fin o sin datos en el rango) se usa el último mes de información y se registra una advertencia en el log.
- Drivers: orden de la hoja `Prioridad`; `Remover`=si excluye, `no` incluye, vacío -> decide `Incluir` día a día.
- Solo `Presente`="si". Lugar `Colombia` por defecto; si hay filas con el nombre de la cuenca, prevalecen para esa variable y día.
- Color por `Impacto Estimado`: >1 azul, <1 ocre (más intenso a mayor diferencia), =1 o vacío gris. Texto = `Nota`. El hover de cada barra muestra variable, fechas, nota, impacto, lugar y los `comentarios` de `Driver_Clim`.
- Días consecutivos con misma variable, nota e impacto se unifican en una barra.
- Ambos paneles comparten el eje de fechas: cada dato diario de aportes se dibuja al centro de la columna de su día, igual que las barras.
