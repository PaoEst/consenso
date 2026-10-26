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

## Uso
```
pip install -r requirements.txt
python graficar_aportes.py
```

## Reglas
- Cada columna de `Aportes %MH` (salvo `Fecha`) es una cuenca; sin datos -> no se grafica. Cuencas nuevas se detectan solas.
- Fechas: hoja `parámetros`, col. B, solo si col. C = "si" (inicio y fin por separado). Si no, última fecha con datos y 30 días atrás.
- Drivers: orden de la hoja `Prioridad`; `Remover`=si excluye, `no` incluye, vacío -> decide `Incluir` día a día.
- Solo `Presente`="si". Lugar `Colombia` por defecto; si hay filas con el nombre de la cuenca, prevalecen para esa variable y día.
- Color por `Impacto Estimado`: >1 azul, <1 ocre (más intenso a mayor diferencia), =1 o vacío gris. Texto = `Nota`.
- Días consecutivos con misma variable, nota e impacto se unifican en una barra.
