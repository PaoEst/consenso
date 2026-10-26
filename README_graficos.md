# Gráficos de aportes y moduladores climáticos

`graficar_aportes.py` genera un HTML interactivo (Plotly) por cada cuenca de la hoja
`Aportes %MH`: arriba la serie de aportes (% de la media histórica) y abajo los
drivers climáticos presentes.

```
pip install -r requirements.txt
python graficar_aportes.py --aportes Seguimiento_Aportes.xlsx --drivers Seguimiento_Drivers.xlsx
```

Salida en `salida/` (un HTML por cuenca + `index.html`). Bitácora de auditoría en `logs/ejecucion.log`
(run_id, hash SHA-256 de entradas, ventana de fechas, advertencias, barras graficadas por cuenca).

## Reglas
- Cada columna de `Aportes %MH` (salvo `Fecha`) es una cuenca; sin datos -> no se grafica.
- Fechas: hoja `parámetros`, col. B, solo si col. C = "si" (inicio y fin por separado). Si no, última fecha con datos y 30 días atrás.
- Drivers: orden de la hoja `Prioridad`; `Remover`=si excluye, `no` incluye, vacío -> decide `Incluir` día a día.
- Solo `Presente`="si". Lugar `Colombia` por defecto; si hay filas con el nombre de la cuenca, prevalecen para esa variable y día.
- Color por `Impacto Estimado`: >1 azul, <1 ocre (más intenso a mayor diferencia), =1 o vacío gris. Texto = `Nota`.
- Días consecutivos con misma variable, nota e impacto se unifican en una barra.
