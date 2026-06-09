# CHANGELOG v16.4 — Corrección de 3 bugs de fórmula en Excel

Fecha: 30 de mayo de 2026
Módulos: office/excel_controller.py
Tipo: corrección de bugs de fórmula detectados por auditoría con recálculo real

Una auditoría funcional generó las 95 plantillas (Word, Excel, PowerPoint) y
recalculó las 1.317 fórmulas de Excel con el motor de LibreOffice (forzando
recálculo). 92 de 95 plantillas eran 100% correctas. Esta versión corrige las
3 plantillas de Excel con bugs de fórmula reales (que también fallaban en
Microsoft Excel, no eran artefactos del recalculador).

## 1. break_even — dos correcciones

- Celda C13 ("Ingreso para cubrir costos fijos"): la nota descriptiva empezaba
  con "=" ("= Costos Fijos (punto exacto)"), por lo que Excel la interpretaba
  como fórmula y daba error. Se eliminó el "=" inicial → ahora es texto.
- Hoja Escenarios: la variable `pe_formula` traía un "=" inicial y se insertaba
  dentro de `=ROUND(...)`, generando `=ROUND(='Break-Even'!B11*B5,0)` (doble
  igual = error de sintaxis, en cascada por toda la tabla). Se quitó el "="
  inicial de `pe_formula` → ahora `=ROUND('Break-Even'!B11*B5,0)`.

## 2. roi_calculator — dos referencias a la fila de encabezados

- "Beneficio Total Bruto" usaba `=B13*B9`, pero B13 es la celda del encabezado
  ("Proyecto A", texto) → texto × número = #VALUE!. Corregido a `=B7*B9`
  (Beneficio mensual × vida útil en meses). Verificado: 45.000 × 60 = 2.700.000.
- "Meses para recuperar" usaba `=IFERROR(B6/B13,0)` (B13 encabezado) → quedaba
  enmascarado en 0 por el IFERROR. Corregido a `=IFERROR(B6/B14,0)` (Inversión /
  Beneficio neto mensual). Verificado: 500.000 / 37.000 = 13,5 meses.

## 3. balance_sheet — nombre de hoja y filas de la hoja Ratios

- Las 4 fórmulas de la hoja Ratios referenciaban una hoja `Balance!`, pero la
  hoja real se llama `Balance General` → `#NAME?` (la única visible) y el resto
  enmascarado en 0 por IFERROR (ratios silenciosamente en 0).
- Además, las filas referenciadas estaban corridas: `B5:B8` incluía un
  encabezado y omitía un activo, y `B14` apuntaba a la depreciación en vez del
  pasivo circulante.
- Solución robusta: las fórmulas ahora usan las **variables de fila dinámicas**
  del propio builder (`ac_total` = total activo circulante, `pc_total` = total
  pasivo circulante, `inv_row` = inventario, `pas_total`/`act_total` = totales
  en columna D) y el nombre de hoja correcto entre comillas (`'Balance General'`).
  Verificado: Liquidez Corriente = 2,07 (292.000/141.000) y Capital de Trabajo
  = 151.000.

## Auditoría posterior a las correcciones

- 95/95 plantillas generan un archivo válido y abrible (Word 38, Excel 39, PPT 18).
- **39/39 Excel recalculan sin un solo error** con el motor real (antes: 36/39).
- Suite completa de pruebas: **327/327 tests pasan** (sin regresiones).
- Compilación de todos los .py sin errores de sintaxis.

## Observación (fuera del alcance de esta corrección)

- roi_calculator: la fila "VPN (tasa anual 12%)" usa `REPT(...)` para simular el
  flujo de caja dentro de `NPV`, lo cual no produce un VPN correcto (queda
  enmascarado en 0 por IFERROR, no genera error). Se recomienda rediseñar esa
  fórmula con una columna auxiliar de flujos. No se tocó por no ser un bug que
  rompa el archivo y para no cambiar el comportamiento sin solicitarlo.
