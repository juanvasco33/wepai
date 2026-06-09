# CHANGELOG v15.16.1 — Auditoría Excel: enrutamiento + auto-actualización permanente en hojas

Fecha: 28 de mayo de 2026
Módulos: `office/excel_controller.py`, `office/legal_config.py`, `agent/brain.py`
Tipo: corrección de enrutamiento + extensión de la persistencia a Excel

---

## Resumen

Auditoría del módulo Excel con la misma metodología aplicada a Word. Se
resolvieron dos clases de problemas: (1) tres builders inalcanzables por
defectos de enrutamiento, y (2) la auto-actualización autónoma NO llegaba a las
hojas de cálculo porque Excel lee de una tabla de tasas (`EXCEL_FISCAL`) que no
pasaba por el overlay permanente introducido en v15.16.0.

---

## 1. Tres builders Excel inalcanzables (CRÍTICO)

### Problema
`kpi_tracker`, `presupuesto_anual` y `ratios_financieros` existían en
`EXCEL_TEMPLATES` y funcionaban si se invocaban directamente, pero **ninguna
ruta de usuario llegaba a ellos**:
- `brain.py` no los anunciaba al LLM → el modelo nunca emitía su `template_id`.
- El fallback por keywords los robaba otro builder: "liquidez/ratios" caía a
  `cashflow`, "kpi/métricas" a `dashboard`, "presupuesto anual" a `budget`.

Resultado: eran código muerto a pesar de estar completos y probados.

### Solución
- Se añadieron los tres a la lista de IDs Excel del SYSTEM_PROMPT de `brain.py`,
  con descripción que los distingue de sus equivalentes simples.
- Se añadieron ramas de keyword propias en el dispatcher, colocadas ANTES de
  las ramas genéricas que los robaban, con términos específicos:
  - `ratios_financieros`: "ratios financieros", "razones financieras",
    "solvencia", "rentabilidad", "prueba ácida"... (antes de `cashflow`).
  - `presupuesto_anual`: "presupuesto anual detallado", "12 meses",
    "real vs plan"... (antes del `budget` simple).
  - `kpi_tracker`: "seguimiento de kpi", "tablero de kpi", "metas y
    cumplimiento", "semáforo de indicadores"... (antes de `dashboard`).

### Verificación
Los tres ahora son alcanzables por `template_id` explícito Y por keywords.
Los builders originales (cashflow, budget simple, dashboard) siguen
enrutándose correctamente — sin regresión.

---

## 2. La auto-actualización NO llegaba a las hojas Excel (CRÍTICO)

### Problema
Excel consume sus tasas numéricas (IVA decimal, deducciones de nómina,
retenciones) desde `EXCEL_FISCAL`, un diccionario SEPARADO de `LEGAL_CONFIG`.
El overlay permanente de v15.16.0 se aplicaba a `LEGAL_CONFIG` (textos legales,
usados por Word) pero `get_excel_fiscal()` leía `EXCEL_FISCAL` directo, sin
overlay. Consecuencia: aunque el verifier actualizara el IVA, **las hojas de
cálculo seguían usando la tasa vieja hardcodeada**. La auto-actualización solo
funcionaba para Word.

### Solución
`get_excel_fiscal()` ahora aplica el overlay persistente sobre las tasas de
Excel. Como el verifier guarda el IVA como string ("17%") y Excel necesita el
decimal (0.17), se hace la conversión aquí. También se propagan `iva_label` y
`autoridad_fiscal`. Los helpers `iva_rate`, `iva_label`, `currency_symbol` y
`currency_code` heredan automáticamente el overlay.

### Verificación (end-to-end)
Probado el ciclo completo en una cotización real:
1. Cotización MX generada → celda `Config!B4` (tasa IVA) = **0.16**.
2. Verifier detecta cambio IVA 16% → 17% y lo persiste.
3. Reinicio simulado.
4. Nueva cotización MX → celda `Config!B4` = **0.17**.

El cambio llega hasta la celda real de la hoja de cálculo y persiste entre
sesiones, sin tocar el código fuente.

---

## Regresión
- Smoke test Excel completo: **234/234 OK, 0 crashes**.
- Huérfanos Excel restantes: **NINGUNO** (los 39 builders cubiertos en brain.py).
- Todos los `.py` compilan.

## Archivos modificados
- `office/excel_controller.py` — 3 ramas nuevas en el dispatcher de keywords.
- `office/legal_config.py` — `get_excel_fiscal()` aplica overlay + conversión
  IVA string→decimal.
- `agent/brain.py` — 3 builders Excel añadidos al SYSTEM_PROMPT.

## Estado de los tres módulos tras esta versión
- Word: enrutamiento corregido (v15.15.1) + auto-actualización permanente (v15.16.0).
- Excel: enrutamiento corregido + auto-actualización permanente en hojas (esta versión).
- PowerPoint: pendiente de auditar con la misma metodología.
