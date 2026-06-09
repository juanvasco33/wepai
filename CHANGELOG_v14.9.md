# CHANGELOG — WEP AI v14.9 (Día 1)

**Objetivo:** Excel data integration — los 5 templates de Excel más demo-ables ahora respetan los datos del usuario en lugar de generar "Producto de ejemplo 1, 2, 3...". Esto cierra el P0.1 del audit de demo-readiness.

**Fecha:** 2026-05-26
**Resultado:** 16/16 grupos de test pasando (+Test 16 con 7 sub-tests para v14.9)

---

## Lo que cambia

| Antes (v14.8) | Después (v14.9) |
|---|---|
| 35/36 builders de Excel ignoraban el input del usuario | 5 builders demo-ables leen `datos["items"]` o equivalentes |
| Pedir "inventario con mis productos Nivea, Olay, L'Oreal" → 5 "Producto de ejemplo" | Pedir lo mismo → 3 filas reales con esos productos |
| LLM no sabía qué keys emitir para Excel | SYSTEM_PROMPT documenta schema `items` por template |
| Sin tests de integración Excel ↔ user data | Test 16: 7 sub-tests E2E |

---

## Archivos modificados

### `office/excel_controller.py` (+250, −10)

#### Nuevos helpers (después de `_safe_filename`)

```python
_ITEM_KEY_SYNONYMS = {
    "codigo":       ("codigo", "code", "sku", "id", "ref", "referencia"),
    "nombre":       ("nombre", "name", "producto", "product", "descripcion", ...),
    "stock":        ("stock", "stock_actual", "cantidad", "qty", "quantity", ...),
    # ... 20+ mapeos en/es
}

def _norm_item(raw: dict) -> dict:
    """Normaliza keys: {"sku":"X","name":"Y","qty":5} → {"codigo":"X","nombre":"Y","cantidad":5}"""

def _extract_items(datos, *extra_keys) -> list[dict]:
    """Extrae items del usuario. Prueba: items, rows, + sinónimos del template."""

def _num(value, default=0):
    """Coerce robusto: "$1,234.56" → 1234.56, "50%" → 0.5, None → default."""
```

#### Wrapper `_dispatch` con introspección

```python
def _dispatch(builder, wb, titulo, secciones, detalles, instruc, datos):
    """Llama al builder pasando datos si su firma lo acepta.
    Permite migrar builders uno por uno sin tocar el dispatcher cada vez."""
```

El dispatcher de `generate_excel` ahora invoca `_dispatch(_build_X, ...)` en lugar de `_build_X(...)`. Builders con firma vieja (`def f(wb, titulo, secciones, detalles, instruccion)`) reciben los mismos 5 args. Builders refactorizados (`def f(..., datos=None)`) reciben también `datos`.

#### Builders refactorizados (5)

| Builder | Items que lee | Sinónimos aceptados |
|---|---|---|
| `_build_inventory` | productos con codigo, nombre, stock, precio_costo, precio_venta, categoria | `items`, `productos`, `inventario`, `products` |
| `_build_quotation` | servicios + cliente_nombre | `items`, `cotizacion`, `productos`, `servicios`, `lineas`, `lines` |
| `_build_payroll` | empleados con nombre, puesto, dias, salario_diario | `items`, `empleados`, `nomina`, `planilla`, `employees` |
| `_build_sales` | ventas con fecha, producto, cliente, cantidad, precio | `items`, `ventas`, `movimientos`, `transacciones` |
| `_build_cashflow` | ingresos + egresos como dict {concepto: {mes: monto}} | `ingresos`/`egresos` o `income`/`expenses`/`gastos` |

**Patrón común:** cada builder verifica si hay items del usuario. Si los hay, los usa y muestra una nota `"✓ Generado con tus datos"`. Si no, cae al sample data hardcoded (backward compat 100%) con nota `"⚠ Reemplaza con tus datos reales"`.

**Las fórmulas se ajustan dinámicamente** al número de filas: `=SUM(E5:E{lr-1})`, `=AVERAGE(G5:G{lr-1})`, etc.

### `agent/brain.py` (+45 líneas en SYSTEM_PROMPT)

Sección nueva después del schema general:

```
════════════════════════════════════════════════
CAMPOS PARA EXCEL — v14.9 CRÍTICO
════════════════════════════════════════════════
Cuando el usuario provea datos concretos (productos, empleados, ventas,
montos) para un Excel, INCLUYE en `gen_data["datos"]` el campo `items`
(o variantes según template) con los datos REALES.

INVENTARIO → items: [{codigo, nombre, categoria, stock, stock_min,
                      precio_costo, precio_venta}]
COTIZACIÓN → cliente_nombre + items: [{descripcion, unidad, cantidad,
                                        precio_venta, descuento}]
NÓMINA     → items: [{empleado, puesto, dias, salario_diario}]
VENTAS     → items: [{fecha, producto, cliente, cantidad, precio_venta}]
FLUJO CAJA → ingresos: {concepto: {mes: monto}}, egresos: idem

REGLAS:
- Si el usuario NO provee items, NO inventes — omite `items` y el builder
  caerá a sample data con nota de "reemplaza con tus datos reales".
- Si dice "dame un inventario para mi tienda" sin más detalle, pregunta UNA
  vez qué productos tiene. Si insiste "ponle ejemplos", entonces omite items.
- Para otros templates (income_statement, balance_sheet, etc.) todavía NO
  está conectado el flujo de items en v14.9. Roadmap.
```

### `tests/test_db.py` (+95 líneas)

Test 16 con 7 sub-tests:

- **16.1** Inventory integra productos del usuario (verifica que aparece "Producto Alpha", "X-1", "Beta")
- **16.2** Quotation integra cliente_nombre + items
- **16.3** Payroll integra empleados (Ana Test, Luis Test, CTO)
- **16.4** Sales integra ventas con fecha
- **16.5** Cashflow integra ingresos/egresos por mes
- **16.6** **Backward compat:** sin items → sample data sigue apareciendo
- **16.7** Helpers `_norm_item` y `_num` con casos edge

---

## Lo que sigue pendiente para demo-ready

### Día 2 (PowerPoint placeholder fix)
- [ ] Documentar campos por planner PPT en SYSTEM_PROMPT (`pitch`, `company`, `sales`, `board`, etc.)
- [ ] Cambiar fallbacks `or "[N]"` para que no se vean placeholders en slides

### Día 3 (Correction flow + cleanup)
- [ ] Versionado de archivos con sufijo `_v2`, `_v3` para no sobrescribir Office abierto
- [ ] Borrar `apply_correction` muerto + ajustar `CONTROLLERS` tuple
- [ ] Borrar `excel_validator.py` huérfano (537 líneas no usadas)
- [ ] Polling reemplaza `time.sleep(3)`

### Day 4 (demo prep)
- [ ] Demo dry-run scripted, video grabado

### Backlog (de reviews anteriores, sin urgencia para demo)

- Refactor de los 30 builders Excel restantes (puedo seguir el patrón v14.9 — son ~1h cada uno)
- Prompt caching del SYSTEM_PROMPT (C2)
- Split de controllers monolíticos (C3)
- I1-I11 menores

---

## Cómo probar el fix en tu Mac

```bash
# 1. Correr la suite completa
cd wepai
python3 tests/test_db.py
# Esperado: "All 16 test groups PASSED."

# 2. Demo manual: pedir un inventario con productos reales
# (después de configurar ANTHROPIC_API_KEY)
python3 main.py
# En el chat:
#   "Necesito un inventario para mi cafetería con Café Espresso (50 unidades,
#    $3 costo, $5 venta), Café Americano (40, $2.50, $4), Croissant (20, $1, $2.50)"
# La LLM debe emitir gen_data con items y el .xlsx generado debe mostrar
# esos 3 productos en lugar de "Producto de ejemplo 1, 2, 3..."

# 3. Verificar backward compat
# En el chat:
#   "Crea un inventario para mi negocio, con productos de ejemplo"
# Debe seguir generando los 5 "Producto de ejemplo" como antes.
```

---

*v14.9 Día 1 cierra el bug P0.1 del audit (Excel ignora input del usuario) para los 5 templates demo-ables. Los otros 30 builders siguen el patrón viejo — los priorizamos por demanda post-demo.*
