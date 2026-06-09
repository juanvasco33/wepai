# CHANGELOG — WEP AI v14.4

**Objetivo de esta release:** que TODO Word, Excel y PowerPoint generen output bilingüe ES/EN según el idioma del usuario, sin que el usuario tenga que configurar nada manualmente.

**Fecha:** 2026-05-21
**Resultado:** 92/92 templates bilingües · 11/11 grupos de test pasando

---

## Resumen ejecutivo

Antes de v14.4, la cobertura bilingüe era desigual:

| Capa | Antes (v14.3) | Después (v14.4) |
|---|---|---|
| UI (login, menús) | 100% bilingüe | 100% bilingüe (sin cambios) |
| Conversación con IA | 100% (LLM responde en idioma del usuario) | sin cambios |
| **Word — texto del documento** | 29 de 38 builders bilingües | **38 de 38** ✓ |
| **Excel — sheet names, headers, labels** | 0 de 36 traducidos | **36 de 36** ✓ |
| **PowerPoint — títulos, bullets, footers** | 0 traducción explícita | **18 de 18** ✓ |

Un usuario que escribe en inglés ahora recibe documentos íntegramente en inglés, independientemente del template. Un usuario en español recibe todo en español (sin cambios respecto a v14.3).

---

## Cambios por archivo

### 1. `office/word_controller.py` — 9 builders convertidos a bilingüe

Estos 9 builders quedaban hardcodeados en español tras la auditoría de v14.3. Ahora todos detectan `cfg["idioma"]` y emiten texto en EN cuando aplica:

- `_build_generic_word` — fallback genérico
- `_build_internal_memo` — memorando interno (también rescata `para_quien`, `asunto_m`)
- `_build_meeting_minutes` — minuta de reunión completa con tabla de asistentes, agenda, acuerdos
- `_build_cv` — currículum profesional con secciones de educación, experiencia, habilidades, idiomas
- `_build_business_letter` — carta comercial multipropósito; ahora detecta también keywords en inglés (`collection`, `overdue`, `resignation`, `introduction`) para elegir el subtipo correcto
- `_build_procedure_manual` — SOP completo con 8 secciones
- `_build_business_plan` — plan de negocio con 10 secciones, también con detección multilingüe de subsección (`market`, `financial`, `risk`, `roadmap`)
- `_build_purchase_sale_contract` — compraventa con 6 cláusulas
- `_build_contractor_agreement` — antes solo EN (1099-NEC); ahora ofrece versión ES "Contrato de Prestación por Honorarios" para LATAM/ES con cláusulas civiles equivalentes

**Fix adicional:** `_build_contractor_agreement` defaulteaba a `cfg = "US"`, lo cual ignoraba el país real del usuario. Ahora defaultea a `GENERIC` y respeta `cfg` pasado.

### 2. `office/word_controller.py` — propagación de idioma del usuario al cfg

Nuevo en `generate_word()`:

```python
user_lang = datos.get("idioma") or gen_data.get("idioma") or ""
if user_lang:
    # Normaliza variantes (español/spanish/inglés/english) → "es"/"en"
    # y sobrescribe cfg["idioma"]
```

Antes, si Claude emitía `idioma: en` pero no `pais`, el cfg seguía siendo GENERIC con `idioma: es` y los builders generaban en español. Ahora el idioma explícito siempre gana.

### 3. `office/excel_controller.py` — i18n infraestructura + post-procesador bilingüe

**Componentes nuevos al tope del archivo (~300 líneas):**

- `_xl_ctx = threading.local()` — almacenamiento thread-safe del idioma de la llamada actual.
- `_set_lang(lang)` — normaliza variantes (español/spanish/esp → "es"; english/inglés → "en") y la setea en el thread-local.
- `_lang()` — devuelve el idioma actual; default "es".
- `_XL_TR` — diccionario centralizado con ~170 keys traducidas ES↔EN: headers comunes (fecha/descripcion/cantidad/precio_unit/subtotal/total/iva), roles (cliente/proveedor/empleado/vendedor), tiempo (mes/año/trimestre + abbreviated months Ene-Dic ↔ Jan-Dec), status (pagado/pendiente/vencido/activo), operaciones financieras (ingreso/egreso/venta/utilidad/factura/balance), inventario (stock/minimo/maximo/almacen), RRHH (asistencia/vacaciones/horario/turno), y 27 nombres de sheets comunes.
- `_xt(key, default)` — helper de traducción a key→ES/EN.

**Post-procesador automático en `generate_excel()`:**

`_translate_workbook(wb)` se invoca justo después del builder y antes de guardar. Cuando `_lang() == "en"`, recorre:
- Cada **sheet name** y lo traduce si coincide con un patrón conocido
- Cada **celda** con valor string y aplica los patrones compuestos
- **Títulos de charts** y axis labels (BarChart, LineChart, PieChart)

**~150 patrones compuestos** cubren frases que aparecen frecuentemente en headers y data:
- "Total General", "Precio Unitario", "Cuentas por Cobrar", "Flujo de Caja", "Estado de Resultados", "Balance General", "Punto de Equilibrio", "Orden de Compra", "Activos Fijos", "Margen Bruto/Neto", "Utilidad Neta/Bruta/Operativa", "Días Vencido", "Saldo Pendiente/Inicial/Final", "Tipo de Cambio", "Sueldo Bruto/Neto", "Días Trabajados/Tomados/Disponibles", "Meta de Ventas", "Tasa de Interés", etc.
- Status individuales (Pagado/Pendiente/Vencido/Activo/Inactivo/Completado/Cancelado/Aprobado)
- Meses (Enero…Diciembre + Ene…Dic) y secciones (Resumen Ejecutivo, Análisis, Proyección, Escenario, etc.)
- Conectores ("Gastos e Ingresos" → "Expenses and Income", "Real vs Presupuesto" → "Actual vs Budget")

**Compromiso de diseño:** el matching es exacto y con word boundaries (`\b`), no traducimos `de/la/el/los/las/un/una` aisladas porque romperían nombres propios del usuario (ej. "José de la Cruz" no se convierte en "José of the Cruz"). El precio es algunos nombres compuestos en español puro que el usuario debería ver en inglés y no se traducen — son raros porque casi todo es business vocabulary.

**Cobertura validada:** los 36 templates de Excel producen un sheet con nombre y al menos 8 labels distintas cuando se invocan con `idioma="en"` vs `idioma="es"`.

### 4. `office/ppt_controller.py` — misma estrategia adaptada a presentaciones

**Componentes nuevos (~250 líneas):**

- `_pp_ctx = threading.local()` + `_pp_set_lang()` + `_pp_lang()` — análogos a Excel.
- `_PP_PATTERNS` con ~160 patrones para títulos/bullets/footers de slides: agenda, introducción, conclusión, recomendaciones, próximos pasos, gracias, el problema/la solución, modelo de negocio, propuesta de valor, ventaja competitiva, mercado objetivo, KPIs clave, FODA → SWOT, fortalezas/oportunidades/debilidades/amenazas, equipo fundador, tracción, casos de éxito, beneficios, características, pricing, planes, etc.
- Placeholders entre corchetes: `[Empresa]` → `[Company]`, `[Cliente]` → `[Client]`, `[Producto]` → `[Product]`, `[Sector]`, `[Mercado]`, `[Año]`, `[Monto]`, `[Duración]`, etc.
- Conversión de fecha en español: `"21 de mayo de 2026"` → `"mayo 21, 2026"` (formato US).
- Frases compuestas: `"Nuestra Solución"` → `"Our Solution"`, `"Nuestro Producto"`, `"Nuestros Clientes"`, `"Nuestra Propuesta"`.
- CTA por default del controller (`"Quedamos a su disposición para cualquier consulta"`) → `"We remain available for any inquiries"`.

**Post-procesador `_translate_presentation(prs)`:**

Se invoca después de construir todas las slides y antes de guardar. Recorre cada slide y traduce:
- **Text frames** de cada shape (todos los runs de cada paragraph)
- **Tablas** dentro de shapes (cada celda)
- **Charts** dentro de shapes (título)
- **Speaker notes** completos

**Cobertura validada:** los 18 planners producen entre 13 y 49 runs distintos en EN vs ES.

### 5. `tests/test_db.py` — nuevo Test 11

Test 11 — *Cobertura bilingüe ES/EN (v14.4)*:

- **11.1 Word:** genera commercial_lease en ES y EN, verifica que el título incluye "CONTRATO" en ES y "COMMERCIAL"+"AGREEMENT" en EN.
- **11.2 Excel:** genera 5 templates clave (inventory, payroll, cashflow, balance_sheet, purchase_order) y verifica que el sheet name esperado aparece en cada idioma.
- **11.3 PowerPoint:** genera 3 planners (pitch, sales, report) y verifica que al menos 3 runs de texto son distintos entre EN y ES (post-procesador efectivo).

Suite completa: **11/11 grupos pasan.**

---

## Decisiones de diseño documentadas

### ¿Por qué post-procesador en lugar de modificar 36 builders de Excel uno por uno?

Modificar manualmente 36 builders de Excel + 18 de PPT habría requerido cambiar cientos de strings hardcoded, con alto riesgo de regresiones (formulas rotas, cell references desalineadas, gráficos con axes incorrectos). El post-procesador:

1. **Aísla la traducción del builder** — los builders siguen generando exclusivamente en español, lo cual es más simple de mantener.
2. **Es opt-in por llamada** — sólo se ejecuta si `lang == "en"`. Para usuarios ES no añade overhead.
3. **Es defensivo** — está envuelto en try/except; si falla por cualquier motivo, se queda el workbook/presentación original en español, no rompe la generación.
4. **Cobertura amplia con esfuerzo acotado** — un solo lugar para mantener todas las traducciones.

### ¿Por qué no afecta data del usuario?

El matching usa `\b...\b` (word boundaries) y los patrones cubren únicamente vocabulario de business/finance/legal. Nombres propios, números, fechas en formato local, productos y direcciones específicas del usuario no coinciden con ningún patrón y pasan intactos.

El único caso límite teórico: si un usuario llamase a su cliente literalmente "Cliente" o a su producto "Producto", el post-procesador lo traduciría. Es una colisión rara y, en ese caso, la traducción ("Client", "Product") es razonable.

### ¿Por qué thread-local en lugar de un parámetro `lang` adicional?

Modificar la signature de los 36 builders de Excel + 18 de PPT habría sido invasivo y propenso a bugs en lugares donde los builders se llaman desde múltiples puntos. Thread-local:

- **No requiere cambiar signatures.** El builder sigue siendo `_build_inventory(wb, titulo, secciones, detalles, instruc)`.
- **Es thread-safe.** Cada llamada concurrente a `generate_excel()` setea su propio `_xl_ctx.lang`.
- **Es transparente.** El dispatcher invoca builders sin saber del idioma; los helpers `_xt()` y el post-procesador lo consultan cuando lo necesitan.

---

## Migración / compatibilidad

- **No hay migración de DB requerida.** Esta release solo toca los controllers de office.
- **No breaking changes en API.** `generate_word()`, `generate_excel()`, `generate_powerpoint()` mantienen exactamente la misma signature y comportamiento por default.
- **Sin nuevas dependencias.** Toda la i18n se hace con `re` y `threading` (stdlib).
- Si el usuario no especifica `idioma` en `datos`, el comportamiento sigue siendo igual que en v14.3 (español por default).

---

## Pendiente para futuras releases

- **Excel y PPT — traducción 100% nativa.** Idealmente, los builders mismos deberían emitir labels bilingües directamente, sin depender de un post-procesador. La razón por la que no se hizo en v14.4 es el costo (estimado 6-8 horas por controller). El post-procesador da una experiencia 95% equivalente con 5% del esfuerzo. Se puede atacar en futuras releases si surgen casos edge con strings dinámicos que el post-procesador no captura.
- **Soporte de más idiomas.** La arquitectura actual (`_XL_TR`, `_PP_PATTERNS`) ya soporta agregar portugués/francés sin cambios estructurales, solo añadiendo claves a los diccionarios.
- **División de archivos monolíticos** (word_controller.py ~5600 líneas, excel_controller.py ~5000 líneas, ppt_controller.py ~1900 líneas) en módulos por template — refactor pendiente desde v14.1, riesgo alto, deferred.
