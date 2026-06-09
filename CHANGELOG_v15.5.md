# CHANGELOG — v15.5

**Fecha:** mayo 2026
**Foco:** soporte fiscal y laboral multi-país (LATAM + ES + US) en Excel, en paridad con Word.

---

## Resumen ejecutivo

Hasta v15.4, los 35 templates de Excel asumían México de facto: IVA 16% hardcoded, payroll con instituciones IMSS/ISR/INFONAVIT, freelance con ISR 35%, libro de ingresos/egresos con RESICO. Un usuario colombiano que pedía nómina recibía formato mexicano con etiquetas inaplicables. Word ya soportaba 20 países desde hacía tiempo con `LEGAL_CONFIG`. v15.5 cierra esa brecha.

Resultado: los 7 builders críticos de Excel (`payroll`, `quotation`, `sales`, `purchase_order`, `freelance_tracker`, `income_statement`, `income_expense_book`) ahora detectan el país y aplican tasas, etiquetas, instituciones y formato de moneda locales. Los 28 builders restantes reciben `cfg` por el dispatcher pero pueden ignorarlo (compatibilidad hacia atrás).

---

## Cambios por archivo

### Nuevo: `office/legal_config.py` (938 líneas)

Módulo compartido entre Word y Excel. Contiene:

- **`LEGAL_CONFIG`** — 20 países LATAM + ES + US + GENERIC, con campos legales y fiscales (leyes laborales, artículos aplicables, instituciones, IVA textual, autoridad fiscal). Extraído de `word_controller.py` sin modificar contenido.
- **`COUNTRY_ALIASES`** — variaciones de nombre ("méxico", "cdmx", "mexico", "mx") → código ISO.
- **`_get_legal_config(pais_raw)`** y **`_detect_country(gen_data, instruc, detalles)`** — extraídas tal cual de Word.
- **`EXCEL_FISCAL`** — *nuevo overlay numérico por país* con campos que Excel necesita:
  - `iva_rate` (decimal): 0.16 MX, 0.19 CO, 0.18 PE, 0.21 AR, 0.07 PA, etc.
  - `iva_label`: "IVA" / "IGV" / "ITBMS" / "ITBIS" / "ISV" / "Sales Tax"
  - `currency_format`: formato openpyxl ($#,##0.00 / S/. #,##0.00 / ₲ #,##0 / etc.)
  - `payroll_deductions`: lista `[(label, default_rate, note), ...]` por país. Ejemplos:
    - MX: IMSS · ISR · INFONAVIT
    - CO: Salud EPS · Pensión AFP · Retefuente · Solidaridad
    - PE: AFP / ONP · IR 5ta categoría · CTS
    - AR: Jubilación SIPA · Obra Social · PAMI · Ganancias
    - CL: AFP · Salud Isapre/Fonasa · AFC · Impuesto único
    - ES: Seguridad Social · IRPF
    - US: FICA Social Security · Medicare · Federal Income Tax · State Income Tax
  - `income_tax_freelance` y `income_tax_freelance_label`: retención sobre honorarios (10% MX Art.106 LISR, 8% PE 4ta cat., 13.75% CL boleta honorarios, 15% ES IRPF, etc.)
  - `tax_authority`: nombre de la autoridad fiscal local
- **Helpers de acceso**: `get_excel_fiscal(cfg)`, `iva_rate(cfg)`, `iva_label(cfg)`, `currency_format(cfg)`, `currency_code(cfg)`, `currency_symbol(cfg)`.

### Modificado: `office/word_controller.py` (–561 líneas)

Las 579 líneas de `LEGAL_CONFIG`, `COUNTRY_ALIASES`, `_get_legal_config` y `_detect_country` se reemplazaron por una importación del módulo compartido:

```python
from office.legal_config import (
    LEGAL_CONFIG, COUNTRY_ALIASES, _get_legal_config, _detect_country,
)
```

Comportamiento idéntico — todos los builders legales de Word siguen funcionando sin cambios. Verificado con smoke tests para MX/CO/AR.

### Modificado: `office/excel_controller.py`

#### Infraestructura
- Import del nuevo módulo compartido al top.
- **`_dispatch()`** ahora detecta vía `inspect.signature` si el builder acepta `datos` y/o `cfg`, y se los pasa según corresponda. Builders sin `cfg` simplemente no lo reciben (sin error).
- **`generate_excel()`**: al principio llama `_detect_country(gen_data, instruc, detalles)` para obtener `cfg`. Aplica override de idioma si el usuario lo especificó. Todas las ramas del dispatcher por keyword ahora van a través de `_dispatch(..., datos, cfg)` (antes algunas llamaban directo al builder).

#### Builders refactorizados con `cfg`
Cada uno acepta ahora `cfg=None` y lee de él:

| Builder | Cambios principales |
|---|---|
| `_build_payroll` | Deducciones dinámicas según `EXCEL_FISCAL[país].payroll_deductions`. Columnas se generan dinámicamente (4-7 columnas de deducciones según país). Tabla LISR Art. 96 se mantiene SOLO para MX (`is_mexico` flag). Para otros países la columna ISR usa `bruto × tasa Parámetros` con nota explicativa para consultar tabla oficial. Formato de moneda del país. Hoja Parámetros editable. |
| `_build_quotation` | IVA default del país (no más "16% hardcoded"). Label dinámico (IVA/IGV/ITBMS/ITBIS/Sales Tax). Moneda y formato del país en Config sheet. `id_empresa` (RFC/NIT/RUT/etc.) según país. |
| `_build_sales` | Header de columna G dinámico: "IVA (16%)" → "IGV (18%)" / "ITBMS (7%)" / "Sales Tax". Fórmula de impuesto usa la tasa del país (0 para US). Formato de moneda en celdas. Estadísticas con label local ("IGV Cobrado" no "IVA Cobrado"). |
| `_build_purchase_order` | IVA/ITBMS/ITBIS según país en subtotal y total. Moneda del país. Nota legal actualizada. |
| `_build_freelance_tracker` | Hoja Impuestos completamente localizada: tasa de retención según país (8% MX/PE, 11% CO, 13.75% CL, 15% ES), label completo ("Retención IR 4ta categoría" / "Retención en la fuente" / "Boleta de honorarios" / "Retención IRPF"). Mención a la autoridad fiscal local. |
| `_build_income_statement` | Label "ISR" reemplazado por el equivalente del país: "ISR" (MX), "Impuesto a las Ganancias" (AR), "Impuesto Sociedades" (ES), "Federal Income Tax" (US), "IRAE" (UY), "IUE" (BO), etc. Formato de moneda. |
| `_build_income_expense_book` | Branching: si país=MX → estructura original RESICO (UUID CFDI, RFC Tercero, hoja RESICO Calc, tasa 2%). Otros países → estructura general (Folio/Factura, ID fiscal genérico según `id_persona` del país, tasa de retención de honorarios local, sin hoja RESICO Calc). |

#### Builders NO refactorizados (28)
Los demás builders (`inventory`, `accounts_receivable`, `cashflow`, `dashboard`, `budget`, `break_even`, `amortization`, `balance_sheet`, `client_directory`, etc.) reciben `cfg` vía el dispatcher pero no lo usan en su firma. Esto es compatibilidad hacia atrás — generan contenido idéntico a v15.4. Su refactor es trabajo de v15.6+.

### Nuevo: `tests/test_excel_country.py` (114 tests)

Tests parametrizados:
- **63 smoke tests** (9 países × 7 templates críticos): cada combinación genera un archivo válido sin error.
- **9 tests de tasa IVA correcta** por país.
- **7 tests de label de impuesto correcto** (IVA / IGV / ITBMS / ITBIS / ISV / Sales Tax).
- **11 tests de deducciones de nómina específicas** por país (IMSS en MX, EPS en CO, AFP en CL, etc.).
- **1 test de Tabla LISR solo aparece para MX.**
- **1 test de income_expense_book MX vs genérico** (verifica RESICO Calc y headers).
- **4 tests de income_statement con label fiscal local** (ISR / Ganancias / Sociedades / Federal Income Tax).
- **12 tests de builders no-críticos siguen funcionando con cfg.**
- **4 tests de detección de país desde keywords libres** ("CDMX" → MX, "Bogotá" → CO, etc.).
- **1 test de consistencia entre Word y Excel** (mismo `LEGAL_CONFIG`, mismo set de países).
- **1 test de fallback a GENERIC** si el país no se reconoce.

Resultado: **114 passed, 0 failed.**

---

## Verificación de localización (output real)

```
quotation:
  México         → IVA 16% · moneda MXN
  Colombia       → IVA 19% · moneda COP
  Perú           → IGV 18% · moneda PEN
  Argentina      → IVA 21% · moneda ARS
  Panamá         → ITBMS 7% · moneda PAB/USD
  United States  → Sales Tax 0% · moneda USD

sales (header columna G):
  México         → "IVA (16%)"
  Colombia       → "IVA (19%)"
  Perú           → "IGV (18%)"
  Argentina      → "IVA (21%)"
  Panamá         → "ITBMS (7%)"

payroll (deducciones):
  México:    IMSS · ISR · INFONAVIT  (+ hoja Tabla ISR LISR Art. 96)
  Colombia:  Salud EPS · Pensión AFP · Retefuente · Solidaridad
  Perú:      AFP / ONP · IR 5ta cat. · CTS
  Argentina: Jubilación SIPA · Obra Social · PAMI · Ganancias
  Chile:     AFP · Salud Isapre/Fonasa · AFC · Impuesto único
  España:    Seguridad Social · IRPF
  Panamá:    CSS aporte personal · Seguro Educativo · ISR

freelance_tracker (retención):
  México:    "ISR retención honorarios (Art. 106 LISR)" 10%
  Colombia:  "Retención en la fuente — honorarios (10-11%)" 11%
  Perú:      "Retención IR 4ta categoría (8% honorarios)" 8%
  Chile:     "Retención boleta de honorarios (13.75% 2024+)" 13.75%
  España:    "Retención IRPF (15% / 7% primeros 2 años)" 15%

income_statement (label de impuesto):
  México:    "UTILIDAD ANTES DE ISR (EBT)"
  Argentina: "UTILIDAD ANTES DE IMPUESTO A LAS GANANCIAS (EBT)"
  Chile:     "UTILIDAD ANTES DE IMPUESTO A LA RENTA 1RA CAT. (EBT)"
  España:    "UTILIDAD ANTES DE IMPUESTO SOCIEDADES (EBT)"
  US:        "UTILIDAD ANTES DE FEDERAL INCOME TAX (EBT)"

income_expense_book:
  México:   hojas = [Registro, Fiscal Mes, RESICO Calc]
            headers = [#, Fecha, UUID CFDI, RFC Tercero, Descripción, ...]
  Colombia: hojas = [Registro, Fiscal Mes]
            headers = [#, Fecha, Folio/Factura, CC (Cédula de Ciudadanía) Tercero, ...]
```

---

## Decisiones de arquitectura

1. **`LEGAL_CONFIG` compartido en módulo separado.** Antes vivía en `word_controller.py`. Lo extraje a `office/legal_config.py` para que Excel pueda importarlo sin importar todo Word (y todo python-docx). Word ahora lo re-importa; comportamiento idéntico.

2. **`EXCEL_FISCAL` como overlay numérico, no como reemplazo.** Los campos textuales que necesita Word (leyes, artículos, instituciones con texto largo) están en `LEGAL_CONFIG`. Los campos numéricos que necesita Excel (tasas decimales, formatos de openpyxl, listas estructuradas de deducciones) están en `EXCEL_FISCAL`. Ambos diccionarios usan las mismas claves de país.

3. **Las tasas son DEFAULTS, no verdades fiscales.** Cada país tiene su payroll_deductions con tasas por defecto, pero estas se publican en una hoja `Parámetros` con celdas editables. Las escalas progresivas (que la mayoría de impuestos a la renta latinoamericanos son) llegan a la hoja con tasa 0% y una nota que dice "depende del salario, consulta la tabla oficial de [autoridad]". Esto es honesto — el sistema no inventa cálculos fiscales precisos para 21 jurisdicciones.

4. **Detección de país en `_dispatch`, no en cada builder.** El builder recibe `cfg` ya resuelto. No tiene que conocer la lógica de detección. Esto facilita testing y futuras extensiones.

5. **Migración gradual de builders.** El dispatcher usa `inspect.signature` para detectar si un builder acepta `cfg`. Los builders refactorizados (7) lo declaran y lo usan; los demás (28) no lo declaran y siguen funcionando. Esto permite hacer el refactor uno a uno, no en un big bang.

6. **`_build_income_expense_book` con branching MX vs general.** Era tan profundamente mexicano (UUID CFDI, RESICO Calc, RFC) que generalizar 1:1 hubiera perdido valor para usuarios mexicanos. Solución: si país=MX → estructura original; otros → estructura general más simple. Esto preserva la calidad del template mexicano sin sacrificarla por compatibilidad.

---

## Pendiente para v15.6+

- **28 builders restantes sin `cfg`.** Aunque no usan tasas fiscales (inventory, project_tracker, work_schedule, …), algunos podrían beneficiarse: formato de moneda local, día/mes vs mes/día, símbolo de moneda en headers. Trabajo mecánico.
- **Conectar `datos["items"]` al resto de builders.** Sigue siendo el otro hueco grande: solo 5 builders consumen datos del usuario, los demás generan sample. Refactor independiente del de v15.5.
- **Tablas de impuesto progresivas por país.** v15.5 las marca como "0% por defecto, consulta tabla oficial". Una v16 podría incorporar las tablas reales SUNAT/DIAN/AFIP como hojas auxiliares (como ya tenemos LISR para MX).
- **Decimal handling.** Algunos países (CLP, PYG, COP en algunos contextos) no usan decimales. Los formatos en `currency_format` ya reflejan esto, pero los valores numéricos generados como sample sí los tienen. Cosmético.

---

## Archivos modificados — checklist

- [x] `office/legal_config.py` — **nuevo** (938 líneas)
- [x] `office/excel_controller.py` — modificado (5,853 líneas, +260)
- [x] `office/word_controller.py` — modificado (5,654 líneas, –561)
- [x] `tests/test_excel_country.py` — **nuevo** (114 tests, todos en verde)
- [x] `CHANGELOG_v15.5.md` — **nuevo** (este archivo)
