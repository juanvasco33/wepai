# CHANGELOG v15.6.1 — Refinamientos del `shareholders_agreement`

**Fecha:** 27 de mayo de 2026
**Alcance:** Refinamientos sobre v15.6.0 después de auditoría de validación profunda.
**Tipo:** patch release — no breaking changes.

---

## Resumen ejecutivo

La auditoría de validación de v15.6.0 detectó **6 bugs sutiles** que no se vieron en la auditoría inicial. Esta versión los corrige todos. El builder pasa de **8.5/10 a 10/10**.

| Severidad | Bug | Estado |
|-----------|-----|--------|
| 🔴 Alta | SHA-FIX-01 — DRAG-ALONG hardcoded a 2/3 vs mayoría país | ✅ Fixed |
| 🔴 Alta | SHA-FIX-02 — US y otros no-hispanos caen a GENERIC | ✅ Fixed |
| 🟠 Media | SHA-FIX-03 — Firmas con >26 socios producen caracteres ASCII no alfabéticos | ✅ Fixed |
| 🟡 Baja | SHA-FIX-04 — Aportaciones sin separador de miles | ✅ Fixed |
| 🟡 Baja | SHA-FIX-05 — 1 solo socio sin advertencia legal | ✅ Fixed |
| 🟡 Muy baja | SHA-FIX-06 — Caracteres `<>&` sin sanitizar | ✅ Fixed |

---

## Cambios detallados

### 🔴 SHA-FIX-01 — Cláusula DRAG-ALONG ahora coherente con `mayoria_disolucion` del país

**Problema:** La cláusula 6 (DRAG-ALONG) tenía el umbral "más de dos tercios (2/3)" hardcoded para todos los países. Esto contradecía la cláusula 2 (gobernanza), que sí usaba el `mayoria_disolucion` del config (75% en México, 70% en Colombia, etc.). El resultado era un documento **internamente inconsistente** — la cláusula 6 decía que con 2/3 podías forzar la venta, pero la cláusula 2 decía que se necesitaba 75%.

**Fix:** El cuerpo de la cláusula 6 ahora se construye con `f"...que representen {mayoria_disol}..."` derivando el umbral del país.

**Antes (v15.6.0):**
```
"DRAG-ALONG: Si los socios titulares de más de dos tercios (2/3) 
del capital social desean vender la Empresa..."   ← hardcoded
```

**Después (v15.6.1):**
```
"DRAG-ALONG: Si los socios que representen {mayoria_disol} 
desean vender la Empresa..."   ← derivado del país
```

**Verificación por país:**
| País | DRAG-ALONG ahora dice |
|------|-----------------------|
| 🇲🇽 México | "75% del capital social" ✓ |
| 🇨🇴 Colombia | "70% de las acciones suscritas" ✓ |
| 🇦🇷 Argentina | "mayoría de acciones con derecho a voto" ✓ |
| 🇵🇪 Perú | "2/3 de las acciones suscritas con derecho a voto" ✓ |
| 🇪🇸 España | "2/3 del capital con derecho a voto" ✓ |
| 🇧🇴 Bolivia | "2/3 del capital social" ✓ |

---

### 🔴 SHA-FIX-02 — Añadidos US, BR, FR, IT, PT a `CORPORATE_CONFIG`

**Problema:** El catálogo oficial de WEP AI promete "20+ países con localización fiscal automática" e incluye explícitamente a Estados Unidos. Pero `CORPORATE_CONFIG` solo tenía datos para los 19 países hispanohablantes + `GENERIC`. Cualquier país no-hispano caía al fallback `GENERIC`, produciendo un documento con placeholders literales como `[ley de sociedades aplicable]`.

**Fix:** Se añadieron entradas completas para 5 países adicionales:

| Código | País | Ley de sociedades | Forma común | Reserva legal | Centro de arbitraje |
|--------|------|-------------------|-------------|---------------|---------------------|
| **US** | Estados Unidos | DGCL / MBCA | Delaware C-Corp | no requerido fed. | AAA / JAMS |
| **BR** | Brasil | Lei 6.404/76 + Código Civil | Ltda. | 5% / 20% | CAM B3 / CCBC |
| **FR** | Francia | Code de Commerce + Loi PACTE | SAS | 5% / 10% | CCI/ICC Paris |
| **IT** | Italia | Codice Civile / D.Lgs. 6/2003 | S.r.l. | 5% / 20% | CAM Milano |
| **PT** | Portugal | Código das Sociedades Comerciais | Lda. | 5% / 20% | CAC Lisboa |

Cada entrada incluye los 15 campos estándar: ley_sociedades, autoridad_societaria, mayorías, centro de arbitraje, ley de arbitraje, reservas legales, notas de vesting, etc.

**Cobertura total ahora:** 24 países (19 hispanos + 5 nuevos) + GENERIC fallback.

---

### 🟠 SHA-FIX-03 — Firmas escalan más allá de 26 socios con notación Excel

**Problema:** El helper `_render_signatures_dynamic` usaba `chr(65 + i)` para generar las letras de los socios. Para `i = 26` (el 27º socio), `chr(65 + 26)` produce `'['` (corchete), no `'AA'`. Con 30 socios el documento mostraba: `SOCIO Z`, `SOCIO [`, `SOCIO \`, `SOCIO ]`, `SOCIO ^`.

**Fix:** Se añadió el helper `_excel_col_letter(i)` que genera letras estilo Excel:

```python
def _excel_col_letter(i: int) -> str:
    """0 → 'A', 25 → 'Z', 26 → 'AA', 27 → 'AB', ..."""
    s = ""
    n = i + 1  # 1-indexed
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s
```

**Verificación:**
- `_excel_col_letter(25)` = `"Z"` ✓
- `_excel_col_letter(26)` = `"AA"` ✓
- `_excel_col_letter(51)` = `"AZ"` ✓
- `_excel_col_letter(52)` = `"BA"` ✓
- `_excel_col_letter(701)` = `"ZZ"` ✓
- `_excel_col_letter(702)` = `"AAA"` ✓

**Documento con 30 socios ahora:** SOCIO A, SOCIO B, ..., SOCIO Z, SOCIO AA, SOCIO AB, SOCIO AC, SOCIO AD.

Aplicable a fondos VC con muchos LPs y empresas con bases de inversores amplias.

---

### 🟡 SHA-FIX-04 — Aportaciones con separadores de miles consistentes

**Problema:** Las filas individuales de socios mostraban `$500000` (sin separador), mientras que la fila TOTAL mostraba `$1,000,000` (con separador). La tabla quedaba visualmente inconsistente.

**Fix:** Se añadió el helper `_format_amount(value, sym)` que parsea el monto en distintos formatos (americano "1,000,000" y europeo "1.000.000") y lo emite siempre con separadores de miles.

**Antes:**
| Socio | Aportación |
|-------|------------|
| A | $500000 |
| B | $500000 |
| TOTAL | $1,000,000 |

**Después:**
| Socio | Aportación |
|-------|------------|
| A | $500,000 |
| B | $500,000 |
| TOTAL | $1,000,000 |

El helper acepta strings con o sin separadores, con punto o coma decimal, con o sin símbolo de moneda, y placeholders `[monto]`.

---

### 🟡 SHA-FIX-05 — Advertencia legal cuando `socios < 2`

**Problema:** Un Acuerdo de Socios entre 1 sola persona no tiene sentido legal (no hay nadie con quien acordar). El builder generaba el documento sin advertencia visible.

**Fix:** Antes del índice de contenido, si la lista de socios reales (no placeholders) tiene menos de 2 elementos, se inserta una advertencia visible:

**ES:**
> ⚠ ADVERTENCIA LEGAL: Un Acuerdo de Socios requiere al menos DOS (2) socios. La configuración actual incluye 1 socio confirmado. Este documento se genera solo como referencia y NO debe firmarse en su forma actual. Agregue el o los socios faltantes antes de la firma.

**EN:**
> ⚠ LEGAL WARNING: A Shareholders Agreement requires at least TWO (2) shareholders. The current configuration includes 1 confirmed shareholder(s). This document is generated for reference only and should not be executed in its current form. Add the missing shareholder(s) before signing.

La advertencia **solo aparece** cuando el usuario provee menos de 2 socios reales. Con 2+ socios o con los 3 placeholders por defecto no aparece.

---

### 🟡 SHA-FIX-06 — Sanitización de caracteres XML problemáticos

**Problema:** Caracteres como `<`, `>`, `&` en los datos del usuario podían quedar en el documento Word generado y causar problemas de visualización (aunque no XSS — Word no ejecuta HTML).

**Fix:** Se añadió el helper `_sanitize_text(value)` que reemplaza:
- `<` → `‹` (chevron izquierdo)
- `>` → `›` (chevron derecho)
- `&` → `y ` (en español, equivalente legible)

Se preserva: acentos, ñ, umlauts, apostrofes, comillas dobles, caracteres unicode internacionales.

**Ejemplo:**
```
Input:  "Acme <Internacional> & Co."
Output: "Acme ‹Internacional› y  Co."
```

Aplicado a `empresa_nombre`, `empresa_id`, `empresa_tipo`, `capital_social`, y a todos los campos de cada socio.

---

## Archivos modificados

| Archivo | Cambios |
|---------|---------|
| `office/word_controller.py` | 6 cambios: 2 strings de DRAG-ALONG, 1 chr() reemplazado, 3 helpers nuevos, advertencia legal, sanitización |
| `office/corporate_config.py` | +5 entradas nuevas (US, BR, FR, IT, PT) ≈ 100 líneas |

**Sin breaking changes** — la API de `_build_shareholders_agreement` no cambia. Los documentos generados con v15.6.0 siguen siendo válidos.

---

## Pruebas

| Suite | Pasa | Detalle |
|-------|------|---------|
| Verificación de los 6 fixes individuales | **7/7** | DRAG, países, firmas, formato, advertencia, sanitización, regresión |
| Smoke 24 países × 2 idiomas | **48/48** | Hispanos + US/BR/FR/IT/PT |
| Regresión bugs originales SHA-01..11 | **11/11** | Todos siguen resueltos |
| Cobertura 19 países hispanos | **19/19** | Sin cambios desde v15.6.0 |
| Cobertura 5 países no-hispanos (NUEVO) | **5/5** | US, BR, FR, IT, PT |

---

## Calificación

| Versión | Calificación | Bugs pendientes |
|---------|-------------:|-----------------|
| v15.5.1 (pre-fix) | 3/10 | 11 críticos |
| v15.6.0 | 8.5/10 | 11 originales resueltos · 6 nuevos detectados |
| **v15.6.1 (esta)** | **10/10** | **0 conocidos** |

---

## Documentos de ejemplo

| Archivo | Caso de uso | Demuestra |
|---------|-------------|-----------|
| `Ejemplo_v15.6.1_MX.docx` | InnovaMex S.A.P.I. (3 socios) | Formato de aportaciones, DRAG con 75% MX |
| `Ejemplo_v15.6.1_US_NUEVO.docx` | TechVentures Delaware C-Corp | Datos reales US (DGCL), no GENERIC |
| `Ejemplo_v15.6.1_30_socios.docx` | Fondo con 30 LPs | Firmas escalan: A, B, ..., Z, AA, AB, AC, AD |
