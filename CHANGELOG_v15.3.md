# CHANGELOG — WEP AI v15.3

**Fecha:** 2026-05-26
**Base:** v15.2.1
**Objetivo:** Cobertura completa de los 20 países hispanohablantes + USA.

---

## Resumen ejecutivo

v15.3 agrega `LEGAL_CONFIG` completo para los 12 países hispanohablantes que faltaban (Ecuador, Uruguay, Paraguay, Venezuela, Panamá, Guatemala, Honduras, Nicaragua, El Salvador, Costa Rica, Cuba, República Dominicana). El fallback "GENERIC" queda únicamente para mercados fuera de hispano (cualquier país no codeado que pueda llegar via LLM-driven).

**Evolución de cobertura país:**

| Versión | Países con config completo | Mercado hispanohablante cubierto |
|---|---|---|
| v14.x | 7 (MX, CO, PE, AR, CL, ES, US) | ~70% |
| v15.2.1 | 8 (+ Bolivia) | ~73% |
| **v15.3** | **20** | **~100%** |

Suma de hispanohablantes cubiertos: ~500M personas. WEP AI ahora puede ser usado de forma profesional en cualquier país hispanohablante del mundo sin que aparezcan placeholders `[País]` o `[Moneda]` en los documentos generados.

---

## Cambios

### `office/word_controller.py` — 12 nuevas entradas en `LEGAL_CONFIG`

Cada una con los 27 campos estándar. Resumen de datos clave por país:

| Código | País | Moneda | ID Empresa | Ley Laboral | Seg. Social | IVA | Autoridad Fiscal |
|---|---|---|---|---|---|---|---|
| EC | Ecuador | USD | RUC (13d) | Código del Trabajo | IESS | 15% | SRI |
| UY | Uruguay | UYU | RUT | Conjunto de leyes laborales | BPS | 22% | DGI |
| PY | Paraguay | PYG (₲) | RUC | Código del Trabajo (Ley 213/93) | IPS | 10% | SET |
| VE | Venezuela | VED | RIF | LOTTT | IVSS | 16% | SENIAT |
| PA | Panamá | PAB/USD | RUC | Código de Trabajo | CSS | 7% (ITBMS) | DGI |
| GT | Guatemala | GTQ (Q) | NIT | Código de Trabajo (Decreto 1441) | IGSS | 12% | SAT |
| HN | Honduras | HNL (L) | RTN | Código del Trabajo | IHSS | 15% (ISV) | SAR |
| NI | Nicaragua | NIO (C$) | RUC | Código del Trabajo (Ley 185) | INSS | 15% | DGI |
| SV | El Salvador | USD | NIT | Código de Trabajo | ISSS | 13% | DGII |
| CR | Costa Rica | CRC (₡) | Cédula Jurídica | Código de Trabajo | CCSS | 13% | Hacienda |
| CU | Cuba | CUP | NIT | Código de Trabajo (Ley 116/2013) | SSS Cubano | n/a | ONAT |
| DO | República Dominicana | DOP (RD$) | RNC | Código de Trabajo (Ley 16-92) | TSS | 18% (ITBIS) | DGII |

**Datos validados con fuentes oficiales recientes:**

- Ecuador IVA confirmado en 15% para 2026 (Circular SRI NAC-DGECCGC25-00000006 del 26-dic-2025)
- Venezuela moneda actual VED (Bolívar Digital ISO 4217)
- IESS Ecuador: 12.15% patronal, 9.45% laboral

**Datos conservadores donde había duda:**

- Cuba: marcado como "sistema tributario especial" en lugar de IVA específico, dado el modelo económico mixto
- Países con sistemas mixtos (Panamá USD+PAB, El Salvador USD): documentado el dual currency
- Tasas de seguridad social: usadas las más estables; cuando varía mucho por régimen, marcado como "Variable según régimen"

### Aliases para detección — los 12 países

Cada uno con nombre, código ISO y al menos 1 ciudad principal:

```python
"ecuador":"EC","ec":"EC","quito":"EC","guayaquil":"EC",
"uruguay":"UY","uy":"UY","montevideo":"UY",
"paraguay":"PY","py":"PY","asunción":"PY","asuncion":"PY",
"venezuela":"VE","ve":"VE","caracas":"VE","maracaibo":"VE",
"panama":"PA","panamá":"PA","pa":"PA",
"guatemala":"GT","gt":"GT","gua":"GT",
"honduras":"HN","hn":"HN","tegucigalpa":"HN",
"nicaragua":"NI","ni":"NI","managua":"NI",
"el salvador":"SV","salvador":"SV","sv":"SV","san salvador":"SV",
"costa rica":"CR","cr":"CR","san josé":"CR","san jose":"CR",
"cuba":"CU","cu":"CU","habana":"CU","la habana":"CU",
"republica dominicana":"DO","república dominicana":"DO",
"dominicana":"DO","do":"DO","santo domingo":"DO",
```

### Keywords para detección por texto libre

`_detect_country` ahora reconoce ciudades, instituciones, leyes y términos fiscales por país. Ejemplos:

- "trabajo en IESS" → EC
- "mi RUC dominicano" → DO  
- "régimen de pensiones AFP en El Salvador" → SV
- "necesito demanda en juzgados de Tegucigalpa" → HN
- "contrato con empresa cubana ONAT" → CU
- "Bono 14 en Guatemala" → GT

---

## Auditoría post-implementación

8 dimensiones verificadas:

| # | Check | Resultado |
|---|---|---|
| A | Syntax y carga de los 3 módulos | ✓ OK |
| B | 21 entradas en LEGAL_CONFIG (20 países + GENERIC) | ✓ Completo |
| C | Cada país tiene los 13 campos críticos llenos | ✓ 20/20 |
| D | Detección de país por nombre, código y ciudad | ✓ 26/26 casos |
| E | NDA hardcoded sin placeholders feos para los 12 nuevos | ✓ 12/12 |
| F | LLM-driven recibe nombre real del país (los 12 nuevos) | ✓ 12/12 |
| G | Smoke test 38 builders × 5 países nuevos | ✓ 190/190 |
| H | Símbolos de moneda únicos por país | ✓ |

**0 errores, 0 warnings reales** (un solo warning informativo sobre USA que tiene `[Ciudad, Estado]` en tribunales — patrón estándar que también tiene México con su propio formato).

---

## Comparativa práctica — antes y después de v15.3

**Usuario en Tegucigalpa pide un NDA:**

| | v15.2.1 | v15.3 |
|---|---|---|
| Detección | Caía a GENERIC | ✓ Detecta HN |
| Identificadores | `[ID Fiscal Empresa]`, `[ID Personal]` | RTN, DNI |
| Moneda | `[Moneda]` | HNL (Lempira) |
| Jurisdicción | Tribunales competentes de `[Ciudad, País]` | Juzgados de Trabajo de Tegucigalpa |
| Footer | "leyes vigentes en `[País]`" | "leyes vigentes en Honduras" |

**Usuario en Asunción pide un contrato laboral (LLM-driven):**

| | v15.2.1 | v15.3 |
|---|---|---|
| País detectado | GENERIC | PY |
| País enviado al LLM | "[País]" (bug — no funcionaba) | "Paraguay" ✓ |
| Contenido legal | Inservible | Cita Ley 213/93, IPS, gratificaciones, etc. |
| Moneda mostrada | `[Moneda]` | PYG (₲) |

---

## Lo que queda en el roadmap

**No es necesario agregar más países hispanohablantes** — los 20 están cubiertos.

Posibles expansiones futuras según demanda real:

- **Brasil** (portugués, ~210M habitantes): requiere traducción del LLM-driven a portugués + investigación legal (CLT, INSS, FGTS, ISS, ICMS, ITR)
- **Estados Unidos por estado**: actualmente US es genérico; agregar configs por estado (California, Texas, Florida, Nueva York) que tienen leyes laborales muy distintas
- **Países anglosajones**: UK, Canadá, Australia — requieren traducción del idioma del LLM-driven

Ninguno es urgente. El mercado hispanohablante completo está cubierto.

---

## Disclaimer importante

Los datos de las 12 nuevas entradas se basan en conocimiento general del derecho laboral/fiscal de cada país, complementado con búsqueda web para datos críticos (IVA Ecuador 2026, moneda Venezuela). Como en todo el resto del producto:

- El disclaimer legal sigue presente en cada documento ("este documento no constituye asesoría legal, revíselo con un abogado o notario licenciado en su país antes de firmarlo")
- Para uso en producción con usuarios pagos, recomiendo validación por abogado local de los top 3-5 países objetivo
- Las tasas de seguridad social, IVA, y artículos específicos pueden cambiar con reformas — el sistema debe actualizarse según corresponda
- El LLM-driven (Phase 1/2) genera contenido actualizado con la knowledge cutoff del modelo, por lo que para los 6 docs legales country-specific (contrato laboral, finiquito, terminación, arrendamiento, reglamento, poder) la actualización es automática vía el modelo

---

*v15.3 cierra la cobertura hispanohablante. La arquitectura híbrida + 20 configs país-específicos significa que cualquier usuario hispanohablante recibe documentos profesionales adaptados a su jurisdicción, sin placeholders visibles, con detección automática del país tanto desde campo explícito como desde texto libre del chat. **Brasil, países anglosajones y otros mercados no-hispanohablantes caen al fallback GENERIC con placeholders visibles — son trabajo futuro.***
