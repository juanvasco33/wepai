# CHANGELOG — WEP AI v15.2

**Fecha:** 2026-05-26
**Base:** v15.1
**Objetivo:** Phase 2 de la arquitectura híbrida — 5 builders legales adicionales LLM-driven multi-país.

---

## Resumen

v15.1 estableció el patrón con `employment_contract`. v15.2 lo aplica a los 5 builders restantes donde la sustancia legal varía drásticamente por país. Total: **6 templates LLM-driven**, **32 templates hardcoded** (sin cambios).

| Phase | Template | Razón para LLM-driven |
|---|---|---|
| 1 | `employment_contract` | LFT mexicana hardcoded para todo país (LAW-01) |
| 2 | `termination_letter` | Causales legales y notice period varían radicalmente |
| 2 | `settlement` | Cálculo de liquidación es el MÁS country-variable de todos |
| 2 | `commercial_lease` | Marco civil/comercial específico, IVA, ajuste por inflación |
| 2 | `work_rules` | Obligación legal y contenido específico de cada código laboral |
| 2 | `power_of_attorney` | Tipos y requisitos notariales específicos por país |

---

## Cambios

### `office/llm_legal_generator.py` — 5 nuevos system prompts (1,180 líneas totales)

Cada uno mantiene la estructura del prompt de `employment_contract`:

- **Regla 1 — Anti-alucinación:** citar artículos solo con alta confianza, usar formulación general en caso contrario.
- **Regla 2 — Terminología correcta por país:** "empleador" vs "patrón", "arrendador/arrendatario", "otorgante/apoderado", según el país.
- **Regla 3 — Marco legal específico por país:** referencias al código laboral, civil, comercial correcto de cada jurisdicción.
- **Regla 4 — Estructura requerida:** número de cláusulas y orden estándar para el tipo de documento.
- **Regla 5 — Placeholders para datos faltantes:** nunca inventar nombres, montos, fechas.
- **Regla 6 — Idioma de salida estricto.**
- **Schema JSON unificado:** mismo formato que employment_contract, así el renderer común funciona para todos.

**Cobertura legal por template:**

`termination_letter`:
- México: Art. 47 LFT (rescisión patrón), Art. 51 LFT (rescisión trabajador)
- Colombia: Art. 62 CST (causales justas), Art. 64 CST (indemnización sin justa causa)
- Chile: Art. 159, 160 Código del Trabajo
- Argentina: Arts. 242, 245 LCT
- Perú: Arts. 22-25 LPCL
- España: Arts. 49-56 ET (despido objetivo, disciplinario, colectivo)
- Notice period (preaviso) específico por país

`settlement`:
- México: aguinaldo proporcional, vacaciones proporcionales, prima vacacional 25%, prima de antigüedad si aplica, indemnización por rescisión injustificada (3 meses + 20 días/año + salarios caídos)
- Colombia: cesantías acumuladas + intereses 12% anual + prima de servicios proporcional + vacaciones proporcionales + indemnización Art. 64 CST si despido sin justa causa
- Argentina: SAC proporcional, vacaciones proporcionales, indemnización Art. 245 LCT, sustitutiva del preaviso, integración mes de despido
- Chile: indemnización por años de servicio Art. 163, sustitutiva del aviso previo, feriado proporcional
- Perú: CTS proporcional, gratificaciones proporcionales, vacaciones truncas
- España: paga extra proporcional, vacaciones no disfrutadas, indemnización según tipo de despido

`commercial_lease`:
- México: CCF Arts. 2398+, IVA 16% arrendamiento comercial
- Colombia: Decreto-Ley 444/1996, Ley 820/2003, IVA 19%
- Chile: Ley 18.101, Código Civil
- Argentina: CCC Arts. 1187+, ajuste por ICL
- España: LAU Ley 29/1994 sección uso distinto vivienda, IVA 21%
- Cláusulas: identificación inmueble, destino comercial, plazo, renta, ajuste, garantía, gastos, terminación, jurisdicción

`work_rules`:
- México: Arts. 422-425 LFT, registro ante JCA/Tribunal Laboral, obligatorio para empresas
- Colombia: Capítulo II CST Arts. 104-125, aprobación Ministerio del Trabajo, obligatorio ≥5/10 trabajadores
- Chile: Arts. 153-157 Código del Trabajo, depósito en Inspección del Trabajo, obligatorio ≥10 trabajadores
- Argentina: implícito en LCT Arts. 67-69 (régimen disciplinario)
- Perú: obligatorio ≥100 trabajadores, aprobación Ministerio
- 10 capítulos estándar: objeto, horarios, descansos, pagos, permisos, obligaciones, higiene/seguridad, quejas, sanciones, vigencia

`power_of_attorney`:
- México: CCF Arts. 2546-2604, tipos (pleitos/administración/dominio), Art. 2554 (poder general)
- Colombia: CC Arts. 2142-2199, escritura pública para actos solemnes
- Argentina: CCC Arts. 362-381, Arts. 1319-1335 (mandato)
- Chile: CC Arts. 2116-2173 (mandato)
- Perú: CC Arts. 145-167, Arts. 1790+ (mandato)
- España: CC Arts. 1709-1739, poder notarial obligatorio para actos de disposición
- Tipos: general, especial, judicial, administrativo
- Warning explícito sobre certificación notarial en el closing

### `office/word_controller.py` — `LLM_DRIVEN_TEMPLATES` expandido

```python
LLM_DRIVEN_TEMPLATES = {
    "employment_contract",   # Phase 1
    "termination_letter",    # Phase 2
    "settlement",            # Phase 2
    "commercial_lease",      # Phase 2
    "work_rules",            # Phase 2
    "power_of_attorney",     # Phase 2
}
```

El resto del patch del dispatcher (try/except con fallback al hardcoded) se aplica automáticamente a los 6 templates.

### Modelo actualizado (no-deprecado)

`word_translator.py` y `llm_legal_generator.py` ahora usan `claude-sonnet-4-6` por defecto (antes `claude-sonnet-4-20250514` que sale de servicio el 15 de junio de 2026). Variables de entorno `WEPAI_MODEL_LEGAL` y `WEPAI_MODEL_TRANSLATOR` permiten override.

### Nuevo: `validate_llm_legal.py`

Script de validación con la API real, en la raíz del proyecto. Uso:

```bash
# Validación rápida — 1 doc, ~$0.15 USD
python3 validate_llm_legal.py

# Validación full — 30 docs (6 templates × 5 países), ~$5 USD
python3 validate_llm_legal.py --all

# Validación específica
python3 validate_llm_legal.py --template=settlement --country=Colombia
```

El script verifica automáticamente:
- API responde correctamente
- JSON parsea
- Doc se genera con tamaño razonable
- Contiene marcadores correctos del país (CST/LFT/LCT/ET según corresponda)
- NO contaminación cruzada (sin LFT en doc colombiano, sin EPS en mexicano, etc.)
- Latencia de generación

Y guarda los docs en `./test_outputs/` para revisión manual.

---

## Auditoría post-implementación

10 dimensiones verificadas:

1. **Syntax y carga** — 3 módulos parsean e importan, sin ciclos
2. **SYSTEM_PROMPTS** — 6 templates registrados, 6 user-message builders definidos
3. **_CACHE_RELEVANT_KEYS** — cobertura para los 6 templates
4. **LLM_DRIVEN_TEMPLATES en generate_word** — los 6 declarados
5. **Modelo no-deprecado** — claude-sonnet-4-6 en ambos módulos
6. **End-to-end con mocks** — los 6 templates generan docs con contenido del LLM
7. **Multi-país en settlement** — Colombia/México/Chile/Argentina/España cada uno con su contenido legal correcto, cero contaminación cruzada
8. **Fallback** — cuando LLM falla, los 6 templates usan el builder hardcoded
9. **Templates NO LLM-driven** — NDA, service_contract, CV, business_letter, executive_report siguen hardcoded sin invocar al LLM
10. **Smoke test** — 38/38 builders hardcoded siguen funcionando

**Resultado: 0 errores reales. El único warning era un falso positivo del script de test (caché en disco que no se limpiaba entre iteraciones — el código es correcto).**

---

## Cobertura del audit LAW-01

| Builder | Estado v14.9 | Estado v15.1 | Estado v15.2 |
|---|---|---|---|
| `employment_contract` | LFT hardcoded | ✓ LLM-driven multi-país | ✓ |
| `termination_letter` | Genérico | Genérico | ✓ LLM-driven multi-país |
| `settlement` | Mex-centric | Mex-centric | ✓ LLM-driven multi-país |
| `commercial_lease` | Genérico | Genérico | ✓ LLM-driven multi-país |
| `work_rules` | Mex-centric | Mex-centric | ✓ LLM-driven multi-país |
| `power_of_attorney` | Mex-centric | Mex-centric | ✓ LLM-driven multi-país |

**LAW-01 cerrado para los 6 builders legales country-specific.** Si el LLM falla, fallback al builder hardcoded de toda la vida (que sigue teniendo el bug LAW-01 pero al menos genera un doc).

---

## Cómo probarlo

**Validación rápida (1 doc, $0.15):**

```bash
cd wepai
export ANTHROPIC_API_KEY=sk-ant-...
python3 validate_llm_legal.py
```

**Validación full (6 templates × 5 países = 30 docs, ~$5):**

```bash
python3 validate_llm_legal.py --all
```

**Validación específica:**

```bash
python3 validate_llm_legal.py --template=settlement --country=México
python3 validate_llm_legal.py --template=power_of_attorney --country=Argentina
```

Los docs generados quedan en `./test_outputs/` para revisión manual. Abrirlos en Word y verificar visualmente que:
- El contenido legal es coherente para el país
- Las citas de artículos son verificables (buscarlas en el código respectivo)
- La terminología es la correcta de la jurisdicción
- Los placeholders `[Bracketed]` están donde corresponden y los datos del usuario están donde deben

---

## Costo proyectado por documento

| Doc tipo | Costo nuevo | Costo cacheado |
|---|---|---|
| `employment_contract` | $0.10–0.15 | $0.00 |
| `termination_letter` | $0.06–0.10 | $0.00 |
| `settlement` | $0.10–0.15 | $0.00 |
| `commercial_lease` | $0.10–0.15 | $0.00 |
| `work_rules` | $0.15–0.25 (más largo) | $0.00 |
| `power_of_attorney` | $0.08–0.12 | $0.00 |

Caché crece con cada documento generado. Una vez generados los 6 templates × 5 países × ~3 variantes c/u = ~90 entradas, el costo marginal por usuario tiende a cero excepto cuando piden parámetros nuevos.

---

## Lo que sigue (Phase 3 y siguientes)

**Phase 3 — Decisión basada en datos reales (cuando haya usuarios):**

Mirar telemetría real de uso:
- ¿Qué países piden los usuarios? Si ves volumen alto en Bolivia, Ecuador, Uruguay, validá esos países con un abogado local.
- ¿Qué tipos de documentos? Si ves que `commercial_lease` no se usa pero `residential_lease` sí, agregalo a `LLM_DRIVEN_TEMPLATES`.
- ¿Qué builders hardcoded muestran problemas? Si los reportes ejecutivos para España salen raros, considerá migrarlos también.

**Phase 4 — Validación humana (cuando tengas usuarios pagos):**

- Top 3 países por volumen → contratar revisión por abogado local de las plantillas LLM-generated
- Plantillas validadas pasan a caché como "lawyer-validated"
- Sello de calidad real para esos países: "Documentos revisados por abogados locales"

**Pricing tier insight (cuando arranques cobro):**

Plan Free: solo builders hardcoded (32 templates). Plan Pro: los 32 hardcoded + 6 LLM-driven country-specific. Esto:
- Mantiene Free sostenible (sin costo de API)
- Comunica el valor de upgrade ("documentos adaptados a tu país")
- Margen Pro: ~70% incluso con uso intensivo

---

*v15.2 cierra la categoría de documentos legales country-specific. La infraestructura compartida (caché, dispatch, renderer, fallback, validation script) está construida. Las próximas iteraciones son aplicaciones del mismo patrón cuando los datos de uso real lo justifiquen.*
