# CHANGELOG — WEP AI v15.1

**Fecha:** 2026-05-26
**Base:** v15.0.1
**Objetivo:** Arquitectura híbrida — Phase 1: contrato laboral LLM-driven multi-país.

---

## Contexto y motivación

El audit anterior dejó LAW-01 abierto: `_build_employment_contract` hardcodea referencias a la LFT mexicana (`Art. 47 LFT`, `Art. 50 LFT`, `EL PATRÓN`) que se inyectan incluso cuando el usuario está en Colombia, Argentina, Chile o cualquier otro país. Resultado: contratos laborales legalmente incorrectos para 7 de los 8 países cubiertos por `LEGAL_CONFIG`.

La arquitectura híbrida de v15.1 resuelve este problema de raíz para los documentos donde la sustancia legal es country-specific. La idea: los builders se dividen en dos grupos.

**Hardcoded (~25 builders, sin cambios):** documentos country-neutral donde la estructura importa más que la jurisdicción — NDAs, cartas comerciales, CVs, memos, propuestas, planes de negocio, reportes ejecutivos, T&Cs, manuales, facturas. Estos siguen siendo gratis, instantáneos y deterministas como hasta ahora.

**LLM-driven (Phase 1: 1 builder, próximas fases sumarán ~10–13):** documentos donde el contenido legal varía drásticamente por país — contratos laborales, cartas de despido, finiquitos, arrendamientos residenciales, reglamentos internos, poderes notariales. Estos se generan vía Claude con system prompt afinado para producir contenido legal específico al país del usuario.

Phase 1 implementa el primer builder con la nueva arquitectura: `employment_contract`. El patrón validado aquí se replica en las próximas fases.

---

## Cambios

### Nuevo archivo: `office/llm_legal_generator.py` (635 líneas)

Módulo standalone que genera documentos legales con LLM. Arquitectura paralela al `word_translator.py` de v15.0 pero con un propósito distinto: en lugar de traducir contenido ya generado, genera el contenido desde cero adaptado al país.

**API pública:**

```python
build_via_llm(doc, doc_type, instruccion, secciones, detalles, cfg, datos)
    # Builder con firma compatible con los builders hardcoded.
    # Genera estructura vía LLM y la renderiza usando helpers de word_controller.

generate_legal_structure(doc_type, country, lang, datos) -> dict
    # Llama al LLM con system prompt específico del doc_type.
    # Devuelve JSON estructurado con: titulo, subtitulo, intro, cláusulas[], closing, partes[].

render_legal_doc(doc, structure, datos, cfg)
    # Toma el JSON del LLM y produce el .docx con formato visual consistente
    # (mismos colores, fuente, headings, signature_block que el resto del app).

cache_stats() / clear_cache() / is_llm_driven(template_id)
    # Utilidades para monitoreo y dispatch decisions.
```

**System prompt diseñado para precisión legal:**

El prompt de `employment_contract` cubre 8 reglas críticas:

1. **Anti-alucinación de citas legales** — solo citar artículos cuando hay alta confianza, usar formulación general cuando hay duda ("conforme a la legislación laboral vigente del país") en lugar de fabricar números de artículo.

2. **Terminología correcta por país** — `empleador` en Colombia/Chile/Argentina/Perú/España (nunca `patrón`), `patrón` o `empleador` en México, `trabajador` formal en la mayoría de LATAM, `Employee/Employer` en USA.

3. **Sistemas de identificación correctos** — NIT/CC en Colombia, RFC/CURP en México, CUIT/CUIL en Argentina, RUT en Chile, RUC/DNI en Perú, CIF/NIF en España, EIN/SSN en USA.

4. **Instituciones correctas por país** — IMSS/AFORE en México, EPS/AFP/Caja en Colombia, AFP/Fonasa en Chile, ANSES en Argentina, EsSalud en Perú, SS en España. Nunca cruzar.

5. **Prestaciones obligatorias por país** — cesantías + prima de servicios + auxilio de transporte para Colombia, aguinaldo + PTU para México, SAC + obra social para Argentina, gratificación legal + AFP para Chile, gratificaciones + CTS para Perú, pagas extras para España.

6. **Estructura estándar de 8–10 cláusulas** con orden esperado: tipo de relación, objeto, jornada, remuneración, prestaciones, lugar de trabajo, obligaciones, confidencialidad, terminación, jurisdicción.

7. **Inyección de datos del usuario** — usar los datos exactos provistos, placeholders `[Bracketed]` para los faltantes, nunca inventar.

8. **Idioma de salida** — generar entero en `lang` solicitado, sin mezclar.

**Caché en disco:**

- Ruta: `~/.wepai/llm_legal_cache.json`
- Key: hash SHA256 de `(doc_type, country, lang, params_relevantes)`
- Params que entran en la key: `tipo_contrato`, `puesto`, `salario`, `fecha_inicio`, `departamento`, `tipo_jornada`, `dias_laborales`, `patron_nombre`, `trabajador_nombre`.
- Una segunda corrida con los mismos parámetros hace 0 API calls.

**Fallback robusto:**

Si el LLM falla (API caída, JSON inválido, doc_type no soportado), `build_via_llm` lanza la excepción. El caller (`generate_word`) la captura y ejecuta el builder hardcoded de toda la vida — el usuario recibe el documento aunque la calidad legal sea la "vieja" (con bug LAW-01) en ese caso particular.

### Patch en `office/word_controller.py`

Bloque de dispatch híbrido insertado al inicio de la cadena de dispatchers en `generate_word()`:

```python
LLM_DRIVEN_TEMPLATES = {"employment_contract"}  # Phase 1
template_id = (gen_data.get("template_id") or "").strip().lower()
_llm_handled = False
if template_id in LLM_DRIVEN_TEMPLATES:
    try:
        from office.llm_legal_generator import build_via_llm
        build_via_llm(doc, template_id, instruccion, secciones, detalles, cfg, datos)
        _llm_handled = True
    except Exception as e:
        import logging
        logging.getLogger("wepai.word").error(
            "LLM legal generation failed for %s, falling back to hardcoded: %s",
            template_id, e
        )

if _llm_handled:
    pass  # ya generado via LLM
elif template_id and template_id in WORD_TEMPLATES:
    WORD_TEMPLATES[template_id](doc, titulo, instruccion, secciones, detalles, cfg, datos)
# ... resto del dispatcher de keywords sin cambios
```

Cambios mínimos en código existente: solo se convierte el primer `if` del dispatcher en `elif` (un caracter). Cero cambios en los 38 builders.

---

## Cómo se conectan las tres capas

El flujo completo de un contrato laboral colombiano en v15.1 es:

1. **Usuario** pide en chat: "Necesito un contrato laboral indefinido para Colombia, posición Senior Engineer, salario 5 millones COP"
2. **`agent/brain.py`** detecta intención y emite `gen_data` con `template_id="employment_contract"` y datos del país, salario, etc.
3. **`generate_word()`** ve `employment_contract` en `LLM_DRIVEN_TEMPLATES` → invoca `llm_legal_generator.build_via_llm()`
4. **`build_via_llm`** consulta caché por `(employment_contract, Colombia, es, hash_params)` — caché miss
5. **`_call_llm`** envía system prompt + datos del usuario a Claude Sonnet
6. **Claude** devuelve JSON estructurado con 10 cláusulas adaptadas a Colombia (Art. 78 CST, Art. 249 CST, EPS, EMPLEADOR, etc.)
7. **`render_legal_doc`** convierte el JSON a .docx usando `_heading`, `_para`, `_divider`, `_signature_block` de `word_controller`
8. **`generate_word()`** continúa: añade `_user_data_block` (safety net), `_disclaimer`, `_wep_footer`
9. **Si `lang="en"`**, el `word_translator.py` de v15.0 procesa el doc (aunque en este caso no hay nada que traducir porque el LLM ya generó en EN si se le pidió)
10. **Doc guardado** y devuelta la ruta al usuario

Si en el paso 5 la API falla, el flujo cae al builder hardcoded de toda la vida — el usuario recibe el documento con el bug LAW-01 conocido pero al menos recibe algo.

---

## Costos y métricas

| Métrica | Valor |
|---|---|
| Costo por contrato laboral nuevo (Sonnet) | ~$0.10–0.20 USD |
| Costo por contrato laboral repetido (mismos params) | $0.00 (caché HIT) |
| Latencia primera generación | 15–25 segundos |
| Latencia generación cacheada | ~100 ms |
| Builders convertidos a LLM-driven | 1 de 38 |
| Builders hardcoded sin tocar | 37 (incluidos los 38 originales — uno pasa por la ruta hardcoded como fallback) |
| Líneas nuevas en código | 635 (`llm_legal_generator.py`) + 20 (patch en `word_controller.py`) |
| Líneas modificadas en código existente | 1 (`if` → `elif` en el dispatcher) |

---

## Auditoría post-implementación

El audit verificó 8 dimensiones:

1. **Syntax y carga** — 3 módulos parsean e importan, sin ciclos.
2. **API pública** — 6 funciones exportadas, `is_llm_driven` discrimina correctamente.
3. **Dispatch híbrido** — orden correcto (LLM antes que hardcoded), flag de éxito, manejo de error con fallback.
4. **End-to-end con LLM mockeado** — doc colombiano generado contiene `Art. 78 CST`, `Art. 249 CST`, `EPS`, `EMPLEADOR`; NO contiene `LFT`, `IMSS`, `EL PATRÓN`. Disclaimer presente, footer correcto.
5. **Caché funciona** — segunda corrida con mismos params hace 0 API calls.
6. **Fallback robusto** — cuando LLM falla, el builder hardcoded se ejecuta sin que `generate_word` rompa.
7. **Templates NO LLM-driven** — NDA y demás siguen hardcoded sin invocar al generador legal.
8. **Smoke test** — los 38 builders hardcoded siguen funcionando intactos.

**Resultado: 0 errores. La integración Phase 1 está limpia.**

---

## Lo que sigue (próximas fases)

**Phase 2** — Migrar 4–5 builders adicionales country-specific al patrón LLM-driven:

- `termination_letter` (carta de despido)
- `settlement` (finiquito / liquidación de prestaciones)
- `commercial_lease` (contrato de arrendamiento comercial)
- `work_rules` (reglamento interno de trabajo)
- `power_of_attorney` (poder notarial)

Cada uno requiere su propio system prompt en `SYSTEM_PROMPTS` y su builder de mensaje en `_build_user_message`. La infraestructura común (caché, fallback, renderer, dispatch) ya está hecha.

**Phase 3** — Decisión basada en datos reales de uso:
- ¿Qué países están solicitando los usuarios reales?
- ¿Qué tipos de documentos están pidiendo?
- Migrar al patrón LLM-driven solo los que el volumen justifique.

**Phase 4** — Validación humana progresiva:
- Para los top 3 países (probablemente MX, CO, AR), contratar revisión por abogado laboralista local de las plantillas LLM-generated.
- Las plantillas validadas pasan al caché como "lawyer-validated" y dejan de consumir API.
- WEP AI puede comunicar legítimamente "documentos revisados por abogados locales" para esos países.

**Pricing tier insight** — el plan Free podría restringirse a builders hardcoded (gratis, sin API cost). Los documentos country-specific generados por LLM quedan disponibles solo en Pro/Enterprise. Esto:
- Elimina pérdida de margen en Free
- Crea razón legítima para upgrade
- Comunica el valor: "Pro/Enterprise = documentos adaptados a la ley específica de tu país"

---

*v15.1 establece el patrón arquitectónico. Las próximas fases son aplicaciones del mismo patrón a más builders. El costo marginal de cada migración baja conforme avanza, porque la infraestructura compartida (caché, dispatch, renderer, fallback) ya está construida.*
