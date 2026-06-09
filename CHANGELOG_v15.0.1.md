# CHANGELOG — WEP AI v15.0

**Fecha:** 2026-05-26
**Base:** v14.9 Día 1
**Objetivo:** Cobertura bilingüe ES/EN real en Word sin reescribir los 38 builders.

---

## Resumen ejecutivo

La auditoría del módulo Word reveló que la cobertura bilingüe declarada en v14.4 era superficial: el test 11.1 solo validaba el título del documento, mientras que las cláusulas del cuerpo seguían en su mayoría hardcoded en español (~963 de 2,990 strings literales, 32%). Un usuario configurado en EN recibía un documento con título en inglés y cuerpo legal en español — peor que un documento monolingüe.

Esta versión introduce una **capa de traducción post-generación con LLM** que resuelve el problema sin tocar los 38 builders existentes.

| | v14.9 | v15.0 |
|---|---|---|
| Cobertura bilingüe real en Word | Solo títulos + algunos encabezados | Documento completo |
| Builders modificados | n/a | 0 (cero) |
| Costo por documento EN nuevo | n/a | ~$0.05–0.10 USD |
| Costo por documento EN repetido | n/a | ~$0.00–0.02 (caché hit) |
| Builders Word totales | 38 | 38 (sin cambios) |

---

## Cambios

### Nuevo archivo: `office/word_translator.py` (469 líneas)

Módulo standalone que expone `translate_doc_to_en(doc, doc_type, cfg)`. Recibe un objeto `docx.Document` ya generado por un builder (en español) y lo traduce in place a inglés usando Claude Sonnet.

**Pipeline:**

1. Carga el caché desde `~/.wepai/word_translations_cache.json` (lazy, una vez por proceso).
2. Recorre todos los párrafos del documento: body, celdas de tabla, headers, footers.
3. Heurística `_should_translate` decide qué fragmentos van al LLM:
   - Salta números puros, montos ("$50,000 MXN"), IDs ("CC 12345678"), texto claramente inglés ("The agreement between the parties").
   - Manda al LLM cualquier texto con marcadores ortográficos del español, palabras claras del español, o ambiguos (el LLM pasa el inglés sin tocar y el caché lo absorbe).
4. Los fragmentos no cacheados se envían a la API en batches de hasta 60 párrafos por llamada.
5. Las traducciones se aplican in place preservando el formato del primer run del párrafo (bold, tamaño, color, fuente).
6. Las nuevas entradas se persisten al caché.

**System prompt diseñado para español legal → inglés legal:**

- Glosario de términos legales con equivalentes correctos: `rescisión → termination` (no `rescission`), `patrón → employer`, `caso fortuito → force majeure`, `fuero → jurisdictional defense`, `comparecen → the parties hereby enter into`.
- Ordinales: `PRIMERA → FIRST`, `SEGUNDA → SECOND`, etc.
- Preservación estricta de: números, montos, nombres propios, IDs fiscales, referencias a artículos de ley (`Art. 50 LFT` queda verbatim, opcionalmente anotado).
- Placeholders entre corchetes: `[Nombre] → [Name]`, `[Dirección del Prestador] → [Provider's Address]`.
- Fechas: `21 de mayo de 2026 → May 21, 2026`.

**Failure mode:**

Si la API de Anthropic falla (sin red, sin API key, timeout, JSON malformado), la función:

1. No lanza excepción.
2. Devuelve el documento en español original.
3. Inserta un banner rojo arriba del documento: `⚠ ENGLISH TRANSLATION UNAVAILABLE — Document shown in original Spanish.`

Así, `generate_word()` nunca rompe el flujo de la UI por un fallo de traducción.

**API pública del módulo:**

```python
translate_doc_to_en(doc, doc_type="general", cfg=None) -> Document
cache_stats() -> dict                       # útil para monitoreo
clear_cache()                               # útil para tests
```

### Modificado: `office/word_controller.py`

Patch de 12 líneas al final de `generate_word()`, después de `_wep_footer(doc, cfg)` y antes de `doc.save(filepath)`:

```python
if lang == "en":
    try:
        from office.word_translator import translate_doc_to_en
        doc = translate_doc_to_en(doc, doc_type=template_id, cfg=cfg)
    except Exception as e:
        import logging
        logging.getLogger("wepai.word").error(
            "Word translation failed (returning Spanish doc): %s", e
        )
```

Cero cambios en los 38 builders. Cero cambios en `_disclaimer`, `_wep_footer`, `_user_data_block`, ni ninguna otra función.

---

## Configuración

**Variables de entorno (opcionales):**

- `WEPAI_MODEL_TRANSLATOR`: modelo a usar para traducción. Default `claude-sonnet-4-20250514`. Para cortar costos a ~1/5, exportar `claude-haiku-4-5-20251001` (la calidad legal es ligeramente menor pero para texto repetitivo es indistinguible).
- `ANTHROPIC_API_KEY`: la misma key que ya usa `agent/brain.py`. No requiere variable nueva.

**Caché:**

- Ubicación: `~/.wepai/word_translations_cache.json`
- Schema: `{ "texto_es": "texto_en", ... }`
- Crece con cada traducción exitosa; persistido tras cada batch.
- Para limpiarlo: `from office.word_translator import clear_cache; clear_cache()`.

---

## Cómo probarlo

Desde la raíz del proyecto, con `ANTHROPIC_API_KEY` exportada:

```bash
python3 -c "
from office.word_controller import generate_word
fp = generate_word({
    'template_id': 'service_contract',
    'titulo': 'Test_EN_Service_Contract',
    'instrucciones': 'Professional software development services agreement.',
    'datos': {
        'idioma': 'en',
        'parte_a_nombre': 'Acme Software Inc.',
        'parte_a_id': 'EIN 12-3456789',
        'parte_b_nombre': 'Client Corp',
        'parte_b_id': 'EIN 98-7654321',
        'objeto': 'Custom software development services',
        'monto': '50,000',
        'duracion': '6 months',
    }
})
print(f'Generado en: {fp}')
"
```

Abrir el .docx en Word y verificar:

- Título: `PROFESSIONAL SERVICES AGREEMENT` (no `CONTRATO DE PRESTACIÓN...`)
- Cláusulas: `FIRST. — SCOPE`, `SECOND. — TERM`, etc.
- Disclaimer: `⚠ LEGAL DISCLAIMER: This document...`
- Datos del usuario preservados: `Acme Software Inc.`, `EIN 12-3456789`, etc.

---

## Pendientes que NO se cierran en v15.0

Los siguientes hallazgos del audit de Word siguen pendientes (no son bloqueantes para EN básico, pero son los próximos a atacar):

- **BUG-01** — `_add_page_numbers` hardcodea `" de "` en el footer. En docs EN sale `"Page X de Y"`. El traductor no toca campos OOXML de footer (son `fldChar`, no runs de texto plano), así que esto sigue requiriendo un fix dedicado en `_add_page_numbers`.
- **BUG-02** — Subtítulo del NDA con ternario invertido (líneas 1029–1031 de word_controller). Cosmético pero queda raro: el doc se traduce correctamente al EN pero hay un punto donde el ES y EN están literalmente intercambiados a nivel código.
- **DEP-01** — `matplotlib` y `numpy` siguen sin estar declarados en `requirements.txt`. Builders con charts (`executive_report`, `commercial_proposal`) siguen crasheando en instalación limpia.
- **LAW-01** — Referencias a la LFT mexicana hardcoded en `_build_employment_contract` para usuarios fuera de México. El traductor anota `Art. 50 LFT (Mexican Federal Labor Law)` pero no corrige la ley aplicable: un usuario colombiano sigue recibiendo un contrato con citas mexicanas (ahora en inglés, lo cual paradójicamente lo hace más evidente).

Estos cuatro items deberían atacarse antes de habilitar Word EN en producción, pero v15.0 es la base sobre la cual fixearlos.

---

## Métricas

**Cobertura del traductor:**

| Componente del .docx | Cubierto |
|---|---|
| Body paragraphs | ✓ |
| Headings | ✓ (son paragraphs con style) |
| Table cells | ✓ |
| Headers (section) | ✓ |
| Footers (section) | ✓ |
| Charts (texto interno) | ✗ — son PNG embebidos |
| Footer fields (PAGE, NUMPAGES) | ✗ — son OOXML XML, no texto plano |

**Validación local:**

- Sintaxis: módulo parsea sin errores.
- Heurística `_should_translate`: 20/20 casos correctos (spanish con/sin tildes, inglés puro, números, IDs, montos, ordinales, labels cortos, nombres propios).
- Reemplazo in-place: preserva bold, tamaño de fuente, color del primer run.
- Iteración: encuentra párrafos en body, tablas, headers y footers.
- Caché: segunda corrida del mismo documento hace **0** API calls.

**Costo proyectado (Sonnet @ $3/M in, $15/M out):**

- Doc nuevo de tipo nunca generado: ~$0.05–0.10 USD
- Doc nuevo del mismo tipo: ~$0.00–0.02 USD (mayoría desde caché)
- Después de generar los 38 templates al menos una vez, el caché debería tener ~500–1500 entradas

---

*v15.0 cierra el gap entre la promesa de bilingüe declarada en v14.4 y la realidad del output. La estrategia de "traducción post-generación con LLM + caché creciente" se inspira en cómo Excel resolvió i18n con `_translate_workbook`, pero adaptada a la naturaleza de prosa larga del Word.*

---

## v15.0.1 — Fixes de los 4 bugs pre-existentes (2026-05-26)

Cuatro bugs detectados en el audit previo de Word, cerrados en una sola iteración sin tocar los builders.

### DEP-01 — matplotlib y numpy declarados en requirements.txt

Antes: imports lazy en `_chart_bar` y `_chart_line` (Word) más equivalentes en PPT crasheaban con `ModuleNotFoundError` en instalación limpia. Reportes ejecutivos y propuestas comerciales eran inutilizables sin instalación manual extra.

Después:
```
matplotlib>=3.7.0
numpy>=1.24.0
```

### BUG-02 — Subtítulo del NDA con ternario invertido (línea 1029)

Antes: el `if lang == "en"` mostraba el texto en español y viceversa. Un NDA en inglés tenía título correcto "NON-DISCLOSURE AGREEMENT" y debajo "(Acuerdo de Confidencialidad — NDA)" en español. Cosmético pero embarazoso.

Después: ternario corregido. EN muestra "(Non-Disclosure Agreement — NDA)", ES muestra "(Acuerdo de Confidencialidad — NDA)".

### SAFE-01 — `_safe_filename` con fallback para títulos vacíos

Antes: `_safe_filename("")` devolvía cadena vacía, generando un archivo `.docx` (oculto en macOS al empezar con punto). Si el LLM emitía título vacío por algún motivo, el usuario "perdía" el archivo.

Después: si el resultado quedaría vacío, devuelve `Documento_YYYYMMDD_HHMMSS`. Tested contra `""`, `"..."`, `"!!!@@@"`, `"  "`.

### BUG-01 — Footer "Page X de Y" en documentos en inglés

Antes: `_add_page_numbers(doc)` hardcoded `' de '`. Documentos en inglés mostraban "Page 1 de 5". La función no recibía `lang`, así que no había forma de fixearlo desde el caller. El traductor LLM tampoco podía arreglarlo porque `' de '` (2 letras alfa) está debajo del umbral de la heurística `_should_translate`.

Después: cadena de 3 funciones propaga `lang`:
- `generate_word` → pasa `lang=lang` a `_doc_base`
- `_doc_base(titulo, company="", ref="", margins=..., lang="es")` → pasa `lang=lang` a `_add_page_numbers`
- `_add_page_numbers(doc, lang="es")` → usa `" of "` o `" de "` según el idioma

Verificado funcionalmente: doc EN ahora muestra "Page X of Y", doc ES sigue mostrando "Página X de Y".

---

## Estado del audit post v15.0.1

| Item | Estado v14.9 | Estado v15.0 | Estado v15.0.1 |
|---|---|---|---|
| I18N en cuerpo de docs EN | ❌ Spanish leak | ✓ Traductor LLM | ✓ Sin cambios |
| DEP-01 (matplotlib/numpy) | ❌ Crash | ❌ Crash | ✓ Declarado |
| BUG-01 (footer "de"/"of") | ❌ Hardcoded ES | ❌ Hardcoded ES | ✓ Bilingüe |
| BUG-02 (NDA subtitle) | ❌ Invertido | ❌ Invertido | ✓ Correcto |
| SAFE-01 (filename vacío) | ❌ `.docx` oculto | ❌ `.docx` oculto | ✓ Fallback timestamp |
| LAW-01 (LFT hardcoded) | ❌ Sigue | ❌ Sigue | ❌ Sigue (pendiente) |
| SAFE-02 (sin dedup filename) | ❌ Sigue | ❌ Sigue | ❌ Sigue (pendiente) |
| I18N-03 (sin dict centralizado) | ❌ | ✓ Mitigado por traductor | ✓ |
| ARCH-01 (monolítico 5,7k LOC) | ❌ | ❌ | ❌ (defer) |

Lo que sigue en backlog para post-demo: LAW-01 (citas LFT mexicanas en builders multi-país), SAFE-02 (deduplicación de filenames), split de controllers monolíticos.

**Suite de tests pasada:**
- 38/38 builders generan sin crashear (smoke test)
- Footer EN dice "of", footer ES dice "de"
- NDA ES tiene subtítulo en español; NDA EN tiene subtítulo en inglés
- `_safe_filename("")` → `Documento_YYYYMMDD_HHMMSS.docx`
- Charts generan PNG válido con matplotlib
- Path ES sigue intacto: 0 llamadas al traductor
- Path EN: traductor invocado correctamente
- Path de fallo: banner rojo insertado al inicio

