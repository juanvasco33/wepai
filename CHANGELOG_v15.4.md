# CHANGELOG — WEP AI v15.4

**Fecha:** 2026-05-26
**Base:** v15.3
**Tipo:** Hotfix de estabilidad + refactor de configuración + tests.

---

## Resumen ejecutivo

v15.4 aplica tres fixes críticos detectados en auditoría externa y agrega la primera batería seria de tests automatizados para los builders de Word. Ningún cambio funcional visible para el usuario — solo estabilidad.

| Bug | Estado v15.3 | Estado v15.4 |
|---|---|---|
| Modelo retirado en `agent/brain.py` | App no chateaba (modelo retirado 20-abr-2026) | Usa `claude-sonnet-4-6`, centralizado en `config.py` |
| Crash en `_build_executive_report` y otros | `AttributeError: 'dict' object has no attribute 'upper'` cuando LLM devuelve secciones como list[dict] | Normalización robusta con `_normalize_sections()` |
| Messaging engañoso "cobertura LATAM" | Sugería Brasil cubierto | Aclarado como "cobertura hispanohablante" |
| Cobertura de tests de Word | 2 tests (~5%) | 111 tests parametrizados |

---

## Cambios

### `config.py` — NUEVO archivo

Configuración centralizada de model IDs y rutas. Resuelve el problema histórico de que distintos módulos quedaran apuntando a modelos diferentes (en v15.3, `brain.py` tenía el modelo retirado mientras `word_translator.py` y `llm_legal_generator.py` ya estaban actualizados).

```python
DEFAULT_CHAT_MODEL       = "claude-sonnet-4-6"
DEFAULT_TITLE_MODEL      = "claude-haiku-4-5-20251001"
DEFAULT_VISION_MODEL     = DEFAULT_CHAT_MODEL
DEFAULT_TRANSLATOR_MODEL = DEFAULT_CHAT_MODEL
DEFAULT_LEGAL_MODEL      = DEFAULT_CHAT_MODEL
```

Override por variable de entorno sigue funcionando: `WEPAI_MODEL_CHAT`, `WEPAI_MODEL_TRANSLATOR`, `WEPAI_MODEL_LEGAL`, `WEPAI_MODEL_VISION`, `WEPAI_MODEL_TITLE`.

### `agent/brain.py` — fix crítico

```diff
-MODEL_CHAT   = os.environ.get("WEPAI_MODEL_CHAT",   "claude-sonnet-4-20250514")  # ← retirado
-MODEL_VISION = os.environ.get("WEPAI_MODEL_VISION", "claude-sonnet-4-20250514")
-MODEL_TITLE  = os.environ.get("WEPAI_MODEL_TITLE",  "claude-haiku-4-5-20251001")
+from config import (
+    DEFAULT_CHAT_MODEL  as MODEL_CHAT,
+    DEFAULT_VISION_MODEL as MODEL_VISION,
+    DEFAULT_TITLE_MODEL  as MODEL_TITLE,
+)
```

`claude-sonnet-4-20250514` fue retirado por Anthropic el 20-abr-2026. La app no podía chatear ni hacer vision verification en producción.

### `office/word_translator.py` — refactor

Consume `config.DEFAULT_TRANSLATOR_MODEL` en vez de leer la env var directamente. Mismo comportamiento, una sola fuente de verdad.

### `office/llm_legal_generator.py` — refactor

Idem: consume `config.DEFAULT_LEGAL_MODEL`.

### `ui/chat_window.py` — fix de display

```diff
-text="Claude Sonnet 4"
+text="Claude Sonnet 4.6"
-text="claude-sonnet-4-20250514"
+text="claude-sonnet-4-6"
```

La pantalla de configuración mostraba al usuario el string de un modelo retirado.

### `office/word_controller.py` — fix crítico del crash

Nuevo helper `_normalize_sections(secciones)` que acepta `list[str]` o `list[dict]` y devuelve siempre `list[str]`. Aplicado en 4 builders que iteraban `for sec in secciones` con `sec.upper()` o `sec.lower()`:

- `_build_executive_report` (línea ~1796) — el crash original
- `_build_business_plan` (línea ~2690) — descubierto por los tests parametrizados
- `_build_internal_memo` (línea ~2610) — el dict se imprimía como string
- `_build_generic_word` (línea ~2832)
- `_build_affidavit` (línea ~3654)

Edge cases cubiertos:
- `secciones = None`
- `secciones = []`
- `secciones = [42, None, "OK", {"name": "Otro"}]` (tipos mezclados/raros)
- `secciones = [{"titulo": "", "contenido": "..."}]` (dicts vacíos descartados)
- `secciones = [{"titulo": "ES", "title": "EN"}]` (prioridad ES sobre EN)

### `CHANGELOG_v15.3.md` — corrección de messaging

```diff
-*v15.3 cierra la cobertura LATAM...*
+*v15.3 cierra la cobertura hispanohablante. ... **Brasil, países anglosajones
+y otros mercados no-hispanohablantes caen al fallback GENERIC con placeholders
+visibles — son trabajo futuro.***
```

Brasil no estaba cubierto (es lusófono, no hispano) y el messaging anterior lo sugería.

### `tests/test_word_builders.py` — NUEVO archivo

Primera batería seria de tests del módulo Word. 111 tests en 6 grupos:

1. **Smoke** (38 tests): cada `template_id` registrado en `WORD_TEMPLATES` genera un .docx sin crashear.
2. **Regresión v15.4** (38 tests): cada builder acepta `secciones` como `list[dict]`.
3. **Helper unitario** (12 tests): `_normalize_sections` con todos los edge cases.
4. **Preservación de datos** (1 test): los datos del usuario aparecen en el .docx final.
5. **Cobertura de países** (20 tests): contrato laboral en los 20 países hispanohablantes sin placeholders del fallback GENERIC.
6. **Config central** (2 tests): ningún modelo apunta al retirado, brain.py consume config central.

Ejecutar:

```bash
pytest tests/test_word_builders.py -v
```

Resultado actual: `111 passed in 9.66s`. **Cero dependencia de red, API key o Mac.** Listo para CI.

---

## Bugs descubiertos por los nuevos tests

Los tests parametrizados encontraron **un cuarto builder** con el mismo bug del `sec.upper()` que no había sido detectado en la auditoría inicial — `_build_business_plan` línea ~2693. Ya está parchado en v15.4.

Esto valida el valor de los tests parametrizados: con un solo archivo de tests detectaron 4 instancias del mismo patrón de bug a través de 38 builders. Sin ellos, el cuarto bug habría llegado a producción.

---

## Lo que queda pendiente

Estos son los hallazgos de la auditoría que **NO** se atacan en v15.4 (planificados para sprints siguientes):

| Hallazgo | Severidad | Sprint sugerido |
|---|---|---|
| Refactor del monolito `word_controller.py` (6,188 líneas) en módulos | Alta | Sprint 4-6 |
| Dispatcher por keywords frágil con `KEYWORD_DISPATCH` declarativo | Alta | Sprint 3 |
| Validación de input con `pydantic` en `generate_word` | Media | Sprint 3 |
| Headings ES/EN mezclados en builders bilingües (NDA, etc.) | Alta | Sprint 7+ |
| Costo de traducción post-hoc (~$2.50/mes/usuario Pro a 50 docs EN) | Alta | Sprint 7+ |
| Auto-ruteo a LLM-driven cuando país no está en LEGAL_CONFIG | Crítica para messaging | Sprint 2 |
| Rate limiting de llamadas a API por usuario | Media | Sprint 4 |

---

## Métricas del release

| Métrica | v15.3 | v15.4 | Δ |
|---|---|---|---|
| Líneas en `word_controller.py` | 6,188 | 6,216 | +28 (helper) |
| Builders con normalización de secciones | 0/38 | 5/38 (los vulnerables) | +5 |
| Modelos activos correctos | 2/3 | 3/3 | +1 |
| Tests automatizados de Word | 2 | 111 | +109 |
| Cobertura builders ejercitada por tests | ~3% | 100% | +97 pp |
| Archivos modificados | — | 5 + 2 nuevos | — |

---

*v15.4 es un release de estabilidad sin nuevas features. La app vuelve a funcionar en producción (modelo vigente), el crash del reporte ejecutivo está resuelto, y la primera red de seguridad de tests automatizados existe.*
