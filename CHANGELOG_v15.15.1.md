# CHANGELOG v15.15.1 — Auditoría Word: fixes de dispatch

Fecha: 28 de mayo de 2026
Módulo: Word (`office/word_controller.py`, `agent/brain.py`)
Tipo: corrección de bugs de enrutamiento (no afecta builders ni estabilidad)

## Contexto

Auditoría del módulo Word sobre v15.15.0. El smoke test de 228 combinaciones
(38 builders × 6 países) seguía pasando sin crashes, pero se detectaron dos
defectos de **enrutamiento** (dispatch) que entregaban el documento equivocado
al usuario en flujos de lenguaje natural. Ninguno de los dos estaba reportado
en el informe de continuidad v15.15.0.

---

## Fix 1 — Colisión de keyword "trabajador" / "empleo" (CRÍTICO)

### Problema
La rama de keywords de `employment_contract` en el dispatcher de fallback
(`generate_word`) incluía los términos genéricos `"trabajador"` y `"empleo"`.
Como esta rama se evalúa ~90 líneas antes que las ramas de documentos laborales
específicos, **capturaba cualquier petición que mencionara al trabajador**.

Impacto medido (peticiones en lenguaje natural sin `template_id` explícito):

| Petición del usuario              | Esperado          | Generaba (bug)        |
|-----------------------------------|-------------------|-----------------------|
| "amonestación al trabajador..."   | written_warning   | ❌ Contrato de trabajo |
| "carta de despido del trabajador" | termination_letter| ❌ Contrato de trabajo |
| "oferta de empleo para trabajador"| offer_letter      | ❌ Contrato de trabajo |
| "finiquito del trabajador"        | settlement        | ❌ Contrato de trabajo |

El bug solo se manifestaba en el path de fallback por keywords (cuando
`brain.py` omite `template_id`). Con `template_id` explícito el dispatch ya era
correcto.

### Solución
Se eliminaron los keywords genéricos `"trabajador"` y `"empleo"` de la rama
`employment_contract`, dejando solo frases inequívocas que únicamente describen
un contrato laboral: `"contrato de trabajo"`, `"contrato laboral"`,
`"contrato individual"`, `"contrato de empleo"`, `"employment contract"`,
`"labor contract"`, `"individual employment"`.

Se reforzó adicionalmente la rama `offer_letter` con la variante
`"oferta de empleo"` (el matching por substring no cubría la forma con "de").

### Verificación
7/7 casos de prueba correctos tras el fix, incluyendo que el contrato laboral
legítimo ("contrato de trabajo", "contrato laboral indefinido") sigue
enrutándose correctamente a `employment_contract`.

---

## Fix 2 — Builders huérfanos no anunciados al LLM

### Problema
Tres builders existían en `WORD_TEMPLATES` pero no estaban listados en el
SYSTEM_PROMPT de `agent/brain.py`, por lo que el LLM nunca emitía su
`template_id` y dependían exclusivamente del fallback por keywords:
`affidavit`, `remote_work_agreement`, `written_warning`.
(`written_warning`, además, era víctima del Fix 1.)

### Solución
Se añadieron los tres como entradas explícitas de primera clase en la lista de
IDs Word del SYSTEM_PROMPT, con su descripción, para que el LLM los reconozca
y emita el `template_id` directamente sin depender del fallback.

### Verificación
Diff de IDs: los 38 builders de `WORD_TEMPLATES` quedan ahora cubiertos en
`brain.py` (antes faltaban 3). Los tres generan el documento correcto vía
`template_id` explícito.

---

## Regresión
- Smoke test completo Word: **228/228 OK, 0 crashes** (sin cambios respecto a
  v15.15.0).
- `agent/brain.py` compila sin errores de sintaxis.
- Contrato laboral legítimo: enruta correctamente.

## Archivos modificados
- `office/word_controller.py` — rama `employment_contract` y `offer_letter`.
- `agent/brain.py` — lista de IDs Word del SYSTEM_PROMPT.

## Pendiente (no incluido en este parche)
- Excel: los builders `kpi_tracker`, `presupuesto_anual` y `ratios_financieros`
  siguen inalcanzables (no anunciados + colisión de keywords). Requiere su
  propio parche.
- Reconciliar cifras del informe: 20 países reales (no 24), 12 verificadores
  (no 10).
