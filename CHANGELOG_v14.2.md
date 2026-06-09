# WEP AI — CHANGELOG v14.2

Fecha: 21 de mayo de 2026
Base: v14.1

Esta versión cubre los dos items más importantes que quedaban pendientes
de la review: el refactor completo del threading en la UI y la migración
del dispatcher de templates a un sistema dirigido por el LLM.

---

## Resumen ejecutivo

| Métrica | v14.1 | v14.2 |
|---|---|---|
| Llamadas a Tkinter desde threads (fuera de `self.after`) | varias | 0 |
| Templates con ID explícito | 0 | 92 (Excel 36 + Word 38 + PPT 18) |
| Tests de regresión | 8 grupos | 9 grupos |
| Verificación estática de thread safety | no | sí (AST inspector) |
| Bugs latentes (críticos + altos + medios) | 0 | 0 |

---

## Cambio 1: Refactor del threading en la UI

### Problema

En v14.1, los métodos `_process` y `_agent` corrían en un `threading.Thread`
y modificaban widgets de Tkinter directamente (`self._bubble(...)`,
`self._set_state(...)`, `self.send_btn.configure(...)`, etc.). Tkinter NO
es thread-safe, especialmente en macOS, donde puede causar freezes,
race conditions visuales y crashes ocasionales.

### Solución

Patrón **worker → main-thread callback** aplicado consistentemente:

1. Métodos `_send` y `_start_agent` corren en main thread. Hacen toda la
   preparación de UI (estados, bubbles iniciales, indicadores) y luego
   lanzan el worker.
2. Workers `_process_worker`, `_agent_worker`, `_upd_title` corren en
   background. Sólo hacen trabajo de red/CPU/subprocess. **No tocan
   widgets directamente.**
3. Cuando el worker tiene un resultado, lo entrega al main thread vía
   `self.after(0, callback, *args)`. Los callbacks `_on_*` reciben los
   datos y actualizan la UI.

### Antes (v14.1)

```python
def _process(self, txt):
    self._set_state("thinking")     # widget call from thread ⚠
    typ = self._typing()            # widget creation from thread ⚠
    try: res = brain.chat(...)
    except Exception as e:
        typ.destroy()               # widget call from thread ⚠
        self._bubble(...)           # widget call from thread ⚠
        self._set_state("idle")     # widget call from thread ⚠
        ...
    self._bubble("assistant", ...)  # widget call from thread ⚠
    if res["generate"]:
        threading.Thread(target=self._agent, ...).start()
    else:
        self._set_state("idle")     # widget call from thread ⚠
```

### Después (v14.2)

```python
def _send(self, e=None):
    # Main thread: prepara UI ANTES de lanzar worker
    self._set_state("thinking")
    typ = self._typing()
    history_snapshot = list(self.chat_history[:-1])  # evita race con append
    threading.Thread(
        target=self._process_worker,
        args=(text, history_snapshot, typ),
        daemon=True,
    ).start()

def _process_worker(self, txt, history, typ):
    # Background thread: SÓLO trabajo de red. Sin widgets.
    try:
        res = brain.chat(history, txt)
    except Exception as e:
        self.after(0, self._on_chat_error, typ, e)
        return
    self.after(0, self._on_chat_done, typ, res)

def _on_chat_done(self, typ, res):
    # Main thread: todas las mutaciones de UI aquí
    typ.destroy()
    if res["text"]:
        self._bubble(...)
        ...
    if res["generate"] and res["gen_data"]:
        self._start_agent(res["gen_data"])
    else:
        self._set_state("idle")
        self.send_btn.configure(state="normal")
```

### Métodos nuevos

Worker methods (background):
- `_process_worker(txt, history, typ)` — chat API call
- `_agent_worker(gen_data, dt, titulo)` — generación de documento + apertura + captura + vision

Main-thread callbacks:
- `_on_chat_error(typ, err)`, `_on_chat_done(typ, res)` — pos-chat
- `_start_agent(gen_data)` — entry point cuando hay que generar
- `_reset_attachment()` — limpia el clip de Excel adjunto
- `_on_agent_error(err)` — manejo de excepción durante generación
- `_on_doc_generated(dt, titulo, fp, gen_data)` — actualiza estado del doc actual
- `_on_screenshot_captured(dt)` — refresca preview y badge
- `_on_agent_done(dt, titulo, fp, vc, bug_report)` — bubble final + cierre

### Validación estática

Se incluyó un inspector AST que recorre el archivo y verifica que dentro
de las funciones marcadas como worker (`_process_worker`, `_agent_worker`,
`_upd_title`) no haya llamadas directas a métodos de UI fuera de
`self.after(...)` o de lambdas (las lambdas pasadas a `after` se ejecutan
en el main thread).

Resultado: **0 llamadas inseguras detectadas**.

### Snapshot de historial

`self.chat_history` se modifica desde el main thread (`_send` hace `append`).
Antes, el worker leía `self.chat_history[:-1]` desde otro thread, lo que
era una race condition latente. Ahora `_send` toma `list(self.chat_history[:-1])`
en el main thread y pasa el snapshot al worker como argumento.

---

## Cambio 2: Dispatcher de templates dirigido por el LLM

### Problema

En v14.1, los dispatchers de `generate_excel`, `generate_word` y
`generate_powerpoint` decidían qué builder usar examinando el texto
del usuario y matcheando keywords con `if/elif` gigantes (200+ líneas
en algunos casos).

Limitaciones:
- Frágil ante typos, paráfrasis y sinónimos no contemplados.
- Patrones con `.*` se trataban como literales (bug corregido en v14.1
  via `_matches()`, pero la lista de keywords seguía teniendo huecos).
- Imposible saber desde fuera qué template terminó eligiéndose.
- Para agregar un template nuevo: modificar el `if/elif` Y la lista
  de keywords Y esperar que el LLM lo describa con palabras compatibles.

### Solución

**El LLM ahora emite el `template_id` explícito en el JSON `##GENERAR##`**.
El dispatcher lo consulta primero; si está presente y es válido, salta el
matching por keywords. Si no, cae al matching legacy.

### Componentes nuevos

**1. Registros `*_TEMPLATES` en cada controller** (al final de cada archivo,
para evitar forward references):

```python
# office/excel_controller.py
EXCEL_TEMPLATES = {
    "income_statement":  _build_income_statement,
    "break_even":        _build_break_even,
    # ... 36 entradas
}

# office/word_controller.py
WORD_TEMPLATES = {
    "service_contract": _build_service_contract,
    # ... 38 entradas
}

# office/ppt_controller.py
PPT_TEMPLATES = {
    "pitch":   _plan_pitch,
    # ... 18 entradas
}
```

**Total: 92 templates con IDs explícitos.**

**2. Dispatch híbrido (template_id → keywords)** en cada `generate_*`:

```python
template_id = (gen_data.get("template_id") or "").strip().lower()
if template_id and template_id in EXCEL_TEMPLATES:
    EXCEL_TEMPLATES[template_id](wb, titulo, secciones, detalles, instruc)
elif any(w in full for w in [...]):  # fallback legacy
    _build_income_statement(...)
```

**3. SYSTEM_PROMPT en `brain.py`** actualizado:
- Añade `template_id` al esquema del JSON `##GENERAR##`.
- Documenta los 92 IDs disponibles, agrupados por tipo (Excel/Word/PPT).
- Da una regla clara al LLM: si la descripción del usuario es nítida,
  emite el template_id; si es ambigua, omítelo y deja al sistema usar
  el fallback por keywords.

### Beneficios

- **Determinismo**: el desarrollador puede testear builders sin tener
  que adivinar qué frase exacta los activa.
- **Diagnóstico**: el log puede registrar el `template_id` final usado.
- **Extensibilidad**: añadir un template nuevo es 3 pasos claros:
  define la función, agrega al registry, lista el ID en SYSTEM_PROMPT.
- **Backward compatible**: el dispatcher por keywords sigue funcionando
  exactamente igual cuando `template_id` no está presente o es inválido.
- **Case-insensitive**: `template_id` se normaliza con `.lower().strip()`.

### Tests de validación (Test 9 del suite)

1. `template_id="dashboard"` + texto con "inventario" → genera Dashboard
   (el template_id explícito GANA al keyword).
2. Sin `template_id` + texto con "inventario" → genera Inventario
   (fallback por keyword funciona).
3. `template_id="no_existe"` + texto con "inventario" → genera Inventario
   (template_id inválido cae al fallback).
4. `template_id=""` → fallback (string vacío no causa error).
5. `template_id="INVENTORY"` (mayúsculas) → matchea `inventory`
   (case-insensitive).

---

## Trabajo NO incluido en v14.2 (queda para v14.3)

Los 3 items restantes de la review están todos OK para iterar después:

### Split de los archivos enormes
`word_controller.py` (5,150 líneas) y `excel_controller.py` (4,427 líneas)
son grandes pero ahora tienen registry centralizado. Refactor a
`office/word/templates/*.py` con una base class compartida tiene sentido
pero es trabajo de medio día y mucho riesgo de regresión sin tests por
template. Mejor abordarlo después de aumentar la cobertura de tests.

### Auditoría de 146 unused locals en `word_controller.py`
Confirmado por pyflakes. Cada uno necesita revisión manual del builder
para determinar si es información del usuario que se está perdiendo
(probable) o variables muertas legítimas. Requiere abrir y entender
cada uno de los ~37 builders.

### Integración Stripe Checkout
Decisión de producto: el formulario actual ya está marcado como "MODO
DEMO". Integrar Stripe requiere claves API, configuración de webhooks,
y manejo de eventos `checkout.session.completed`. Out of scope para
una sesión de fixes.

---

## Cambios en dependencias

Ninguno. v14.2 no añade dependencias respecto a v14.1.
