# Reporte de Auditoría de UI — WEP AI v16.9.1

**Proyecto:** WEP AI · Asistente de Office con IA para Colombia
**Componente auditado:** Interfaz de escritorio — `ui/chat_window.py` (CustomTkinter)
**Versión:** v16.9 → corregida a **v16.9.1**
**Fecha:** 3 de junio de 2026
**Método:** análisis estático (pyflakes) + render en vivo bajo display virtual
(Xvfb) con aislamiento de la capa de datos (`storage.db` simulada).

> **Alcance y límites:** esta auditoría cubre la integridad de ejecución y la
> construcción visual de la UI de escritorio. NO evalúa lógica de negocio,
> seguridad, ni fallos en runtime que dependan de datos reales (red, DB, Stripe,
> sistema de archivos del usuario). Pueden existir defectos fuera de este alcance.

---

## 1. Resumen ejecutivo

Se auditó la interfaz de escritorio renderizándola realmente, no solo leyéndola.
Esto reveló un bug **crítico** que impedía el arranque y que el análisis estático
por sí solo no marca (los atributos faltantes solo fallan en ejecución). En total
se detectaron **5 bugs** y se corrigieron los 5, sin regresiones.

| Métrica | Resultado |
|---|---|
| Bugs detectados | 5 (1 crítico · 3 altos · 1 medio) |
| Bugs corregidos | 5 / 5 ✅ |
| pyflakes (antes → después) | 3 → 0 ✅ |
| Arranque de la app | crash → limpio ✅ |
| Regresiones introducidas | 0 |

---

## 2. Inventario de bugs

| # | Bug | Ubicación | Severidad | Síntoma | Estado |
|---|-----|-----------|-----------|---------|--------|
| 1 | `_prefs` / `_prefs_path` nunca asignados | `__init__`, `_load_prefs` | Crítico | `AttributeError` al arrancar — la app no abre | ✅ |
| 2 | `self.state` pisa `Tk.state()` | `__init__`, `_set_state` + 2 | Alto | `TypeError: 'str' object is not callable` | ✅ |
| 3 | Nombre indefinido `FG` | ventana de pago | Alto | `NameError` al abrir "esperando pago" | ✅ |
| 4 | Nombre indefinido `FG_DIM` | ventana de pago | Alto | `NameError` al abrir "esperando pago" | ✅ |
| 5 | Cursor de tarjeta salta al final | `fmt_card` | Medio | No se puede editar el nº en medio | ✅ |

---

## 3. Detalle y causa raíz

**#1 (crítico).** Atributos leídos pero nunca inicializados. Patrón típico de una
refactorización en la que se eliminó la asignación pero no las lecturas. Es el
caso clásico que el render detecta y el linter no: pyflakes no marca atributos de
instancia faltantes porque podrían definirse en cualquier método.

**#2 (alto).** Colisión de espacio de nombres con la clase base. `CTk` hereda de
`Tk`, que expone `state()` como método; asignarle un string lo inutiliza para
todo el framework. La corrección renombra a `self.app_state`.

**#3 / #4 (altos).** Constantes de color huérfanas (`FG`, `FG_DIM`) de una paleta
anterior, no migradas a los tokens actuales (`WHITE`, `GRAY1`…). Solo se disparan
al abrir la ventana de pago, por eso no se notaban en el flujo principal.

**#5 (medio).** Variable calculada y descartada: se medía la posición del caret
pero no se restauraba tras reescribir el campo. Defecto de UX, no de ejecución.

---

## 4. Verificación post-corrección

1. `py_compile` → OK.
2. `pyflakes` → 0 hallazgos (antes: `FG`, `FG_DIM`, variable `pos` sin usar).
3. Búsqueda de residuos: 0 ocurrencias de `self.state`, `FG`, `FG_DIM`.
4. Render en vivo **sin parches de prueba**: `WEPAIApp` se instancia sola, expone
   `app_state == "idle"`, construye sidebar + header + chat + preview + input, y
   lanza el onboarding de primer arranque. El `TypeError` del tracker de DPI que
   aparecía antes ya no se produce.

---

## 5. Recomendaciones de seguimiento (fuera de alcance de este patch)

- Separar `chat_window.py` (≈2.900 líneas) en `login.py`, `register.py`,
  `main_window.py`, `dialogs.py` y un `theme.py` para los tokens de color.
- Encapsular el acceso a `self.chat_scroll._parent_canvas` (API privada de CTk)
  en un helper con `hasattr`, para no romper con futuras versiones.
- Reemplazar el cálculo manual de altura de burbujas por `CTkLabel` con
  `wraplength` (autoajuste), eliminando el conteo frágil de ~50 chars/línea.
- Añadir binds `<Control-n>` / `<Control-comma>` para Windows y Linux (hoy solo
  hay atajos de macOS).
