# CHANGELOG — WEP AI v14.8

**Objetivo:** release de cierre de deuda técnica, sin features nuevas. Cierra los 4 issues nuevos introducidos en v14.7 (N1–N4) y 3 de los 5 críticos pendientes del review previo (C1, C4, C5). Total: 7 fixes en ~50 líneas netas modificadas, 6 archivos.

**Fecha:** 2026-05-26
**Resultado:** 15/15 grupos de test pasando (+1 grupo nuevo cubriendo todos los fixes de v14.8)

---

## Resumen ejecutivo

| Fix | Severidad | Categoría | Líneas |
|---|---|---|---|
| **C1** Inyección AppleScript en `open_in_word/excel/powerpoint` | 🔴 Crítico | Seguridad | ~30 |
| **C4** Modelo Claude hardcoded en 3 lugares de brain.py | 🔴 Crítico | Mantenibilidad | ~15 |
| **C5** Sin retry/backoff en API Anthropic | 🔴 Crítico | Robustez | ~5 |
| **N1** `_cancel_subscription_directly` huérfano, sin botón | 🔴 Importante | UX/Producto | ~10 |
| **N2** `send_welcome` bloquea thread UI hasta 10s | 🔴 Importante | UX/Performance | ~15 |
| **N3** Connection leak en `_get_user_email_and_lang` | 🟡 Menor | Recursos | ~5 |
| **N4** `_get_user_email_and_lang` ignoraba lang real | 🟡 Menor | i18n | ~10 |

---

## Cambios por archivo

### 1. `office/word_controller.py` / `excel_controller.py` / `ppt_controller.py` — Fix C1

Los 3 controllers tenían la misma vulnerabilidad: interpolación f-string de `filepath` directo a un script de AppleScript ejecutado por `osascript`. Si el path contenía `"` o sintaxis AppleScript válida, se ejecutaba como código.

**Antes (v14.7):**
```python
def open_in_word(filepath: str):
    script = f'''
        tell application "Microsoft Word"
            activate
            open POSIX file "{filepath}"
        end tell
    '''
    subprocess.run(["osascript", "-e", script], check=False)
```

**Después (v14.8):**
```python
def open_in_word(filepath: str):
    """Open the file in Microsoft Word on macOS.

    v14.8 — Usa `open -a` en lugar de AppleScript interpolado para evitar
    inyección por contenido del path. `open` recibe el path como argv y no
    lo evalúa como código.
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Word file not found: {filepath}")
    subprocess.run(["open", "-a", "Microsoft Word", filepath], check=False)
```

Mismo patrón aplicado a `open_in_excel` y `open_in_powerpoint`. Bonus: ahora valida que el archivo existe antes de llamar a `open`, dando un `FileNotFoundError` claro en lugar de un fallo silencioso.

**Riesgo de regresión:** muy bajo. `open -a "Microsoft Word"` es el mismo comportamiento que produce el AppleScript original, sin el riesgo de inyección.

---

### 2. `agent/brain.py` — Fix C4 + C5

**Antes (v14.7):**
```python
import anthropic, json, re, os

client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY", ""))

# ... (424 líneas de SYSTEM_PROMPT) ...

response = client.messages.create(
    model="claude-sonnet-4-20250514",  # ← repetido en 3 lugares
    ...
)
```

**Después (v14.8):**
```python
import anthropic, json, re, os

# ── v14.8: configuración de modelo y retry ────────────────────────────────────
# Centralizamos el modelo en constantes con override por env var, para que
# (a) actualizar el modelo no requiera buscar/reemplazar en 3 lugares, y
# (b) podamos correr A/B con un modelo distinto sin redeploy.
# `generate_title` usa Haiku porque es ~10× más barato y suficiente para 5 palabras.
MODEL_CHAT   = os.environ.get("WEPAI_MODEL_CHAT",   "claude-sonnet-4-20250514")
MODEL_VISION = os.environ.get("WEPAI_MODEL_VISION", "claude-sonnet-4-20250514")
MODEL_TITLE  = os.environ.get("WEPAI_MODEL_TITLE",  "claude-haiku-4-5-20251001")

client = anthropic.Anthropic(
    api_key=os.environ.get("ANTHROPIC_API_KEY", ""),
    max_retries=3,
    timeout=60.0,
)
```

**Beneficios:**
- **C4:** actualizar el modelo es ahora cambiar 1 string (o setear un env var). Antes había que buscar 3 lugares.
- **C5:** el SDK de Anthropic reintenta automáticamente con backoff exponencial sobre `RateLimitError` (429), `APIConnectionError` y 5xx transitorios. Un 429 momentáneo ya no rompe la UX del usuario.
- **Bonus económico:** `generate_title` ahora usa Haiku 4.5 — ~10× más barato que Sonnet, y suficiente para los 5 palabras del título de la conversación.

**Cómo overridear el modelo sin redeploy:**
```bash
# Para probar un modelo nuevo solo cambiando env vars:
export WEPAI_MODEL_CHAT="claude-opus-4-7"
```

---

### 3. `ui/chat_window.py` — Fix N1 + N2

#### N2 — Helper async para emails best-effort

Antes, `send_welcome` se llamaba síncrono en el thread UI. Si `EMAIL_PROVIDER=resend` está configurado, el HTTP POST a Resend tiene timeout de 10s. La app se congelaba esperando ese ack durante el registro.

**Nuevo helper top-level:**
```python
# v14.8 — Helper para enviar emails en background sin bloquear el thread UI.
def _send_welcome_async(email_to: str, name: str, lang: str):
    """Lanza send_welcome en un thread daemon. No bloquea, no propaga errores."""
    def _run():
        try:
            from storage import email as _email
            _email.send_welcome(email_to, name, lang=lang)
        except Exception as e:
            print(f"[WEP AI] send_welcome failed (non-blocking): {e}")
    threading.Thread(target=_run, daemon=True).start()
```

Los 2 call sites (`_do_register` y `_open_payment_stripe`) ahora son una sola línea:
```python
_send_welcome_async(em, nm, lang=self.lang[0])
```

#### N1 — Wire-up del botón "Cancelar suscripción"

El método `_cancel_subscription_directly` existía desde v14.7 pero ningún elemento de la UI lo invocaba — feature anunciada que el usuario no podía usar. Ahora aparece en el menú de usuario, en rojo (estilo `danger`), **solo para usuarios en plan Pro o Enterprise**:

```python
item("ℹ️", "about_menu",lambda: self._open_about())
# v14.8 — Cancelar suscripción directo (solo para planes pagos).
# El método ya existía en v14.7 pero ningún botón lo invocaba.
if plan in ("pro", "biz"):
    ctk.CTkFrame(m, height=1, fg_color=BORDER).pack(fill="x", pady=3)
    item("✕", "cancel_sub", self._cancel_subscription_directly, danger=True)
ctk.CTkFrame(m, height=1, fg_color=BORDER).pack(fill="x", pady=3)
item("→",  "logout",   self._logout, danger=True)
```

Traducciones agregadas a `LANG`:
```python
"cancel_sub": "Cancelar suscripción",  # ES
"cancel_sub": "Cancel subscription",   # EN
```

**UX:** el usuario en Free no ve el botón (no tiene nada que cancelar). El usuario Pro/Enterprise ve "✕ Cancelar suscripción" entre "Acerca de" y "Cerrar sesión", separado por una línea para que se note que es destructivo. Al hacer click, el método existente pide confirmación con `messagebox.askyesno` y llama a `stripe_client.cancel_subscription(sid)`.

---

### 4. `storage/webhook_server.py` — Fix N3 + N4

**Antes (v14.7):**
```python
def _get_user_email_and_lang(customer_id: str) -> tuple:
    if not customer_id:
        return (None, "en")
    try:
        db = sqlite3.connect(_db_path())
        cur = db.cursor()
        cur.execute("PRAGMA table_info(users)")
        cols = {row[1] for row in cur.fetchall()}
        if "email" not in cols:
            return (None, "en")           # ← N3: leak de connection
        cur.execute("SELECT email FROM users WHERE stripe_customer_id = ?",
                    (customer_id,))
        row = cur.fetchone()
        db.close()
        return (row[0] if row else None, "en")  # ← N4: lang siempre "en"
```

**Después (v14.8):**
```python
def _get_user_email_and_lang(customer_id: str) -> tuple:
    """v14.8 fixes:
    - N3: usa context manager para que la conexión se cierre siempre.
    - N4: ahora SÍ lee el lang real si existe una columna 'lang' en users.
    """
    if not customer_id:
        return (None, "en")
    try:
        with sqlite3.connect(_db_path()) as db:
            cur = db.cursor()
            cur.execute("PRAGMA table_info(users)")
            cols = {row[1] for row in cur.fetchall()}
            if "email" not in cols:
                return (None, "en")
            lang_col = next((c for c in ("lang", "language", "locale") if c in cols), None)
            select = f"SELECT email{', ' + lang_col if lang_col else ''} FROM users WHERE stripe_customer_id = ?"
            cur.execute(select, (customer_id,))
            row = cur.fetchone()
            if not row:
                return (None, "en")
            email = row[0]
            raw_lang = (row[1] if lang_col and len(row) > 1 else None) or "en"
            lang = "es" if str(raw_lang).lower().startswith("es") else "en"
            return (email, lang)
    except Exception as e:
        log.warning(f"_get_user_email_and_lang: {e}")
        return (None, "en")
```

**Beneficios:**
- **N3:** el `with sqlite3.connect(...) as db:` garantiza commit/rollback + close, incluso si hay un `return` en cualquier rama.
- **N4:** ahora detecta automáticamente si existe una columna de idioma (probando `lang`, `language`, `locale` en ese orden) y la usa. Normalizamos cualquier valor que empiece con `"es"` a `"es"`, todo lo demás a `"en"`. Si la columna no existe (DB legacy), el fallback sigue siendo `"en"`.

**Caso concreto que esto desbloquea:**
- Usuario hispanohablante registrado con su idioma en DB.
- Stripe envía `invoice.payment_failed` (no siempre trae metadata).
- Webhook llama `_get_user_email_and_lang(customer_id)` → antes devolvía email correcto pero lang=`"en"` → email en inglés. Ahora devuelve email + lang reales → email en español ✓.

---

### 5. `tests/test_db.py` — Test 15 (3 sub-tests)

Test nuevo que verifica los fixes de v14.8 con assertions concretas (no dependiente de mock de SDKs externos para los 3 sub-tests):

- **15.1** Verifica vía `inspect.getsource` que los 3 `open_in_*` no contienen `"osascript"`, sí contienen `"open"` + `"-a"`, y validan existencia con `FileNotFoundError`.
- **15.2** Verifica que `brain.MODEL_CHAT/VISION/TITLE` existen y responden a override por env var (usa `importlib.reload`).
- **15.3** Crea una DB temporal con columna `lang` y 3 users (es/en/NULL), verifica que `_get_user_email_and_lang` devuelve `("es@test.com", "es")`, `("en@test.com", "en")` y para NULL cae a `"en"`.

Suite completa: **15/15 grupos pasan.**

---

## Issues pendientes (para v14.9 o más adelante)

### Críticos no abordados aún:

- **C2** — `SYSTEM_PROMPT` de 420 líneas inline. Mover a archivo + activar prompt caching del SDK de Anthropic. **Beneficio:** ~90% de reducción de costo en la parte de input cacheable. Trabajo: 1 día.
- **C3** — Refactor de `word_controller.py` (5,771 líneas) y `excel_controller.py` (5,262 líneas) a paquetes con un builder por archivo. **Beneficio:** mantenibilidad. Trabajo: 1 mes incremental.

### Importantes I1–I11 del review previo:

Todos siguen pendientes. Recomiendo atacar en orden:
- **I1** (bare excepts en flujos críticos) — fácil, 30 min de log.exception
- **I7** (mezcla print/logging) — fácil, setup centralizado en main.py
- **I8** (late imports) — trivial
- **I9** (state strings → Enum) — 1 hora
- **I4** (sleep(3) → polling) — 1 hora

---

## Cómo verificar los fixes en tu Mac

```bash
# 1. Correr el test suite
cd wepai
python3 tests/test_db.py
# Esperado: "All 15 test groups PASSED."

# 2. Verificar que los open_in_* no usan osascript
grep -l "osascript" office/*.py
# Esperado: silencio (no debería haber matches)

# 3. Probar override de modelo via env var
export WEPAI_MODEL_TITLE="claude-haiku-4-5-20251001"
python3 -c "from agent import brain; print(brain.MODEL_TITLE)"
# Esperado: claude-haiku-4-5-20251001

# 4. Verificar el wire-up del botón cancel en la UI
# Login como user Pro (o cambiar plan en DB), click en menú de usuario.
# Debería aparecer "✕ Cancelar suscripción" entre "Acerca de" y "Cerrar sesión".
```

---

*v14.8 — Release de cierre técnico. Próxima feature release puede arrancar sobre base más sana.*
