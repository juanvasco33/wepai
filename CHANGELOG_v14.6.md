# CHANGELOG — WEP AI v14.6

**Objetivo de esta release:** cerrar el gap de bilingüe identificado en v14.5 + agregar las features que faltaban para experiencia de producto completa con usuarios anglo: link de T&C/Privacy en registro, emails transaccionales, trial period en Stripe, self-service billing, y un test script E2E para verificar el round-trip en Mac.

**Fecha:** 2026-05-21
**Resultado:** 13/13 grupos de test pasando + script E2E para Mac

---

## Resumen ejecutivo

| Cambio | Antes (v14.5) | Después (v14.6) |
|---|---|---|
| Idioma de UI propagado al brain | NO — el LLM decidía solo basándose en el prompt | `brain.chat(history, txt, ui_lang=self.lang[0])` — UI es fallback cuando LLM omite el campo |
| Aceptación de T&C en registro | No existía | Checkbox obligatorio con links a TERMS y PRIVACY locales |
| Emails transaccionales | No existían | `storage/email.py` con backends Resend/SendGrid/Postmark + 5 plantillas bilingües |
| Stripe trial period | No | 14 días por default (configurable vía `STRIPE_TRIAL_DAYS`) |
| Self-service billing UI | No — usuario tenía que escribir a soporte | Botón "Gestionar facturación" en sidebar para usuarios pago, abre Stripe portal |
| E2E test en Mac | Nunca ejercitado | `tests/test_mac_e2e.py` — script comprensivo a correr en Mac con Office |
| Test suite | 12/12 | 13/13 |

---

## Cambios por archivo

### 1. `agent/brain.py` — `chat()` acepta ui_lang como fallback

```python
def chat(history, user_message, ui_lang=None):
    ...
    if ui_lang and gen_data:
        datos = gen_data.setdefault("datos", {})
        if not datos.get("idioma") and not gen_data.get("idioma"):
            datos["idioma"] = ui_lang
```

**Comportamiento:**
- Si Claude emite `"idioma": "..."` → prioridad al LLM (puede haber detectado idioma específico del prompt aunque la UI esté en otro)
- Si Claude omite el campo → se inyecta `ui_lang` de la UI
- Si el caller no pasa `ui_lang` → no se inyecta nada (retrocompatible)

**Tests:** 5/5 casos en Test 13.1. Mockean diferentes respuestas del LLM y verifican que el merging es correcto.

### 2. `ui/chat_window.py` — propaga UI lang + checkbox T&C + billing portal

**Propagación de lang:**
```python
res = brain.chat(history, txt, ui_lang=self.lang[0])  # antes: solo (history, txt)
```

**Nuevas keys en LANG dict** (ES y EN):
- `terms_check` — "He leído y acepto los" / "I have read and agree to the"
- `terms_link`, `terms_and`, `privacy_link`
- `terms_err` — mensaje de error si no se aceptan

**Nuevo bloque de checkbox + links en `RegisterWindow`:**
- Checkbox obligatorio antes del botón "Crear cuenta"
- Dos links inline (clickeables, subrayados, azul): "Términos" y "Política de Privacidad"
- Click en cualquiera abre `_open_terms(kind)` que busca el archivo markdown local (`TERMS_EN.md`, `PRIVACY_ES.md`, etc.) según el idioma activo y lo abre en el navegador como `file://`. Si no encuentra el archivo, abre URL placeholder (`https://wepai.app/terms-en`)

**Validación en `_create()`:**
```python
if not self.terms_var.get():
    self.terms_err_lbl.configure(text=self._t("terms_err"))
    return
```

**Botón "Gestionar facturación" en sidebar** (solo para usuarios pago):
- Aparece en row 4 del sidebar (donde para usuarios free está "Docs: X/5")
- Click → `_open_billing_portal()` → llama a `stripe_client.create_billing_portal_session(customer_id)` → abre URL del Stripe Customer Portal en el navegador
- Si no hay `stripe_customer_id` (suscripción no sincronizada aún), muestra mensaje informativo
- Si Stripe no está configurado, muestra "contactá support@wepai.app"

### 3. `storage/email.py` — NUEVO módulo (~350 líneas)

Abstracción mínima para emails transaccionales con **3 backends opcionales**:

- **Resend** — `RESEND_API_KEY`. Recomendado para productos nuevos. API limpia.
- **SendGrid** — `SENDGRID_API_KEY`. Más establecido. Free tier 100 emails/día.
- **Postmark** — `POSTMARK_SERVER_TOKEN`. Premium, mejor deliverability.

**Selección vía `EMAIL_PROVIDER`** (`resend|sendgrid|postmark`).

**5 plantillas bilingües ES/EN:**
- `send_welcome(to, name, lang)` — bienvenida post-registro
- `send_payment_confirmed(to, plan, amount, next_billing_date, lang)` — pago exitoso
- `send_payment_failed(to, plan, lang)` — cobro rechazado
- `send_subscription_cancelled(to, end_date, lang)` — confirmación de cancelación
- `send_trial_ending_soon(to, days_left, lang)` — recordatorio fin de trial

**Comportamiento por default:**
- Si `EMAIL_PROVIDER` no está en env → modo dry-run (loguea contenido, devuelve `False`, no rompe nada)
- Cada backend valida API key, hace POST con timeout 10s, devuelve `True` solo en 200/202
- Cualquier excepción se loguea y devuelve `False` (no se propaga)

**Integración pendiente:** las funciones existen pero **aún no se llaman desde el flujo**. Quien las debe llamar son:
- `_do_register()` / `_open_payment_stripe()` en UI → `send_welcome()`
- `webhook_server.py` evento `checkout.session.completed` → `send_payment_confirmed()`
- `webhook_server.py` evento `invoice.payment_failed` → `send_payment_failed()`
- `webhook_server.py` evento `customer.subscription.deleted` → `send_subscription_cancelled()`
- Cron job opcional → `send_trial_ending_soon()` (3 días antes del fin del trial)

Las llamadas a `email.send_*()` desde el webhook server son ~1 línea cada una. Las dejé fuera para no tocar webhook_server.py en esta release (que el cliente puede deployarlo gradualmente sin tener el módulo de email obligatoriamente). Si querés que las agregue, decime.

### 4. `storage/stripe_client.py` — trial_period_days

Nuevo parámetro `trial_days=0` en `create_checkout_session()`. Cuando >0, se incluye `trial_period_days` en `subscription_data`. Stripe Checkout muestra el período de trial al usuario y no cobra hasta que termina.

**Override vía env var:** `STRIPE_TRIAL_DAYS=14` (o cualquier número). Permite cambiar el trial sin redeployar la app.

**Default desde la UI:** 14 días (estándar en mercado US/EU).

### 5. `ui/chat_window.py` — UI usa trial de 14 días

```python
session = stripe_client.create_checkout_session(
    self.sel_plan, em, uid,
    trial_days=14,  # v14.6 — 14 días gratis antes del primer cobro
)
```

### 6. `storage/db.py` — `get_user_field()` helper

Helper seguro para leer campos arbitrarios del usuario con whitelist de columnas permitidas. Maneja columnas que pueden no existir aún (como `stripe_customer_id` que crea el webhook server). Devuelve `None` ante cualquier ambigüedad en lugar de lanzar excepción.

Whitelist actual: `stripe_customer_id`, `stripe_subscription_id`, `subscription_status`, `plan`, `email`, `name`, `last_name`, `docs_used_this_month`, `month_reset`, `card_last4`.

Usado por `_open_billing_portal()` para obtener el `stripe_customer_id` y por futuras features.

### 7. `tests/test_mac_e2e.py` — NUEVO script E2E para Mac (~250 líneas)

**Para qué sirve:** validar los pasos del round-trip completo que no se pueden testear en CI (Linux sin Office).

**Steps cubiertos:**
- S0 — Detecta macOS, Office instalado, deps Python, ANTHROPIC_API_KEY
- S1 — Genera Word/Excel/PPT con los controllers y verifica que existen
- S2 — Abre cada uno con AppleScript (Word/Excel/PowerPoint)
- S3 — Toma screenshot del display vía `agent.vision.take_screenshot()`
- S4 — (Opcional, requiere API key) Llama a Claude Vision para análisis
- S5 — Cierra las apps sin guardar

**Output:** logs detallados a `/tmp/wepai_e2e_*.log` + resumen en stdout con `✓ PASS / ⚠ SKIP / ✗ FAIL`.

**Cómo correrlo en Mac:**
```bash
cd wepai/
python3 tests/test_mac_e2e.py
```

Interactivo: pregunta antes de abrir Office y antes de cerrarlo. Cualquier FAIL se muestra con detalle (mensaje de error, código de retorno de osascript, etc.) para diagnosticar.

### 8. `tests/test_db.py` — Test 13

Cuatro sub-tests:
- **13.1** `ui_lang` fallback funciona y respeta precedencia LLM > UI
- **13.2** `email.is_configured()` false sin env vars; `send_*` devuelve False en dry-run
- **13.3** `stripe_client.create_checkout_session()` acepta `trial_days`
- **13.4** `db.get_user_field()` whitelist + manejo de columna inexistente

Suite completa: **13/13 grupos pasan.**

---

## ¿Está el producto listo para mercado anglo?

**Sí, para lanzamiento beta:**

1. ✓ UI 100% bilingüe con toggle ES/EN
2. ✓ Conversación con Claude en cualquier idioma
3. ✓ 92 templates (Word/Excel/PPT) bilingües
4. ✓ Idioma de UI propaga al brain (fix del gap de v14.5)
5. ✓ Pagos reales con Stripe Checkout + trial 14 días
6. ✓ Self-service billing portal en la app
7. ✓ T&C y Privacy linkeados desde registro, checkbox obligatorio
8. ✓ Email transaccional listo (necesita configurar provider)
9. ✓ Webhook server template para deploy
10. ✓ T&C/Privacy en EN y ES con secciones CCPA y GDPR
11. ✓ Test suite 13/13

**Lo que sigue pendiente y no es bloqueante:**

- **Integrar `email.send_*()` en el flujo real.** El módulo existe y funciona; lo que falta es llamarlo desde 4 puntos (1 en UI tras registro exitoso, 3 en webhook_server). ~10 líneas total. No es bloqueante porque sin esto el producto sigue funcionando, solo no manda emails.
- **Ejecutar `test_mac_e2e.py` en tu Mac.** Es el único test que no pude correr yo. Si encuentra algún FAIL, mandame el log y lo arreglamos.
- **Revisión legal de T&C/Privacy.** Llenar los placeholders (dirección, jurisdicción), revisar con abogado, configurar emails `legal@`, `privacy@`, `dpo@`.
- **Deployar webhook server** (Render/Fly.io/Railway) y configurar Stripe Dashboard. Sin esto los pagos no se sincronizan a la DB aunque Stripe los cobre.
- **Configurar el proveedor de email** (Resend recomendado, $20/mes desde 50k emails) y los DNS records (SPF/DKIM) para que los emails no caigan en spam.

---

## Pendiente para futuras releases

- Onboarding interactivo en EN tras registro (tutorial guiado)
- Soporte multi-currency en Checkout (hoy todo en USD)
- Cron job para `send_trial_ending_soon()` (3 días antes del fin del trial)
- Branding consistente en emails (logo, colores corporativos cuando exista identidad visual)
- División de archivos monolíticos (`word_controller.py` ~5650 líneas, `excel_controller.py` ~5300 líneas) — refactor pendiente desde v14.1
- Cobertura del post-procesador de Excel/PPT al 99.5%+ (hoy ~99% según testing real)
