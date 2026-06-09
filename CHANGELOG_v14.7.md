# CHANGELOG — WEP AI v14.7

**Objetivo:** integrar el módulo de email transaccional al flujo real de la app (`email.send_*()` que en v14.6 estaba como código disponible pero no conectado), agregar status banner para suscripciones con pago vencido, y un botón de cancelación directa como alternativa al portal completo de Stripe. Con esto el ciclo completo paga → onboarding → uso → fallo de pago → cancelación queda conectado.

**Fecha:** 2026-05-21
**Resultado:** 14/14 grupos de test pasando

---

## Resumen ejecutivo

| Pieza | Antes (v14.6.1) | Después (v14.7) |
|---|---|---|
| Email de bienvenida tras registro | Código existía pero no se llamaba | `_do_register()` y `_open_payment_stripe()` llaman `send_welcome()` automáticamente |
| Email tras pago confirmado | No | webhook `checkout.session.completed` envía `send_payment_confirmed()` |
| Email tras fallo de cobro | No | webhook `invoice.payment_failed` envía `send_payment_failed()` |
| Email tras cancelación | No | webhook `customer.subscription.deleted` envía `send_subscription_cancelled()` |
| Idioma de emails del webhook | Default EN siempre | Stripe metadata transporta `lang`+`email` del usuario para que el webhook use el correcto |
| Status banner past_due/cancelled | No existía | Banner ámbar/gris en chat principal con CTA "Update card" / "Reactivate" |
| Cancelar suscripción | Solo abriendo Stripe portal completo (5+ clicks) | Método directo `cancel_subscription()` con confirmación (1 paso) |
| Test suite | 13/13 | 14/14 |

---

## Cambios por archivo

### 1. `storage/stripe_client.py` — metadata enriquecido para webhook

`create_checkout_session()` ahora acepta `user_lang: str = "en"`. Este idioma se inyecta en el `metadata` de Stripe (tanto a nivel session como subscription) junto con el email del usuario. Cuando Stripe envía eventos webhook, el servidor los recibe con `lang` y `email` ya disponibles, sin tener que consultar la DB en cada evento.

```python
metadata = {"user_id": "42", "plan": "pro",
            "lang": "en", "email": "user@example.com"}
```

### 2. `ui/chat_window.py` — emails + banner + cancel directo

**`_do_register()`** (registro free): tras crear el usuario llama `email.send_welcome(em, nm, lang=self.lang[0])`. Best-effort, envuelto en try/except que loguea pero no bloquea. Si `EMAIL_PROVIDER` no está configurado, `send_welcome` es no-op (dry-run) y todo sigue funcionando.

**`_open_payment_stripe()`** (registro pago): mismo `send_welcome()` antes del checkout. Pasa `user_lang=lbl` a Stripe para que el webhook use el idioma correcto al confirmar.

**`_refresh_billing_banner()`** (nuevo): lee `subscription_status` del usuario. Si es `past_due`, muestra banner ámbar con texto urgente "Tu último pago falló. Actualizá tu tarjeta o tu acceso se suspenderá en 7 días." y botón "Actualizar tarjeta" que abre el portal. Si es `cancelled`, banner gris informativo con CTA "Reactivar". Si es `active` o sin estado, oculto.

Se llama al final de `__init__` del WEPAIApp principal y tras volver del portal de Stripe (cuando refresh manual sea necesario).

**`_cancel_subscription_directly()`** (nuevo): leen `stripe_subscription_id` de la DB, pide confirmación con `messagebox.askyesno`, llama `stripe_client.cancel_subscription(sid)`. Si OK, muestra mensaje de "Cancelación programada, mantenés acceso hasta fin de período". El webhook `customer.subscription.deleted` se encarga de enviar el email + downgrade en la DB.

### 3. `storage/webhook_server.py` — emails por evento

Tres nuevos helpers wrapper no-bloqueantes:

```python
def _safe_send_payment_confirmed(to_email, plan, amount_str, next_billing, lang)
def _safe_send_payment_failed(to_email, plan, lang)
def _safe_send_subscription_cancelled(to_email, end_date, lang)
```

Cada uno: import lazy de `storage.email`, llamada con try/except, loguea pero no propaga la excepción. Crítico porque Stripe espera HTTP 200 de los webhooks o los reintenta — un fallo en el envío de email no debe romper el procesamiento del evento.

**`checkout.session.completed`** → además de actualizar el plan en DB, envía email de confirmación con el monto cobrado. Stripe envía `amount_total` en centavos (e.g. 499 = $4.99), formateado por `_amount_from_invoice()`.

**`invoice.payment_failed`** → además de marcar `past_due`, envía email avisando al usuario. Usa `_get_user_email_and_lang(customer_id)` como fallback porque el evento de invoice no siempre trae metadata (la metadata viaja con la subscription).

**`customer.subscription.deleted`** → además del downgrade, envía email con la fecha de fin del período (calculada de `current_period_end` que viene como Unix timestamp).

### 4. `tests/test_db.py` — Test 14

Cuatro sub-tests:
- **14.1** `create_checkout_session` con `user_lang="en"` produce metadata correcto (mockea Stripe API)
- **14.2** `email.send_welcome` en dry-run devuelve `False` sin lanzar excepción
- **14.3** `stripe_client.cancel_subscription` en dry-run devuelve `False` sin error
- **14.4** `db.get_user_field("subscription_status")` con DB nueva devuelve `None` sin excepción

Suite completa: **14/14 grupos pasan.**

---

## ¿Cuándo se envía cada email?

| Evento | Quién lo envía | Plantilla |
|---|---|---|
| Usuario se registra (Free o pago) | UI (`_do_register` / `_open_payment_stripe`) | `send_welcome` |
| Stripe confirma el primer cobro | webhook (`checkout.session.completed`) | `send_payment_confirmed` |
| Cobro mensual recurrente falla | webhook (`invoice.payment_failed`) | `send_payment_failed` |
| Usuario cancela (UI directo o portal) | webhook (`customer.subscription.deleted`) | `send_subscription_cancelled` |

Si `EMAIL_PROVIDER` no está configurado, todos los `send_*()` son no-op (dry-run, loguean pero no envían). La app sigue funcionando, simplemente no manda emails.

---

## Para activar emails en producción

1. **Elegir proveedor.** Recomiendo Resend (https://resend.com) para arrancar — API más simple, $20/mes desde 50k emails, integración limpia. Alternativas: SendGrid (100 emails/día gratis), Postmark (premium $15/mes).

2. **Verificar dominio.** En el dashboard del proveedor, agregar `wepai.app` (o tu dominio) y configurar los DNS records:
   - SPF: `v=spf1 include:_spf.resend.com ~all`
   - DKIM: registro TXT que te da Resend
   - DMARC (opcional pero recomendado): `v=DMARC1; p=quarantine; rua=mailto:dmarc@wepai.app`
   Esto tarda 24-48h en propagarse. Sin esto los emails caen en spam.

3. **Crear API key** en el dashboard del proveedor.

4. **Configurar env vars** en el webhook server (y opcionalmente en la app desktop si los emails de welcome se envían desde la UI):
   ```
   EMAIL_PROVIDER=resend
   RESEND_API_KEY=re_xxx
   EMAIL_FROM=WEP AI <noreply@wepai.app>
   EMAIL_REPLY_TO=support@wepai.app
   ```

5. **Test.** Enviar un email manualmente desde el servidor:
   ```python
   from storage import email
   email.send_welcome("tu-email@gmail.com", "Tu Nombre", lang="es")
   ```
   Si llega, listo. Si no, revisar logs del provider (Resend/SendGrid tienen un panel con cada envío y rebote).

---

## Lo que sigue pendiente

**Operativo (no es código):**
- Correr `tests/test_mac_e2e.py` en tu Mac
- Configurar provider de email + DNS
- Deploy webhook server (Render/Fly.io/Railway)
- Revisión legal de T&C/Privacy con un abogado
- Beta con 5-10 usuarios

**Producto (próximas releases si se necesitan):**
- Onboarding interactivo bilingüe tras registro (tour guiado)
- Multi-currency en Checkout (hoy todo USD)
- Branding visual de los emails (logo, colores corporativos cuando exista identidad)
- Detección de país en el registro para sugerir el plan según poder adquisitivo
- Polling automático del estado de suscripción cada X minutos (hoy se refresca solo al iniciar la app)
- Refactor de archivos monolíticos (`word_controller.py` ~5650 líneas)
