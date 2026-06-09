"""
WEP AI — Transactional email (v14.6)

Abstracción mínima para emails transaccionales con 3 backends opcionales:

  • Resend  (https://resend.com)    — moderno, simple API, $20/mes desde 50k emails
  • SendGrid (https://sendgrid.com) — más establecido, free tier 100 emails/día
  • Postmark (https://postmarkapp.com) — premium, mejor deliverability, $15/mes

USO:
────
    from storage import email
    email.send_welcome(user_email, user_name, lang="en")
    email.send_payment_confirmed(user_email, plan, amount, lang="es")
    email.send_payment_failed(user_email, plan, lang="en")
    email.send_subscription_cancelled(user_email, end_date, lang="es")

VARIABLES DE ENTORNO:
─────────────────────
    EMAIL_PROVIDER          "resend" | "sendgrid" | "postmark"  (default: ninguno)
    EMAIL_FROM              "WEP AI <noreply@wepai.app>"
    EMAIL_REPLY_TO          "support@wepai.app"  (opcional)

    # Según el proveedor elegido, una de estas:
    RESEND_API_KEY          re_xxx
    SENDGRID_API_KEY        SG.xxx
    POSTMARK_SERVER_TOKEN   xxx

FALLBACK:
─────────
Si EMAIL_PROVIDER no está en el entorno (o no hay API key del proveedor),
los métodos send_* loguean el contenido y devuelven False. La app sigue
funcionando — los emails no se mandan, pero nada se rompe.
"""

import os
import logging
from typing import Optional

log = logging.getLogger("wepai.email")


# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────────────────────────
def is_configured() -> bool:
    """True si hay proveedor configurado con su key correspondiente."""
    provider = (os.getenv("EMAIL_PROVIDER") or "").lower().strip()
    if provider == "resend":   return bool(os.getenv("RESEND_API_KEY"))
    if provider == "sendgrid": return bool(os.getenv("SENDGRID_API_KEY"))
    if provider == "postmark": return bool(os.getenv("POSTMARK_SERVER_TOKEN"))
    return False


def _from() -> str:
    return os.getenv("EMAIL_FROM", "WEP AI <noreply@wepai.app>")


def _reply_to() -> Optional[str]:
    return os.getenv("EMAIL_REPLY_TO") or None


# ─────────────────────────────────────────────────────────────────────────────
# BACKENDS
# ─────────────────────────────────────────────────────────────────────────────
def _send_resend(to: str, subject: str, html: str, text: str) -> bool:
    """Backend: Resend.com (recomendado para nuevos productos)."""
    try:
        import requests
    except ImportError:
        log.error("Para usar Resend, instalá: pip install requests")
        return False
    api_key = os.getenv("RESEND_API_KEY", "")
    payload = {
        "from": _from(),
        "to": [to],
        "subject": subject,
        "html": html,
        "text": text,
    }
    if _reply_to():
        payload["reply_to"] = _reply_to()
    try:
        r = requests.post(
            "https://api.resend.com/emails",
            json=payload,
            headers={"Authorization": f"Bearer {api_key}",
                     "Content-Type": "application/json"},
            timeout=10,
        )
        if r.status_code in (200, 202):
            return True
        log.error(f"Resend error {r.status_code}: {r.text[:200]}")
        return False
    except Exception as e:
        log.error(f"Resend exception: {e}")
        return False


def _send_sendgrid(to: str, subject: str, html: str, text: str) -> bool:
    """Backend: SendGrid (clásico, free tier amplio)."""
    try:
        import requests
    except ImportError:
        log.error("Para usar SendGrid, instalá: pip install requests")
        return False
    api_key = os.getenv("SENDGRID_API_KEY", "")
    # Parsear "Name <email>" en email puro
    sender = _from()
    if "<" in sender:
        sender_name = sender.split("<")[0].strip()
        sender_email = sender.split("<")[1].rstrip(">")
    else:
        sender_name, sender_email = "WEP AI", sender
    payload = {
        "personalizations": [{"to": [{"email": to}]}],
        "from": {"email": sender_email, "name": sender_name},
        "subject": subject,
        "content": [
            {"type": "text/plain", "value": text},
            {"type": "text/html", "value": html},
        ],
    }
    if _reply_to():
        payload["reply_to"] = {"email": _reply_to()}
    try:
        r = requests.post(
            "https://api.sendgrid.com/v3/mail/send",
            json=payload,
            headers={"Authorization": f"Bearer {api_key}",
                     "Content-Type": "application/json"},
            timeout=10,
        )
        if r.status_code in (200, 202):
            return True
        log.error(f"SendGrid error {r.status_code}: {r.text[:200]}")
        return False
    except Exception as e:
        log.error(f"SendGrid exception: {e}")
        return False


def _send_postmark(to: str, subject: str, html: str, text: str) -> bool:
    """Backend: Postmark (premium, deliverability superior)."""
    try:
        import requests
    except ImportError:
        log.error("Para usar Postmark, instalá: pip install requests")
        return False
    token = os.getenv("POSTMARK_SERVER_TOKEN", "")
    payload = {
        "From": _from(),
        "To": to,
        "Subject": subject,
        "HtmlBody": html,
        "TextBody": text,
        "MessageStream": "outbound",
    }
    if _reply_to():
        payload["ReplyTo"] = _reply_to()
    try:
        r = requests.post(
            "https://api.postmarkapp.com/email",
            json=payload,
            headers={"X-Postmark-Server-Token": token,
                     "Content-Type": "application/json",
                     "Accept": "application/json"},
            timeout=10,
        )
        if r.status_code == 200:
            return True
        log.error(f"Postmark error {r.status_code}: {r.text[:200]}")
        return False
    except Exception as e:
        log.error(f"Postmark exception: {e}")
        return False


def _send(to: str, subject: str, html: str, text: str) -> bool:
    """Despacha al backend configurado. Si no hay backend, sólo loguea."""
    provider = (os.getenv("EMAIL_PROVIDER") or "").lower().strip()
    if not is_configured():
        log.info(f"[EMAIL DRY-RUN] to={to} | subj={subject}")
        log.debug(f"[EMAIL DRY-RUN] text:\n{text}")
        return False
    if provider == "resend":   return _send_resend(to, subject, html, text)
    if provider == "sendgrid": return _send_sendgrid(to, subject, html, text)
    if provider == "postmark": return _send_postmark(to, subject, html, text)
    log.warning(f"EMAIL_PROVIDER='{provider}' no reconocido")
    return False


# ─────────────────────────────────────────────────────────────────────────────
# PLANTILLAS — bilingüe (ES/EN)
# ─────────────────────────────────────────────────────────────────────────────
def _wrap_html(body_html: str, lang: str = "es") -> str:
    """Plantilla HTML estándar con header/footer mínimos."""
    footer = ("Si tenés dudas, escribinos a support@wepai.app"
              if lang == "es" else
              "Questions? Email us at support@wepai.app")
    return f"""<!DOCTYPE html><html><head><meta charset="utf-8"></head>
<body style="font-family: Arial, sans-serif; color: #333; max-width: 560px; margin: 0 auto; padding: 20px;">
  <div style="border-bottom: 2px solid #185FA5; padding-bottom: 10px;">
    <h2 style="color: #185FA5; margin: 0;">WEP AI</h2>
  </div>
  <div style="padding: 20px 0;">
    {body_html}
  </div>
  <div style="border-top: 1px solid #ccc; padding-top: 10px; font-size: 11px; color: #888;">
    {footer}
  </div>
</body></html>"""


def send_welcome(to: str, name: str, lang: str = "es") -> bool:
    """Email de bienvenida tras registro."""
    if lang == "en":
        subject = "Welcome to WEP AI"
        text = (f"Hi {name},\n\nWelcome to WEP AI! Your account is ready.\n\n"
                "You can now generate professional Word, Excel and PowerPoint "
                "documents from natural-language prompts.\n\n"
                "Questions? Reply to this email.\n\n— The WEP AI team")
        html = _wrap_html(
            f"<p>Hi <b>{name}</b>,</p>"
            "<p>Welcome to WEP AI! Your account is ready.</p>"
            "<p>You can now generate professional Word, Excel and PowerPoint "
            "documents from natural-language prompts.</p>"
            "<p>— The WEP AI team</p>",
            lang="en")
    else:
        subject = "Bienvenido a WEP AI"
        text = (f"Hola {name},\n\nBienvenido a WEP AI. Tu cuenta está lista.\n\n"
                "Ya podés generar documentos profesionales de Word, Excel y "
                "PowerPoint desde lenguaje natural.\n\n"
                "Cualquier consulta, respondé a este email.\n\n— El equipo de WEP AI")
        html = _wrap_html(
            f"<p>Hola <b>{name}</b>,</p>"
            "<p>Bienvenido a WEP AI. Tu cuenta está lista.</p>"
            "<p>Ya podés generar documentos profesionales de Word, Excel y "
            "PowerPoint desde lenguaje natural.</p>"
            "<p>— El equipo de WEP AI</p>",
            lang="es")
    return _send(to, subject, html, text)


def send_payment_confirmed(to: str, plan: str, amount: str,
                            next_billing_date: str = "", lang: str = "es") -> bool:
    """Email tras pago exitoso de Pro/Enterprise."""
    plan_display = {"pro": "Personal Pro", "biz": "Enterprise",
                    "enterprise": "Enterprise"}.get(plan.lower(), plan)
    if lang == "en":
        subject = f"Payment received — WEP AI {plan_display}"
        next_line = (f" Your next billing date is {next_billing_date}."
                     if next_billing_date else "")
        text = (f"Your payment of {amount} for {plan_display} has been received.\n\n"
                f"Your subscription is now active.{next_line}\n\n"
                "You can manage your subscription anytime from the app.\n\n— WEP AI")
        html = _wrap_html(
            f"<p>Your payment of <b>{amount}</b> for <b>{plan_display}</b> has been received.</p>"
            f"<p>Your subscription is now active.{next_line}</p>"
            "<p>You can manage your subscription anytime from the app.</p>",
            lang="en")
    else:
        subject = f"Pago recibido — WEP AI {plan_display}"
        next_line = (f" Tu próximo cobro será el {next_billing_date}."
                     if next_billing_date else "")
        text = (f"Recibimos tu pago de {amount} por {plan_display}.\n\n"
                f"Tu suscripción está activa.{next_line}\n\n"
                "Podés gestionar tu suscripción desde la app cuando quieras.\n\n— WEP AI")
        html = _wrap_html(
            f"<p>Recibimos tu pago de <b>{amount}</b> por <b>{plan_display}</b>.</p>"
            f"<p>Tu suscripción está activa.{next_line}</p>"
            "<p>Podés gestionar tu suscripción desde la app cuando quieras.</p>",
            lang="es")
    return _send(to, subject, html, text)


def send_payment_failed(to: str, plan: str, lang: str = "es") -> bool:
    """Email cuando un cobro recurrente falla."""
    plan_display = {"pro": "Personal Pro", "biz": "Enterprise"}.get(plan.lower(), plan)
    if lang == "en":
        subject = "Payment failed — please update your card"
        text = ("We tried to charge your card for your WEP AI subscription "
                f"({plan_display}) but the payment was declined.\n\n"
                "Please update your card from your billing portal — the app "
                "has a 'Manage billing' button. Otherwise your access will be "
                "suspended in 7 days.\n\n— WEP AI")
        html = _wrap_html(
            "<p>We tried to charge your card for your WEP AI subscription "
            f"(<b>{plan_display}</b>) but the payment was declined.</p>"
            "<p>Please update your card from your billing portal — the app "
            "has a 'Manage billing' button. Otherwise your access will be "
            "suspended in 7 days.</p>",
            lang="en")
    else:
        subject = "Fallo de pago — actualizá tu tarjeta"
        text = ("Intentamos cobrar tu suscripción de WEP AI "
                f"({plan_display}) pero el pago fue rechazado.\n\n"
                "Por favor actualizá tu tarjeta desde el portal de facturación "
                "— la app tiene un botón 'Gestionar facturación'. De lo "
                "contrario, tu acceso será suspendido en 7 días.\n\n— WEP AI")
        html = _wrap_html(
            "<p>Intentamos cobrar tu suscripción de WEP AI "
            f"(<b>{plan_display}</b>) pero el pago fue rechazado.</p>"
            "<p>Por favor actualizá tu tarjeta desde el portal de facturación "
            "— la app tiene un botón 'Gestionar facturación'. De lo "
            "contrario, tu acceso será suspendido en 7 días.</p>",
            lang="es")
    return _send(to, subject, html, text)


def send_subscription_cancelled(to: str, end_date: str,
                                  lang: str = "es") -> bool:
    """Email tras cancelar suscripción."""
    if lang == "en":
        subject = "Your WEP AI subscription was cancelled"
        text = (f"Your WEP AI subscription has been cancelled. You'll continue "
                f"to have access until {end_date}, after which your account will "
                "downgrade to the Free plan automatically.\n\n"
                "You can re-subscribe anytime from the app.\n\n— WEP AI")
        html = _wrap_html(
            f"<p>Your WEP AI subscription has been cancelled. You'll continue "
            f"to have access until <b>{end_date}</b>, after which your account "
            "will downgrade to the Free plan automatically.</p>"
            "<p>You can re-subscribe anytime from the app.</p>",
            lang="en")
    else:
        subject = "Tu suscripción de WEP AI fue cancelada"
        text = (f"Tu suscripción de WEP AI fue cancelada. Mantenés acceso hasta "
                f"el {end_date}, fecha en la que tu cuenta volverá automáticamente "
                "al plan Free.\n\n"
                "Podés volver a suscribirte cuando quieras desde la app.\n\n— WEP AI")
        html = _wrap_html(
            f"<p>Tu suscripción de WEP AI fue cancelada. Mantenés acceso hasta "
            f"el <b>{end_date}</b>, fecha en la que tu cuenta volverá automáticamente "
            "al plan Free.</p>"
            "<p>Podés volver a suscribirte cuando quieras desde la app.</p>",
            lang="es")
    return _send(to, subject, html, text)


def send_trial_ending_soon(to: str, days_left: int, lang: str = "es") -> bool:
    """Email recordando que el trial termina en N días."""
    if lang == "en":
        subject = f"Your WEP AI trial ends in {days_left} days"
        text = (f"Your free trial of WEP AI ends in {days_left} days.\n\n"
                "If you don't cancel, you'll be charged automatically. "
                "Otherwise, you can cancel from the app's billing portal.\n\n"
                "— WEP AI")
        html = _wrap_html(
            f"<p>Your free trial of WEP AI ends in <b>{days_left} days</b>.</p>"
            "<p>If you don't cancel, you'll be charged automatically. "
            "Otherwise, you can cancel from the app's billing portal.</p>",
            lang="en")
    else:
        subject = f"Tu prueba gratuita de WEP AI termina en {days_left} días"
        text = (f"Tu prueba gratuita de WEP AI termina en {days_left} días.\n\n"
                "Si no cancelás, se realizará el cobro automáticamente. "
                "Si querés cancelar, hacelo desde el portal de facturación de "
                "la app.\n\n— WEP AI")
        html = _wrap_html(
            f"<p>Tu prueba gratuita de WEP AI termina en <b>{days_left} días</b>.</p>"
            "<p>Si no cancelás, se realizará el cobro automáticamente. "
            "Si querés cancelar, hacelo desde el portal de facturación de la app.</p>",
            lang="es")
    return _send(to, subject, html, text)
