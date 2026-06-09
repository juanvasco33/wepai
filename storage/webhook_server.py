"""
WEP AI — Stripe Webhook Server (v14.5)

ESTE ARCHIVO NO CORRE EN LA APP DESKTOP. Es un servidor separado que
debe desplegarse en la nube para recibir eventos de Stripe y actualizar
la base de datos de usuarios cuando un pago se completa o se cancela.

DESPLIEGUE RECOMENDADO:
───────────────────────
• Render.com (gratis tier disponible): conecta repo, agrega env vars,
  deploy automático en cada push
• Fly.io: similar, también gratis tier
• Railway.app: similar
• AWS Lambda + API Gateway: para producción seria

REQUISITOS:
───────────
pip install flask stripe

Variables de entorno requeridas en el servidor:
  STRIPE_SECRET_KEY        sk_live_... o sk_test_...
  STRIPE_WEBHOOK_SECRET    whsec_... (lo da Stripe al crear el endpoint)
  DATABASE_URL             URL a la misma DB que usa la app desktop
                           (postgres recomendado para multi-cliente,
                            o servir el SQLite por la red)

CONFIGURACIÓN EN STRIPE DASHBOARD:
──────────────────────────────────
1. Ir a https://dashboard.stripe.com/webhooks
2. Add endpoint → URL de tu webhook (ej. https://wepai-webhook.onrender.com/stripe/webhook)
3. Suscribirse a estos eventos:
     ✓ checkout.session.completed       (cuando el usuario completa el pago)
     ✓ customer.subscription.updated    (cambios de plan)
     ✓ customer.subscription.deleted    (cancelaciones)
     ✓ invoice.payment_succeeded        (renovaciones exitosas)
     ✓ invoice.payment_failed           (fallos de cobro)
4. Copiar el "Signing secret" (whsec_...) y ponerlo en STRIPE_WEBHOOK_SECRET

EJEMPLO RUN LOCAL (para testing con stripe CLI):
─────────────────────────────────────────────────
  pip install flask stripe
  export STRIPE_SECRET_KEY=sk_test_...
  export STRIPE_WEBHOOK_SECRET=$(stripe listen --print-secret)
  export DATABASE_URL=sqlite:///path/to/wepai.db
  stripe listen --forward-to localhost:5000/stripe/webhook &
  python webhook_server.py
"""

import os
import logging
import sqlite3
from datetime import datetime, timezone

try:
    from flask import Flask, request, jsonify
    import stripe
except ImportError:
    print("Instalá las dependencias: pip install flask stripe")
    raise SystemExit(1)

app = Flask(__name__)
log = logging.getLogger("wepai.webhook")
logging.basicConfig(level=logging.INFO)

stripe.api_key       = os.getenv("STRIPE_SECRET_KEY", "")
WEBHOOK_SECRET       = os.getenv("STRIPE_WEBHOOK_SECRET", "")
DATABASE_URL         = os.getenv("DATABASE_URL", "sqlite:///wepai.db")


# ─────────────────────────────────────────────────────────────────────────────
# DB HELPERS — actualiza la tabla `users` con stripe_customer_id, subscription_id, plan
# ─────────────────────────────────────────────────────────────────────────────
def _db_path():
    """Resuelve la ruta del SQLite desde la URL."""
    if DATABASE_URL.startswith("sqlite:///"):
        return DATABASE_URL.replace("sqlite:///", "", 1)
    raise ValueError("Sólo se soporta sqlite:/// en este template. "
                     "Para Postgres, adaptar usando psycopg2.")


def _update_user_subscription(user_id: int, plan: str,
                              stripe_customer_id: str = None,
                              subscription_id: str = None,
                              status: str = "active"):
    """Actualiza la fila del usuario con los datos de la suscripción."""
    db = sqlite3.connect(_db_path())
    try:
        cur = db.cursor()
        # Aseguramos que existan las columnas (idempotente)
        cur.execute("PRAGMA table_info(users)")
        cols = {row[1] for row in cur.fetchall()}
        if "stripe_customer_id" not in cols:
            cur.execute("ALTER TABLE users ADD COLUMN stripe_customer_id TEXT")
        if "stripe_subscription_id" not in cols:
            cur.execute("ALTER TABLE users ADD COLUMN stripe_subscription_id TEXT")
        if "subscription_status" not in cols:
            cur.execute("ALTER TABLE users ADD COLUMN subscription_status TEXT DEFAULT 'none'")
        # Actualización
        cur.execute("""
            UPDATE users
               SET plan = ?, subscription_status = ?,
                   stripe_customer_id = COALESCE(?, stripe_customer_id),
                   stripe_subscription_id = COALESCE(?, stripe_subscription_id)
             WHERE id = ?
        """, (plan, status, stripe_customer_id, subscription_id, user_id))
        db.commit()
        log.info(f"User {user_id} updated → plan={plan}, status={status}")
    finally:
        db.close()


def _downgrade_user_to_free(user_id: int):
    """Cuando una suscripción se cancela, el usuario vuelve a 'free'."""
    _update_user_subscription(user_id, plan="free", status="cancelled")


# ─────────────────────────────────────────────────────────────────────────────
# EMAIL HELPERS (v14.7) — wrappers no-bloqueantes
# ─────────────────────────────────────────────────────────────────────────────
# Importación lazy para que el webhook server no requiera el módulo si no se
# usa. Si email.send_* falla por cualquier razón (sin provider, red caída,
# etc.) lo logueamos pero no rompemos el procesamiento del webhook — Stripe
# espera HTTP 200 o reintenta.

def _safe_send_payment_confirmed(to_email: str, plan: str, amount_str: str,
                                  next_billing: str, lang: str):
    try:
        from storage import email
        email.send_payment_confirmed(to_email, plan, amount_str,
                                      next_billing, lang=lang)
    except Exception as e:
        log.warning(f"send_payment_confirmed falló (no-bloqueante): {e}")


def _safe_send_payment_failed(to_email: str, plan: str, lang: str):
    try:
        from storage import email
        email.send_payment_failed(to_email, plan, lang=lang)
    except Exception as e:
        log.warning(f"send_payment_failed falló (no-bloqueante): {e}")


def _safe_send_subscription_cancelled(to_email: str, end_date: str, lang: str):
    try:
        from storage import email
        email.send_subscription_cancelled(to_email, end_date, lang=lang)
    except Exception as e:
        log.warning(f"send_subscription_cancelled falló (no-bloqueante): {e}")


def _get_user_email_and_lang(customer_id: str) -> tuple:
    """Fallback: si el evento no trae email/lang en metadata, los leemos de la DB.

    v14.8 fixes:
    - N3: usa context manager para que la conexión se cierre siempre
          (antes había un return entre connect y close cuando faltaba la columna 'email').
    - N4: ahora SÍ lee el lang real si existe una columna 'lang' en users.
          Antes hardcodeaba 'en' incluso para usuarios hispanohablantes.
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
            # Si existe una columna de idioma (lang/language/locale), la incluimos.
            lang_col = next((c for c in ("lang", "language", "locale") if c in cols), None)
            select = f"SELECT email{', ' + lang_col if lang_col else ''} FROM users WHERE stripe_customer_id = ?"
            cur.execute(select, (customer_id,))
            row = cur.fetchone()
            if not row:
                return (None, "en")
            email = row[0]
            # Normalizamos a 'es' / 'en'; cualquier otro valor cae a 'en'.
            raw_lang = (row[1] if lang_col and len(row) > 1 else None) or "en"
            lang = "es" if str(raw_lang).lower().startswith("es") else "en"
            return (email, lang)
    except Exception as e:
        log.warning(f"_get_user_email_and_lang: {e}")
        return (None, "en")


def _amount_from_invoice(obj: dict) -> str:
    """Stripe envía amount_total en cents (e.g. 499 = $4.99). Lo formateamos."""
    amount = obj.get("amount_total") or obj.get("amount_paid") or obj.get("amount_due") or 0
    currency = (obj.get("currency") or "usd").upper()
    if amount and isinstance(amount, (int, float)):
        return f"${amount/100:.2f} {currency}"
    return ""


def _next_billing_date_from_obj(obj: dict) -> str:
    """Stripe envía current_period_end como timestamp Unix."""
    ts = obj.get("current_period_end") or 0
    if ts:
        return datetime.fromtimestamp(ts, tz=timezone.utc).strftime("%Y-%m-%d")
    return ""


# ─────────────────────────────────────────────────────────────────────────────
# WEBHOOK ENDPOINT
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/stripe/webhook", methods=["POST"])
def stripe_webhook():
    payload = request.data
    sig = request.headers.get("Stripe-Signature", "")

    # Verificación de firma (CRÍTICO — sin esto cualquiera puede hacerse pasar por Stripe)
    try:
        event = stripe.Webhook.construct_event(payload, sig, WEBHOOK_SECRET)
    except ValueError:
        log.error("Webhook: payload inválido")
        return jsonify({"error": "invalid payload"}), 400
    except stripe.error.SignatureVerificationError:
        log.error("Webhook: firma inválida")
        return jsonify({"error": "invalid signature"}), 400

    event_type = event["type"]
    obj = event["data"]["object"]
    log.info(f"Evento Stripe recibido: {event_type}")

    # ── checkout.session.completed → activar la suscripción ─────────────────
    if event_type == "checkout.session.completed":
        metadata = obj.get("metadata", {})
        user_id = int(metadata.get("user_id", 0))
        plan = metadata.get("plan", "pro")
        user_lang = metadata.get("lang", "en")
        user_email = metadata.get("email", "") or obj.get("customer_email", "")
        customer_id = obj.get("customer", "")
        subscription_id = obj.get("subscription", "")
        if user_id:
            _update_user_subscription(user_id, plan,
                                      stripe_customer_id=customer_id,
                                      subscription_id=subscription_id,
                                      status="active")
            # v14.7 — Email de pago confirmado
            if user_email:
                amount_str = _amount_from_invoice(obj)
                _safe_send_payment_confirmed(
                    user_email, plan, amount_str, "", user_lang)

    # ── customer.subscription.updated → puede ser upgrade, downgrade, etc ───
    elif event_type == "customer.subscription.updated":
        metadata = obj.get("metadata", {})
        user_id = int(metadata.get("user_id", 0))
        plan = metadata.get("plan", "pro")
        status = obj.get("status", "active")  # active / past_due / cancelled
        if user_id:
            _update_user_subscription(user_id, plan, status=status)

    # ── customer.subscription.deleted → downgrade a free ────────────────────
    elif event_type == "customer.subscription.deleted":
        metadata = obj.get("metadata", {})
        user_id = int(metadata.get("user_id", 0))
        user_lang = metadata.get("lang", "en")
        user_email = metadata.get("email", "")
        if user_id:
            _downgrade_user_to_free(user_id)
            # v14.7 — Email de cancelación confirmada
            end_date = _next_billing_date_from_obj(obj)
            if not user_email:
                customer_id = obj.get("customer", "")
                user_email, _ = _get_user_email_and_lang(customer_id)
            if user_email:
                _safe_send_subscription_cancelled(user_email, end_date, user_lang)

    # ── invoice.payment_failed → marca past_due (UI mostrará banner) ────────
    elif event_type == "invoice.payment_failed":
        customer_id = obj.get("customer", "")
        # Buscar user por stripe_customer_id
        db = sqlite3.connect(_db_path())
        try:
            cur = db.cursor()
            cur.execute("SELECT id, plan FROM users WHERE stripe_customer_id = ?",
                        (customer_id,))
            row = cur.fetchone()
            if row:
                _update_user_subscription(row[0], row[1], status="past_due")
                # v14.7 — Email avisando del fallo de cobro
                user_email, user_lang = _get_user_email_and_lang(customer_id)
                if user_email:
                    _safe_send_payment_failed(user_email, row[1], user_lang)
        finally:
            db.close()

    return jsonify({"received": True}), 200


# ─────────────────────────────────────────────────────────────────────────────
# Páginas de éxito / cancelación que ve el usuario al volver del checkout
# ─────────────────────────────────────────────────────────────────────────────
@app.route("/checkout-success")
def checkout_success():
    return """
    <html><head><title>Payment successful</title></head>
    <body style="font-family:Arial; text-align:center; padding:60px;">
      <h1>✓ Payment received</h1>
      <p>Your WEP AI subscription is now active.</p>
      <p>You can close this tab and return to the WEP AI app.</p>
    </body></html>
    """


@app.route("/checkout-cancel")
def checkout_cancel():
    return """
    <html><head><title>Payment cancelled</title></head>
    <body style="font-family:Arial; text-align:center; padding:60px;">
      <h1>Payment cancelled</h1>
      <p>No charge was made. You can close this tab and try again from the app.</p>
    </body></html>
    """


@app.route("/billing-portal-return")
def portal_return():
    return """
    <html><head><title>Done</title></head>
    <body style="font-family:Arial; text-align:center; padding:60px;">
      <h1>Done</h1>
      <p>Your billing settings have been updated. You can close this tab.</p>
    </body></html>
    """


@app.route("/health")
def health():
    return jsonify({"status": "ok", "time": datetime.now(timezone.utc).isoformat()})


if __name__ == "__main__":
    port = int(os.getenv("PORT", "5000"))
    log.info(f"WEP AI webhook server running on :{port}")
    app.run(host="0.0.0.0", port=port)
