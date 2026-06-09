"""
WEP AI — Stripe Checkout integration (v14.5)

Maneja la creación de Stripe Checkout Sessions para suscripciones (Pro,
Enterprise) y la verificación del estado de suscripción de un usuario.

ARQUITECTURA:
─────────────
WEP AI es una app de escritorio (macOS). Para procesar pagos de forma segura
desde una app desktop, NO se puede embeber un campo de tarjeta dentro de la
app (eso requeriría PCI compliance). El patrón estándar es:

  1. El usuario hace clic en "Suscribirse a Pro"
  2. La app llama a este módulo → crea una Checkout Session en Stripe
  3. Se abre el navegador del usuario con la URL de Checkout (hosted por Stripe)
  4. El usuario completa el pago en la web de Stripe
  5. Stripe envía un webhook a NUESTRO SERVIDOR (ver webhook_server.py)
  6. El servidor actualiza la DB del usuario: plan="pro", subscription_id=...
  7. La app desktop hace polling cada minuto: ¿mi plan cambió? → actualiza UI

REQUISITOS PARA QUE FUNCIONE:
─────────────────────────────
1. Cuenta de Stripe (https://dashboard.stripe.com)
2. Crear 2 productos con precios recurrentes mensuales:
     - WEP AI Pro       → $4.99/mes  → guarda el price_id (price_xxx)
     - WEP AI Enterprise→ $24.99/mes → guarda el price_id (price_xxx)
3. Variables de entorno:
     - STRIPE_SECRET_KEY        (sk_live_... o sk_test_...)
     - STRIPE_WEBHOOK_SECRET    (whsec_...)
     - STRIPE_PRICE_ID_PRO      (price_...)
     - STRIPE_PRICE_ID_BIZ      (price_...)
     - WEBHOOK_BASE_URL         (URL pública del webhook server)
4. Un webhook server desplegado en la nube (Render, Fly.io, Heroku, AWS Lambda)
   con el código de `webhook_server.py` y la URL configurada en Stripe Dashboard
   como endpoint para los eventos:
     - checkout.session.completed
     - customer.subscription.updated
     - customer.subscription.deleted

FALLBACK:
─────────
Si NO hay STRIPE_SECRET_KEY en el entorno, las funciones devuelven None
y la UI debe mostrar el formulario de "MODO DEMO · Sin cobro real" como
hasta v14.4. Esto permite que la app siga funcionando en desarrollo sin
necesidad de claves de Stripe.
"""

import os
import logging

log = logging.getLogger("wepai.stripe")


def is_configured() -> bool:
    """True si las claves de Stripe están presentes en el entorno."""
    return bool(os.getenv("STRIPE_SECRET_KEY"))


def _get_price_id(plan: str) -> str:
    """Mapea 'pro'/'biz' al price_id de Stripe configurado en el entorno."""
    p = (plan or "").lower().strip()
    if p == "pro":
        return os.getenv("STRIPE_PRICE_ID_PRO", "")
    if p in ("biz", "enterprise"):
        return os.getenv("STRIPE_PRICE_ID_BIZ", "")
    return ""


def create_checkout_session(plan: str, user_email: str, user_id: int,
                              trial_days: int = 0,
                              user_lang: str = "en") -> dict:
    """
    Crea una Stripe Checkout Session para suscripción.

    Args:
        plan: "pro" o "biz" / "enterprise"
        user_email: email del usuario (Stripe lo prellena en el checkout)
        user_id: ID del usuario en nuestra DB (lo pasamos en metadata para que
                 el webhook sepa qué fila actualizar)
        trial_days: días de trial gratuito antes del primer cobro.
                    **Default: 0 (sin trial).** Decisión de producto v14.6.1:
                    el plan Personal Free ya permite probar el software (5
                    docs/mes), por lo que un usuario que elige Pro/Enterprise
                    va con intención de pago directo. Si se desea reactivar
                    trial, setear la env var `STRIPE_TRIAL_DAYS=14` (o N) —
                    se respeta sin tocar código.
        user_lang: v14.7 — idioma del usuario ("es" o "en"). Se inyecta en
                    el metadata de Stripe para que el webhook server lo lea
                    al enviar emails transaccionales en el idioma correcto.

    Returns:
        dict con:
          {"url": "https://checkout.stripe.com/...", "session_id": "cs_..."}
        o None si Stripe no está configurado o hay un error.
    """
    if not is_configured():
        log.warning("Stripe no configurado — usando modo demo")
        return None

    price_id = _get_price_id(plan)
    if not price_id:
        log.error(f"No hay price_id configurado para plan '{plan}'")
        return None

    try:
        # Import dentro de la función para no requerir `stripe` si no se usa
        import stripe
    except ImportError:
        log.error("paquete 'stripe' no instalado. Instalá con: pip install stripe>=7.0.0")
        return None

    stripe.api_key = os.getenv("STRIPE_SECRET_KEY")

    try:
        webhook_base = os.getenv("WEBHOOK_BASE_URL", "https://wepai.app").rstrip("/")
        # v14.6 — subscription_data con trial_period_days si aplica.
        # El trial_days desde env var STRIPE_TRIAL_DAYS tiene precedencia si
        # está definido (permite cambiar el trial sin redeployar la app).
        env_trial = os.getenv("STRIPE_TRIAL_DAYS", "").strip()
        if env_trial.isdigit():
            trial_days = int(env_trial)

        subscription_data = {"metadata": {"user_id": str(user_id),
                                          "plan": plan,
                                          "lang": user_lang,
                                          "email": user_email}}
        if trial_days > 0:
            subscription_data["trial_period_days"] = trial_days

        session = stripe.checkout.Session.create(
            mode="subscription",
            payment_method_types=["card"],
            line_items=[{"price": price_id, "quantity": 1}],
            customer_email=user_email,
            metadata={"user_id": str(user_id), "plan": plan,
                      "lang": user_lang, "email": user_email},
            subscription_data=subscription_data,
            success_url=f"{webhook_base}/checkout-success?session_id={{CHECKOUT_SESSION_ID}}",
            cancel_url=f"{webhook_base}/checkout-cancel",
            allow_promotion_codes=True,
            billing_address_collection="auto",
            automatic_tax={"enabled": True},  # Stripe Tax para US/UE
        )
        return {"url": session.url, "session_id": session.id}
    except Exception as e:
        log.error(f"Error creando Stripe Checkout Session: {e}")
        return None


def create_billing_portal_session(stripe_customer_id: str) -> str:
    """Crea una sesión del portal de facturación de Stripe.
    Devuelve la URL del portal (donde el usuario puede cancelar, actualizar
    tarjeta, ver facturas, etc.) o None si falla.
    """
    if not is_configured():
        return None
    try:
        import stripe
    except ImportError:
        return None
    stripe.api_key = os.getenv("STRIPE_SECRET_KEY")
    try:
        webhook_base = os.getenv("WEBHOOK_BASE_URL", "https://wepai.app").rstrip("/")
        portal = stripe.billing_portal.Session.create(
            customer=stripe_customer_id,
            return_url=f"{webhook_base}/billing-portal-return",
        )
        return portal.url
    except Exception as e:
        log.error(f"Error creando billing portal: {e}")
        return None


def cancel_subscription(subscription_id: str) -> bool:
    """Cancela una suscripción al final del período actual.
    Devuelve True si la cancelación se programó OK."""
    if not is_configured():
        return False
    try:
        import stripe
    except ImportError:
        return False
    stripe.api_key = os.getenv("STRIPE_SECRET_KEY")
    try:
        stripe.Subscription.modify(subscription_id, cancel_at_period_end=True)
        return True
    except Exception as e:
        log.error(f"Error cancelando suscripción: {e}")
        return False
