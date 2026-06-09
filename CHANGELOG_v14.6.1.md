# CHANGELOG — WEP AI v14.6.1 (patch)

**Tipo:** patch release. Sólo un cambio funcional + documentación.
**Fecha:** 2026-05-21
**Base:** v14.6 (todos los demás features y fixes intactos)

## Cambio

### Trial de 14 días removido del plan Personal Pro

**Decisión de producto:** el plan Personal Free (5 docs/mes, sin tarjeta) ya cumple el rol de "prueba gratuita del producto". Un usuario que llega a elegir Pro o Enterprise ya conoce WEP AI y va con intención de pago directo. Mantener un trial adicional encima del Free duplicaba caminos y diluía la propuesta:

- **Antes (v14.6):** "Empezar prueba gratis de 14 días — luego $4.99/mes". Stripe Checkout mostraba "$0.00 today, then $4.99/mo on Jun 4".
- **Ahora (v14.6.1):** "Empezar plan Personal Pro · $4.99/mes". Stripe Checkout cobra inmediatamente $4.99 al primer checkout.

### Archivos modificados

- **`ui/chat_window.py`** — `_open_payment_stripe()` ahora pasa `trial_days=0` en lugar de 14. El comentario explica la decisión de producto.
- **`storage/stripe_client.py`** — docstring de `trial_days` actualizada: el default sigue siendo 0, pero ahora se documenta explícitamente por qué (no es "olvido", es decisión).

### Lo que NO cambió

- **El parámetro `trial_days` sigue existiendo** en `create_checkout_session()`. Si en el futuro se decide reactivar trial para una campaña o segmento específico, basta con setear la env var `STRIPE_TRIAL_DAYS=N` y el cliente Stripe lo respeta automáticamente — sin tocar código.
- **`send_trial_ending_soon()`** en `storage/email.py` se mantiene. Es un helper inerte si no hay trial, pero queda disponible para uso futuro.
- **Test 13.3** sigue validando que la signature de `create_checkout_session` acepta `trial_days`. No verifica un valor específico — verifica que el parámetro funciona.

## Cobertura de tests

13/13 grupos siguen pasando. El cambio no requiere test nuevo porque el comportamiento Stripe con `trial_days=0` es el path por default ya cubierto en Test 12.4.

## Impacto en mockups y materiales

El mockup de la pantalla de registro que se mostró en una conversación previa tenía el botón "Empezar prueba gratis de 14 días — luego $4.99/mes". En v14.6.1 ese botón dice simplemente lo que el código emite: el nombre del plan + precio. El texto exacto del botón depende del template `btn_k[self.sel_plan]` del LANG dict — si querés cambiar el copy del botón mismo (ej. "Suscribirme a Personal Pro · $4.99/mes"), decime y lo actualizo en otra release.

## Sin breaking changes

Todo lo demás funciona igual. Quien actualice de v14.6 a v14.6.1 sólo verá un cambio: el primer cobro de Stripe sucede al checkout en lugar de 14 días después.
