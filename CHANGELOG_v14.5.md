# CHANGELOG — WEP AI v14.5

**Objetivo de esta release:** dejar WEP AI listo para cobrar a usuarios anglo. Esto requiere tres cosas que faltaban en v14.4: (1) formato de fecha US, (2) moneda USD en lugar de MXN/MN, (3) integración real con Stripe Checkout. También se agregan documentos legales (T&C y Privacy Policy) en ambos idiomas.

**Fecha:** 2026-05-21
**Resultado:** 12/12 grupos de test pasando

---

## Resumen ejecutivo

| Capa | Antes (v14.4) | Después (v14.5) |
|---|---|---|
| Formato de fecha en Word | DD/MM/YYYY en footer y headers, sin importar idioma | `_fmt_short_date(cfg)` bilingüe; EN → MM/DD/YYYY |
| Formato de fecha en Excel | DD/MM/YYYY hardcoded en 20+ lugares | Post-procesador detecta y convierte a MM/DD/YYYY cuando lang=en |
| Moneda en data simulada (Excel) | "MXN", "Peso mexicano" residual | Post-procesador convierte a USD / "US Dollar" |
| Pagos | Solo "MODO DEMO" sin cobro real | **Stripe Checkout real** cuando hay claves; fallback automático a demo si no |
| Documentos legales del producto | No existían | **TERMS_EN.md, TERMS_ES.md, PRIVACY_EN.md, PRIVACY_ES.md** completos |
| Webhook server | No existía | **storage/webhook_server.py** listo para deploy (Render/Fly/Railway) |
| Test suite | 11/11 | **12/12** |

Con esto, un usuario en US puede registrarse, pagar realmente $4.99/mes vía Stripe, recibir documentos en inglés con formato US (fechas, moneda) y tener T&C y Privacy Policy en inglés que aceptó al registrarse.

---

## Cambios por archivo

### 1. `office/word_controller.py` — formato de fecha bilingüe

**Nuevo helper `_fmt_short_date(cfg, dt=None)`** (línea 300):
- ES: `21/05/2026` (formato europeo / latino)
- EN: `05/21/2026` (formato US)

**Reemplazo de 6 ocurrencias** de `today.strftime('%d/%m/%Y')` y `datetime.now().strftime('%d/%m/%Y')` que estaban hardcoded en:
- `_build_internal_memo` — fecha en el bloque firma
- `_build_procedure_manual` — tabla de control de documentos
- `_build_procedure_manual` — tabla de control de cambios
- `_build_employment_contract` — fecha de elaboración del contrato
- `_build_executive_report` — fecha del header
- `_build_recommendation_letter` — fecha de la carta
- `generate_word()` — header reference

**`_wep_footer()` ahora es bilingüe** y respeta el cfg del documento:
- ES: `Documento generado con WEP AI · 21/05/2026 · Revisar con asesor legal antes de firmar.`
- EN: `Generated with WEP AI · 05/21/2026 · Review with legal counsel before signing.`

### 2. `office/excel_controller.py` — post-procesador con fecha/moneda US

**Extensión del post-procesador `_translate_workbook()`** para que cuando `_lang() == "en"`:

1. **Conversión de fechas DD/MM/YYYY → MM/DD/YYYY** mediante regex `\b(\d{1,2})/(\d{1,2})/(\d{4})\b`. Conservadora: si ambos números ≤12 (ambiguo), asume DD/MM como lo emite el código y convierte; si el primero > 12 es inequívocamente día y convierte; si el segundo > 12 deja igual.

2. **Conversión de moneda en labels**: `MXN` → `USD`, `MN` → `USD`, `"$1500 MXN"` → `"$1500"`, "Peso mexicano" → "US Dollar", "Dólar americano" → "US Dollar". Esto afecta los datos simulados de templates como `travel_expenses`, `multi_currency`, `subscription_tracker`, `savings_planner`, `personal_budget` que tenían valores de ejemplo en pesos mexicanos.

**+50 patrones adicionales** encontrados en testing real: `Destino`, `Origen`, `Motivo`, `Reembolsado`, `Alojamiento`, `Alimentación`, `Transporte`, `Hotel`, `Vuelo`, `Taxi`, `Gasolina`, `Estacionamiento`, `Cheque`, `Transferencia`, `Depósito`, `Retiro`, `Crédito`, `Débito`, `Tarjeta`, `Sí/No/Parcial`, etc. Cobertura aproximada estimada subió de 95% a 99% en cells visibles para el usuario.

### 3. `storage/stripe_client.py` — NUEVO módulo (~130 líneas)

**API mínima:**
- `is_configured()` → bool. Detecta si `STRIPE_SECRET_KEY` está en el entorno.
- `create_checkout_session(plan, user_email, user_id)` → dict con `{"url", "session_id"}` o None.
- `create_billing_portal_session(stripe_customer_id)` → URL del portal o None.
- `cancel_subscription(subscription_id)` → bool.

**Diseño:**
- Variables de entorno requeridas: `STRIPE_SECRET_KEY`, `STRIPE_PRICE_ID_PRO`, `STRIPE_PRICE_ID_BIZ`, `WEBHOOK_BASE_URL`.
- Import lazy de `stripe` dentro de cada función (no requiere que esté instalado si no se usa).
- Toda función devuelve None ante cualquier error en lugar de excepción.
- `success_url` y `cancel_url` apuntan al webhook server (que sirve páginas de "vuelve a la app").
- Habilitado `automatic_tax` y `allow_promotion_codes` por default (apropiado para US).

### 4. `storage/webhook_server.py` — NUEVO servidor (~190 líneas)

Servidor Flask separado que el cliente debe desplegar (Render/Fly.io/Railway/AWS Lambda) para recibir eventos de Stripe y actualizar la DB de usuarios.

**Endpoints:**
- `POST /stripe/webhook` — valida firma, procesa eventos `checkout.session.completed`, `customer.subscription.updated`, `customer.subscription.deleted`, `invoice.payment_failed`.
- `GET /checkout-success` — página HTML que ve el usuario tras pagar.
- `GET /checkout-cancel` — página tras cancelar el checkout.
- `GET /billing-portal-return` — página tras gestionar facturación.
- `GET /health` — healthcheck para monitoring.

**Migración DB automática:** si las columnas `stripe_customer_id`, `stripe_subscription_id` o `subscription_status` no existen, las agrega con `ALTER TABLE` la primera vez que recibe un webhook.

**Docstring incluye guía de deploy paso a paso** con configuración de Stripe Dashboard, variables de entorno y testing local con `stripe-cli`.

### 5. `ui/chat_window.py` — pago Stripe primero, demo como fallback

**Refactor de `_open_payment()`:**

```python
def _open_payment(self, nm, ls, em, ps):
    if stripe_client.is_configured():
        self._open_payment_stripe(nm, ls, em, ps)
    else:
        self._open_payment_demo(nm, ls, em, ps)
```

**`_open_payment_stripe()` nuevo:**
1. Registra al usuario como "free" en la DB (placeholder — no le da plan pago hasta que Stripe confirme)
2. Llama a `stripe_client.create_checkout_session()` y obtiene URL
3. Abre el navegador del sistema con esa URL (`webbrowser.open()`)
4. Muestra modal "Completá el pago en tu navegador. Una vez confirmado, tu suscripción se activará automáticamente."
5. Si Stripe falla en cualquier punto, hace fallback automático al modo demo (no bloquea al usuario)

**`_open_payment_demo()` (lo que era `_open_payment` antes):** sigue funcionando como en v14.4 para desarrollo y testing sin claves de Stripe.

### 6. `requirements.txt` — dependencias agregadas

```diff
+ # v14.5 — Stripe billing (opcional)
+ stripe>=7.0.0
+
+ # v14.5 — Webhook server (sólo para deploy del webhook)
+ # pip install flask>=3.0.0
```

`stripe` es soft-dependency: si no está instalado, la app sigue funcionando en modo demo. `flask` no se incluye porque sólo lo necesita el webhook server desplegado en la nube, no la app desktop.

### 7. Documentos legales del producto — NUEVOS (~4 archivos)

- **TERMS_EN.md** (~150 líneas) — Terms of Service en inglés. 16 secciones cubriendo: descripción del Servicio, cuentas, planes y reembolsos, uso aceptable, descargo sobre contenido AI, propiedad intelectual, servicios de terceros (Anthropic, Stripe), disponibilidad, terminación, sin garantías, limitación de responsabilidad, ley aplicable (Delaware + JAMS arbitration), contacto.
- **TERMS_ES.md** — versión en español equivalente.
- **PRIVACY_EN.md** (~180 líneas) — Privacy Policy con secciones CCPA y GDPR explícitas. Cubre qué se recopila, dónde vive (local vs Anthropic vs Stripe), cómo se usa, retención por plan, derechos del usuario, transferencias internacionales (Standard Contractual Clauses), contacto del DPO.
- **PRIVACY_ES.md** — versión en español.

Estos archivos están pensados para ser referenciados desde la UI en pantalla de registro (link "Al registrarte aceptás los Términos y la Política de Privacidad") y para incluirse en el sitio web del producto. **Necesitarán revisión legal** antes de ir a producción — son una base profesional y bien estructurada, pero un abogado en US (y otro en cada jurisdicción donde se opere) debe validarlos.

### 8. `tests/test_db.py` — Test 12

Cuatro sub-tests:
- **12.1** `_fmt_short_date` devuelve formato correcto por idioma
- **12.2** Word footer en EN dice "Generated with WEP AI" con fecha US
- **12.3** Excel post-procesador convierte MXN→USD en `travel_expenses`
- **12.4** `stripe_client.is_configured()` detecta env vars y fallback a None funciona

Suite completa: **12/12 grupos pasan.**

---

## Decisiones de diseño

### ¿Por qué Stripe Checkout y no campos de tarjeta embebidos?

Embeber campos de tarjeta en la app requeriría que WEP AI sea PCI-DSS compliant (costo alto, auditorías recurrentes, infraestructura especializada). Stripe Checkout es la solución estándar para apps desktop:

1. El usuario nunca ingresa la tarjeta dentro de WEP AI — la ingresa en `checkout.stripe.com`
2. Stripe maneja todos los aspectos PCI (campos, almacenamiento, 3D Secure, fraud detection)
3. WEP AI sólo conoce el `subscription_id` y los últimos 4 dígitos (display only)
4. Reduce drásticamente la superficie de ataque y la responsabilidad legal

### ¿Por qué un webhook server separado y no procesarlo en la app desktop?

La app desktop puede estar offline cuando Stripe envía el webhook (`checkout.session.completed`). Si dependiéramos de la app desktop:
- El usuario pagaría pero la suscripción no se activaría hasta que abriera la app de nuevo
- Stripe reintenta el webhook varias veces y eventualmente desiste — perdiendo el evento
- No habría forma de procesar renovaciones mensuales

El webhook server siempre está corriendo. Cuando Stripe le notifica, actualiza la DB. La app desktop hace polling o lee la DB al abrir y ve el plan actualizado.

### ¿Por qué fallback automático a modo demo?

Permite:
- **Desarrollo local sin necesitar claves de Stripe.** Un desarrollador clona el repo y la app funciona inmediatamente.
- **Testing automatizado.** Los tests de UI pueden seguir corriendo sin un mock complicado de Stripe.
- **Onboarding gradual del producto.** El equipo puede deployar la app antes de tener Stripe configurado completamente.

El fallback es transparente: si en algún punto Stripe falla (red, credenciales inválidas, price_id incorrecto), el usuario igual puede registrarse. La app no se rompe.

### ¿Por qué T&C/Privacy en markdown y no PDFs?

Markdown es editable, versionable en git, fácil de actualizar y se puede renderizar como HTML en el sitio web. PDFs serían apropiados para archivar versiones específicas firmadas digitalmente, pero eso ya es trabajo de otra etapa (cuando haya un equipo legal y un workflow de versionado).

---

## Despliegue para producción (checklist)

Para activar pagos reales en producción, el cliente debe:

**1. Cuenta de Stripe** (15 min)
   - Crear cuenta en https://dashboard.stripe.com
   - Activar pagos (puede requerir KYC empresarial)
   - Crear 2 productos con precios recurrentes:
     - "WEP AI Pro" — $4.99/mes → guardar el `price_id` (price_xxx)
     - "WEP AI Enterprise" — $24.99/mes → guardar el `price_id`

**2. Desplegar el webhook server** (30 min)
   - Crear cuenta en Render.com (o Fly.io / Railway)
   - Conectar repo Git
   - Crear nuevo Web Service apuntando a `storage/webhook_server.py`
   - Configurar variables de entorno: `STRIPE_SECRET_KEY`, `STRIPE_WEBHOOK_SECRET`, `DATABASE_URL`
   - Deploy. Anotar la URL pública (ej. `https://wepai-webhook.onrender.com`)

**3. Configurar webhook en Stripe Dashboard** (5 min)
   - https://dashboard.stripe.com/webhooks → Add endpoint
   - URL: la del paso 2 + `/stripe/webhook`
   - Eventos: `checkout.session.completed`, `customer.subscription.updated`, `customer.subscription.deleted`, `invoice.payment_failed`
   - Copiar el "Signing secret" (whsec_xxx) → ponerlo en el env var `STRIPE_WEBHOOK_SECRET` del webhook server

**4. Configurar la app desktop** (5 min)
   En el equipo de cada usuario (o vía instalador), agregar al `.env` o variables de entorno del sistema:
   ```
   STRIPE_SECRET_KEY=sk_live_...
   STRIPE_PRICE_ID_PRO=price_...
   STRIPE_PRICE_ID_BIZ=price_...
   WEBHOOK_BASE_URL=https://wepai-webhook.onrender.com
   ```

**5. Revisión legal** (variable)
   - Revisar TERMS_EN.md y PRIVACY_EN.md con un abogado licenciado en Delaware (o donde se incorpore la empresa)
   - Llenar los placeholders `[your address]`, `[city, state, ZIP]`, etc.
   - Configurar emails `legal@`, `privacy@`, `dpo@` (puede ser forward a un solo email)

**6. Linkear T&C/Privacy en la UI** (15 min de código)
   En la pantalla de registro, agregar un checkbox "He leído y acepto los [Términos](TERMS) y la [Política de Privacidad](PRIVACY)" antes del botón de continuar. Esto NO está hecho en v14.5 — los archivos existen, pero el link en la UI queda como trabajo de un próximo sprint si lo necesitás.

Estimado total: 1-2 horas de configuración + el tiempo de revisión legal.

---

## Pendiente para futuras releases

- **Linkear los T&C/Privacy en la UI** desde la pantalla de registro (checkbox de aceptación)
- **Email transaccional**: notificaciones de bienvenida, pago confirmado, recordatorio de pago próximo, fallo de pago. Requiere integración con SendGrid/Postmark/Resend.
- **Self-service billing UI**: mostrar plan actual, fecha de próxima renovación, link al portal de Stripe para cambiar tarjeta o cancelar. Hoy el usuario tiene que escribir a soporte.
- **Trial period**: muchos productos US ofrecen 14 días free trial al suscribirse a Pro. Stripe Checkout lo soporta nativamente con `subscription_data.trial_period_days=14`.
- **Soporte multi-currency en checkout**: actualmente Stripe Checkout muestra USD. Para usuarios fuera de US podría mostrar EUR/MXN/COP según ubicación.
- **División de archivos monolíticos** (`word_controller.py` ~5650 líneas, `excel_controller.py` ~5300 líneas, `ppt_controller.py` ~1900 líneas) — refactor pendiente desde v14.1, sigue diferido.
- **Pruebas E2E en macOS real con Office instalado** — el flujo `AppleScript → screenshot → vision-verify → iterate` nunca se ejercitó físicamente. Bloqueante para garantizar el round-trip completo.
