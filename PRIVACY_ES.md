# WEP AI — Política de Privacidad

**Vigencia desde:** 21 de mayo de 2026
**Última actualización:** 21 de mayo de 2026

Esta Política de Privacidad explica qué información recopila WEP AI ("nosotros") sobre vos, cómo la usamos y tus derechos sobre ella. Aplica a tu uso de la aplicación de escritorio WEP AI y servicios relacionados ("el Servicio").

## Resumen en lenguaje claro

- La mayoría de tus datos se quedan en tu equipo. No corremos una base de datos en la nube con tus documentos.
- Las instrucciones que enviás a la IA se envían a **Anthropic** (la empresa detrás de Claude) para procesamiento. No las guardamos en nuestros servidores.
- Si te suscribís a un plan pago, **Stripe** procesa tu pago. Nunca vemos tu número completo de tarjeta.
- No vendemos tu información personal. No mostramos publicidad.
- Podés eliminar tu cuenta en cualquier momento.

## 1. Qué recopilamos

### Información que proporcionás
- **Datos de cuenta**: nombre, apellido, email, contraseña (almacenada como hash bcrypt, nunca en texto plano)
- **Selección de plan**: free, Pro o Enterprise
- **Instrucciones de conversación**: las solicitudes en lenguaje natural que enviás para generar documentos

### Información recopilada automáticamente
- **Métricas de uso**: cantidad de documentos generados por mes (solo para enforcement del plan)
- **Logs de aplicación**: timestamps de generaciones, errores y resultados del detector de bugs — almacenados localmente

### Información de pagos (solo si te suscribís)
- **Últimos 4 dígitos de tu tarjeta**: almacenados localmente para mostrar en tu UI de facturación ("Visa terminada en 1234")
- **ID de cliente de Stripe e ID de suscripción**: almacenados localmente para que la app pueda consultar el estado de tu plan

**No** almacenamos tu número completo de tarjeta, CVV, fecha de expiración ni dirección de facturación. Eso lo maneja exclusivamente Stripe.

## 2. Dónde viven tus datos

| Datos | Dónde viven | Quién accede |
|---|---|---|
| Email + hash de contraseña | SQLite local en `~/.wepai/wepai.db` | Vos |
| Historial de conversación | SQLite local en `~/.wepai/wepai.db` | Vos |
| Documentos generados | Carpeta local `~/Documents/WEP_AI/` | Vos |
| Datos de pago | Sistemas PCI-compliant de Stripe | Stripe + vos |
| Instrucciones (transitorias) | Enviadas a la API de Anthropic en el momento de generación | Anthropic |

**No** tenemos una base de datos en la nube con tus datos. Si desinstalás WEP AI y eliminás `~/.wepai/` y `~/Documents/WEP_AI/`, efectivamente borraste todos tus datos desde nuestro lado.

## 3. Cómo las instrucciones llegan a Anthropic

Cuando enviás una instrucción (ej. "Generá una factura para el cliente ABC por $5,000"), la app WEP AI envía esa instrucción directamente desde tu equipo a la API Claude de Anthropic. Anthropic procesa la instrucción y devuelve el JSON estructurado que se usa para construir tu documento.

El manejo de tu instrucción por parte de Anthropic se rige por la [Política de Privacidad de Anthropic](https://www.anthropic.com/legal/privacy) y los [Términos Comerciales](https://www.anthropic.com/legal/commercial-terms). Según la política de Anthropic, las instrucciones enviadas vía API no se usan para entrenar a Claude.

Si no querés que tus instrucciones salgan de tu equipo, no uses WEP AI.

## 4. Cómo usamos tu información

Usamos la información almacenada localmente para:
- Autenticarte y enforcer los límites de tu plan
- Mostrarte tu historial de conversación y documentos generados
- Detectar y reportar bugs (el validador de Excel corre localmente)
- Para planes pagos: verificar que tu suscripción esté activa

**No** usamos tu información para:
- Entrenar modelos de IA
- Construir perfiles para publicidad
- Vender a terceros
- Compartir con brokers de datos

## 5. Compartir información

Solo compartimos información en estos casos limitados:

- **Anthropic** recibe tus instrucciones en el momento de generación, como se describe en sección 3
- **Stripe** recibe tu información de pago cuando te suscribís
- **Autoridades** ante una orden judicial válida
- **Empresa sucesora** si WEP AI es adquirido o fusionado, tus datos pueden transferirse bajo los mismos términos

No compartimos con ningún otro tercero.

## 6. Tus derechos

Dependiendo de dónde vivas, tenés los siguientes derechos:

### Para todos los usuarios
- **Acceso**: ver los datos que tenemos sobre vos (almacenados localmente, abrí `~/.wepai/wepai.db` con cualquier navegador SQLite)
- **Eliminación**: eliminar tu cuenta desde la app, o manualmente eliminar `~/.wepai/` y `~/Documents/WEP_AI/`
- **Exportación**: copiá tus documentos de `~/Documents/WEP_AI/` y la DB; ambos usan formatos abiertos

### Para usuarios en la Unión Europea (GDPR)
Tenés derecho a:
- Acceder a tus datos personales (Artículo 15)
- Corregir datos inexactos (Artículo 16)
- Borrar tus datos ("derecho al olvido" — Artículo 17)
- Restringir el procesamiento (Artículo 18)
- Portabilidad de datos (Artículo 20)
- Oponerte al procesamiento (Artículo 21)
- Retirar el consentimiento en cualquier momento
- Presentar una queja ante tu autoridad de protección de datos local

Para ejercer derechos GDPR, escribí a **privacy@wepai.app**. Respondemos en 30 días.

La base legal para procesar tus datos es: **cumplimiento contractual** (te registraste para el Servicio) para datos de cuenta; **interés legítimo** (mejorar el producto) para estadísticas anónimas de uso; **consentimiento** para funciones opcionales.

### Para residentes de California (CCPA)
Tenés derecho a:
- Saber qué información personal recopilamos (ver sección 1)
- Saber si la vendemos o compartimos (no lo hacemos)
- Eliminar tu información personal (ver Eliminación arriba)
- Optar por no participar en venta de información personal (no aplica — no vendemos)
- No discriminación por ejercer estos derechos

Para ejercer derechos CCPA, escribí a **privacy@wepai.app**.

## 7. Privacidad de menores

WEP AI no está pensado para uso por menores de 18 años. No recopilamos datos a sabiendas de menores. Si descubrimos que una cuenta pertenece a un menor, la eliminaremos.

## 8. Retención de datos

- **Plan Free**: historial de conversación se mantiene 7 días (luego se borra automáticamente)
- **Plan Pro**: historial se mantiene 90 días
- **Plan Enterprise**: historial se mantiene 1 año
- **Documentos generados**: se mantienen indefinidamente (viven en tu carpeta `~/Documents`; nunca los borramos)
- **Registros de pago**: los mantiene Stripe según su política de retención; nosotros solo guardamos tu ID de suscripción y últimos 4 dígitos de tarjeta

Cuando eliminás tu cuenta, eliminamos datos locales de inmediato. Stripe retiene los registros de pago según requieren las regulaciones financieras (típicamente 7 años).

## 9. Seguridad

- Las contraseñas se hashean con **bcrypt** (estándar de la industria, no reversible)
- La base de datos SQLite local tiene permisos de filesystem limitados a tu cuenta de usuario
- Los datos de pago nunca se almacenan en nuestros sistemas; Stripe maneja todos los datos de tarjeta en infraestructura PCI-compliant
- El servidor de webhooks de Stripe (si está desplegado) verifica las firmas de todos los eventos entrantes para prevenir falsificaciones

A pesar de estas medidas, ningún sistema es 100% seguro. Si nos enteramos de una brecha que afecte tus datos, te notificaremos dentro de las 72 horas como requiere GDPR.

## 10. Transferencias internacionales

Si usás WEP AI desde fuera de Estados Unidos, tus instrucciones se envían a servidores de Anthropic, que pueden estar en EE.UU. Nos basamos en las salvaguardas de Anthropic para transferencias internacionales (ver política de privacidad de Anthropic).

Para usuarios europeos, esta transferencia está cubierta por Cláusulas Contractuales Estándar aprobadas por la Comisión Europea.

## 11. Cambios a esta Política

Podemos actualizar esta Política de Privacidad. Los cambios materiales se anunciarán en la app y por email al menos 14 días antes de su entrada en vigor. El uso continuado luego de los cambios constituye aceptación.

## 12. Contacto

Para consultas relacionadas con privacidad:
- **Email**: privacy@wepai.app
- **Dirección postal**: [WEP AI · dirección · ciudad, estado, código postal, país]
- **Responsable de Protección de Datos** (para asuntos UE/GDPR): dpo@wepai.app
