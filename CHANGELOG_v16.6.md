# CHANGELOG v16.6 — Versión web profesional (auth + UX SaaS)

Fecha: 30 de mayo de 2026
M�dulos: web/server.py (reescrito), web/static/index.html (reescrito)
Reutiliza storage/db.py tal cual (bcrypt, planes, conversaciones).

## Autenticación web (login / contraseña)
- Registro, inicio y cierre de sesión reales sobre la web. Sesión mediante
  cookie firmada con HMAC (HttpOnly, SameSite=Lax, Secure salvo desarrollo),
  sin dependencias nuevas. Secreto desde WEPAI_SECRET_KEY.
- Endpoints: POST /api/register, /api/login, /api/logout; GET /api/me.
- /api/chat ahora requiere sesión (devuelve 401 sin autenticar).
- Validaciones: correo válido, contraseña ≥ 8, correo único (409).

## Historial por usuario
- Conversaciones y mensajes persistidos en SQLite y expuestos en la web.
- GET /api/conversations y GET /api/conversations/{id} (con verificación de
  propiedad). Título de la conversación generado por el motor.

## Límites de plan
- Al generar un documento se consume 1 crédito de forma atómica
  (check_and_consume_doc). Si el plan Gratis llega al tope (5/mes), la API lo
  informa y la UI ofrece mejorar a Pro. Contador visible en la barra lateral.

## Vista previa del documento
- Al generar, el servidor crea una miniatura PNG de la 1ª página (best-effort
  vía LibreOffice + pdftoppm) servida en /api/preview/{token}. Si el servidor
  no tiene LibreOffice, degrada con elegancia (sin miniatura, todo lo demás
  funciona). Recupera la "verificación visual" que tenía la app de escritorio.

## Interfaz (index.html reescrito)
- Pantalla de login / registro.
- Barra lateral: logo, "Nuevo documento", historial de conversaciones y menú de
  usuario (avatar, nombre, plan, contador X/Y, "Mejorar a Pro", cerrar sesión).
- Modo oscuro (persistente) — el logo oscuro encaja de forma natural.
- Tarjeta de documento con miniatura de vista previa + descargar.
- Manejo de errores: 429 (demasiadas solicitudes), 401 (sesión expirada),
  límite de plan (modal de mejora), errores del motor.
- Responsive con barra lateral deslizable en móvil; favicon y Open Graph.

## Notas de despliegue
- Definir WEPAI_SECRET_KEY en producción (obligatorio para sesiones seguras).
- La vista previa requiere LibreOffice y poppler (pdftoppm) en el servidor.
- La pasarela de pago (Stripe) para "Mejorar a Pro" se conecta en el despliegue;
  el botón ya está en la UI.
