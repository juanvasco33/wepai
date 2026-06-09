# CHANGELOG v16.1 — Validación de seguridad + actualización legal robusta

Fecha: 29 de mayo de 2026
Módulos: agent/legal_validators.py (NUEVO), agent/legal_overrides.py,
         agent/fiscal_verifier.py
Tipo: seguridad de datos legales (crítico)

## Contexto

El objetivo es que la auto-actualización legal de WEP AI sea lo más confiable
posible y, sobre todo, que **falle de forma segura**: ante cualquier duda,
conservar el valor base conocido en lugar de arriesgar un dato erróneo que
contaminaría los documentos de todos los usuarios.

NOTA IMPORTANTE — sobre "cero margen de error": ningún sistema de IA con
búsqueda web puede *garantizar* exactitud legal absoluta. Lo que esta versión
logra es minimizar la probabilidad de error con varias capas de defensa y hacer
que el sistema se niegue a actualizar cuando no está altamente seguro. La
exactitud final siempre debe confirmarla un profesional; el disclaimer del
producto refleja esto y es parte de la protección legal.

## 1. Validador de datos legales (NUEVO) — agent/legal_validators.py

Capa que valida cada dato ANTES de persistirlo. Tres controles:

- **Rango / formato:** IVA solo 0-30%; moneda contra lista ISO 4217 de la
  región; etiqueta de impuesto contra lista válida (IVA/IGV/ITBMS/...); textos
  legales con largo razonable y sin valores-basura ("N/A", "unknown", etc.).
- **Salto sospechoso:** un cambio de IVA mayor a 10 puntos se rechaza por
  seguridad (probable error de lectura del modelo).
- **Fuente oficial:** el cambio solo se acepta si la fuente es un dominio
  gubernamental (.gov.co, .gob.*, DIAN, MinTrabajo, etc.). Blogs y noticias
  se rechazan.

## 2. save_overrides ahora valida antes de persistir — legal_overrides.py

- Cada llamada pasa por `validate_changes()`. Solo los campos aprobados se
  guardan; los rechazados se registran para auditoría y NO entran al overlay.
- Si NINGÚN campo pasa, no se persiste nada (fallar seguro).
- Si el propio validador fallara, tampoco se persiste (fallar seguro).
- Nuevos parámetros: `current_cfg` (para comparar contra el valor viejo) y
  `require_official` (exigir fuente gubernamental, por defecto True).

## 3. Actualización constante conectada de verdad — fiscal_verifier.py

- **Brecha corregida:** la política de frescura (TTL 30 días) existía pero el
  verificador nunca la usaba. Ahora, si el overlay persistente tiene todos los
  campos críticos verificados dentro del TTL, se OMITE la búsqueda web. Esto
  hace la re-verificación eficiente: se vuelve a buscar solo cuando un dato
  vence, no en cada documento.
- **Prompt reforzado:** ahora exige explícitamente una URL de fuente
  gubernamental oficial y que el cambio esté YA EN VIGENCIA (no proyectos de
  ley ni borradores). Sin fuente oficial, el verificador devuelve "sin cambios".

## Pruebas

- `tests/test_legal_validators.py` (NUEVO): 22 tests, incluidos los casos
  peligrosos (IVA 80%/99%, fuente de blog, moneda inválida, valor basura) —
  todos rechazados correctamente.
- Suite completa: 323/323 tests pasan (301 previos + 22 nuevos).
- Prueba de integración del ciclo completo: dato absurdo rechazado, fuente no
  oficial rechazada, cambio válido + DIAN aceptado y persistido como fresco,
  mezcla válido/inválido separada correctamente.

## Sobre "actualización verdaderamente constante" (pendiente de decisión)

El sistema actual verifica **en el momento en que el usuario pide un documento**
y persiste lo aprendido con TTL. Esto cubre el caso real: cuando importa, el
dato está fresco. Para una verificación *proactiva diaria* (revisar leyes aunque
nadie pida documentos) se necesitaría un componente de servidor con tarea
programada (cron/worker) — no aplica a una app de escritorio o a un servidor web
sin proceso de fondo. Se puede añadir cuando exista infraestructura de servidor
permanente; el motor de validación ya quedó listo para soportarlo.
