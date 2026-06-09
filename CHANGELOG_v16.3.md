# CHANGELOG v16.3 — Migración a Colombia del prompt + robustez del servidor web

Fecha: 29 de mayo de 2026
Módulos: agent/brain.py, web/server.py, office/word_controller.py
Tipo: corrección de coherencia (enfoque Colombia) + robustez y seguridad

Esta versión cierra cuatro hallazgos detectados en una auditoría de código.
Ningún cambio altera el comportamiento de generación correcto que ya existía;
corrige inconsistencias y endurece el servidor web.

## 1. System prompt completamente migrado a Colombia (agent/brain.py)

El motor ya forzaba Colombia + español en el JSON `##GENERAR##` y los builders
leían el IVA del cfg (19% para CO), pero el `SYSTEM_PROMPT` aún conservaba
residuos de la etapa multi-país (origen mexicano). Esto NO afectaba el documento
final, pero sí lo que la IA **decía al usuario** en el chat (podía preguntar por
el IMSS o mencionar IVA 16%). Se eliminaron 13 residuos:

- "IVA 16%" → "IVA 19%" (tasa general en Colombia).
- "equivalente MXN" → "equivalente COP" (multimoneda, subscripciones).
- "¿Moneda base (MXN, COP, etc.)?" → "¿Moneda base (por defecto COP)?".
- "ISR estimado" → "retención en la fuente estimada".
- "deducciones (IMSS, ISR)" → "salud 4%, pensión 4% y retención en la fuente".
- "Vacaciones según LFT" → "Vacaciones según el CST (15 días hábiles/año)".
- "días LFT" → "días según el CST".
- "RFC/NIT" / "RFC/RUC/NIT/..." → "NIT (empresas) · cédula (personas)".
- Notas de bug (`BUG_CHECKS`): IVA 16%→19%, MXN→COP, ISR 35%→retención.

Además se corrigió el valor por defecto de moneda en dos plantillas de Word:
`cfg.get('moneda','MN')` → `cfg.get('moneda','COP')` (MN = Moneda Nacional
mexicana). Afectaba el contrato de trabajo y una carta de cobro cuando el cfg
no traía el código de moneda.

## 2. Servidor web: expiración, limpieza y rate limiting (web/server.py)

Tres problemas de producción en el servidor FastAPI (v16.0+):

- **Fuga de memoria/disco:** `_DOWNLOAD_TOKENS` era un dict sin expiración y los
  archivos en el directorio temporal nunca se borraban. Crecían sin límite.
  Ahora cada token guarda su `created` y un barrido (`_sweep_expired_downloads`)
  elimina tokens y archivos vencidos (TTL = 1 hora) en cada generación y descarga.
  El acceso al mapa de tokens está protegido con lock (thread-safe).
- **Sin rate limiting:** `/api/chat` quedaba expuesto a abuso → consumo
  descontrolado de la API de Anthropic. Se añadió un limitador por IP de ventana
  deslizante (15 peticiones / 60 s) que responde HTTP 429 al excederse.
- Versión del servidor: 16.0 → 16.3.

## 3. Corrección gramatical en el contrato de trabajo (office/word_controller.py)

El contrato laboral interpolaba el nombre de la ley con artículo femenino fijo:
"de **la** {ley_lab}", produciendo "de la Código Sustantivo del Trabajo"
(discordancia de género; "Código" es masculino). Era un resto de las plantillas
multi-país donde la ley era femenina ("la Ley Federal del Trabajo"). Corregido a
"del / en el {ley_lab}" en las tres cláusulas afectadas (tipo de relación,
prestaciones de ley, rescisión).

## 4. Robustez al leer la respuesta del modelo (agent/brain.py)

`chat()`, `analyze_screenshot()` y `generate_title()` accedían a
`response.content[0].text` a ciegas. Si el primer bloque de la respuesta no
fuera de texto (p. ej. un bloque de herramienta), lanzaría AttributeError/
IndexError y tumbaría el flujo. Se añadió el helper `_first_text(response)` que
recorre los bloques y concatena solo el texto, devolviendo "" si no hay ninguno.

## Pruebas

- Suite completa: **327/327 tests pasan** (sin regresiones).
- Compilación de los 43 archivos .py sin errores de sintaxis.
- Rate limiter verificado: bloquea a partir de la 16ª petición en la ventana.
- Sweep de TTL verificado: elimina tokens vencidos y conserva los vigentes.
- Contrato regenerado: la concordancia ahora lee "en el Código Sustantivo del
  Trabajo" correctamente.
- Generación end-to-end (Word) funciona con verificadores caídos (fail-safe).

## Sin cambios

No se tocó el esquema de la base de datos, el flujo de pago (Stripe), la UI de
escritorio ni la lógica de los 11 verificadores legales. Cambios acotados a los
cuatro puntos anteriores.
