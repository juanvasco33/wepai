# Reporte de Auditoría — WEP AI

**Proyecto:** WEP AI · Asistente de Office con IA (Word · Excel · PowerPoint) para Colombia
**Versión auditada:** v16.2 → corregida a **v16.3**
**Fecha:** 29 de mayo de 2026
**Alcance:** análisis estático y dinámico del código, suite de pruebas, generación real de documentos, revisión de seguridad y análisis competitivo de mercado.

---

## 1. Resumen ejecutivo

WEP AI es un proyecto **maduro y bien construido**, no un prototipo. Tiene una arquitectura limpia por capas, 327 pruebas automatizadas, 31 changelogs versionados y decisiones de ingeniería que reflejan experiencia real (transacciones atómicas, migraciones de esquema, degradación elegante, validación de seguridad de datos legales).

Se realizó una auditoría que detectó **cuatro hallazgos accionables**, todos corregidos en esta entrega (v16.3). Ninguno comprometía la corrección del documento final; corregían inconsistencias de mensaje, robustez y seguridad operativa. Tras las correcciones, **las 327 pruebas siguen pasando sin regresiones**.

| Métrica | Resultado |
|---|---|
| Archivos Python | 43 |
| Líneas de código | ~30.000 |
| Pruebas (antes / después) | 327/327 ✅ |
| Errores de sintaxis | 0 |
| Hallazgos corregidos | 4 (+ 2 mejoras menores) |
| Secretos hardcodeados | 0 |

---

## 2. Metodología

1. Descompresión e inventario del proyecto (estructura, tamaño por módulo).
2. Lectura de los componentes núcleo: motor IA (`brain.py`), config, base de datos, servidor web, verificadores legales y muestra de los builders de Office.
3. Verificación de sintaxis de los 43 archivos.
4. Ejecución de la suite completa de pruebas (`pytest`).
5. Generación end-to-end de un documento real (contrato de trabajo) para validar el flujo y la calidad de salida.
6. Revisión de seguridad: secretos, SQL injection, XSS, manejo de pagos, verificación de webhooks.
7. Análisis de mercado y competencia en Colombia (investigación web).

---

## 3. Arquitectura

```
agent/    → motor IA + 11 verificadores legales autónomos + overlay/validadores
office/   → builders de Word / Excel / PowerPoint (50+ plantillas)
storage/  → SQLite, Stripe, webhooks, email
ui/       → app de escritorio (CustomTkinter, macOS)
web/      → servidor FastAPI + frontend (acceso multiplataforma)
```

**Fortaleza destacada — verificación legal autónoma:** antes de generar, los verificadores consultan la web para confirmar datos fiscales/laborales vigentes, los validan exigiendo fuente gubernamental oficial y rangos plausibles, y los persisten en un overlay JSON con TTL de 30 días. La filosofía de "fallar seguro" (ante la duda, conservar el valor base conocido y nunca bloquear) está correctamente implementada.

---

## 4. Hallazgos y correcciones

### H1 · System prompt no migrado a Colombia — Severidad: Media — ✅ Corregido

El motor forzaba Colombia + español en el JSON generado y los builders leían el IVA del cfg (19% correcto), pero el `SYSTEM_PROMPT` conservaba 13 residuos mexicanos (IVA 16%, MXN, IMSS, ISR, LFT, RFC). No contaminaba el documento, pero sí lo que la IA *decía* en el chat. **Corrección:** sustituidos los 13 residuos por equivalentes colombianos (IVA 19%, COP, salud/pensión, CST, NIT) y corregido el default de moneda `'MN'` → `'COP'` en dos plantillas de Word.

### H2 · Servidor web: fuga de recursos y sin rate limiting — Severidad: Media — ✅ Corregido

`_DOWNLOAD_TOKENS` no expiraba y los archivos temporales nunca se borraban (fuga de memoria y disco). `/api/chat` no tenía límite de tasa (riesgo de consumo descontrolado de la API). **Corrección:** TTL de 1 hora con barrido automático de tokens y archivos (thread-safe), y rate limiter por IP (15 req/60 s → HTTP 429). Verificado funcionalmente.

### H3 · Discordancia gramatical en el contrato de trabajo — Severidad: Baja — ✅ Corregido

El contrato interpolaba "de **la** {ley}", produciendo "de la Código Sustantivo del Trabajo" (Código es masculino). Resto de las plantillas multi-país. **Corrección:** "del / en el {ley}" en las tres cláusulas afectadas. Verificado regenerando el documento.

### H4 · Acceso frágil a la respuesta del modelo — Severidad: Baja — ✅ Corregido

`response.content[0].text` asumía que el primer bloque siempre es texto. **Corrección:** helper `_first_text()` que recorre los bloques y concatena solo el texto, devolviendo "" si no hay ninguno.

---

## 5. Deuda técnica (no corregida — fuera del alcance inmediato)

- **Builders monolíticos:** `word_controller.py` (7.316 líneas) y `excel_controller.py` (6.603 líneas) concentran 50+ plantillas en un solo archivo cada uno. Recomendación: dividir en módulos por plantilla con un dispatcher.
- **Tokens de descarga en memoria:** el TTL ya resuelve la fuga, pero el mapa sigue en memoria — no sobrevive a reinicios ni escala a múltiples workers. Para producción: Redis o tabla en DB.

---

## 6. Seguridad — sin hallazgos críticos

| Control | Estado |
|---|---|
| Contraseñas con bcrypt (+ migración desde SHA-256) | ✅ |
| SQL parametrizado + whitelist de campos dinámicos | ✅ |
| Verificación de firma del webhook de Stripe | ✅ |
| Escape de HTML en el frontend (anti-XSS) | ✅ |
| Datos de tarjeta: solo últimos 4 dígitos (conciencia PCI) | ✅ |
| Sin secretos hardcodeados | ✅ |
| Consumo de documentos atómico (sin race condition) | ✅ |
| Rate limiting en el servidor web | ✅ (añadido en v16.3) |

---

## 7. Posición competitiva (Colombia)

No existe un competidor que combine exactamente lo mismo: IA conversacional + Word/Excel/PowerPoint + enfoque legal/fiscal colombiano + precio bajo + standalone. Cuatro segmentos compiten por partes:

1. **Microsoft 365 Copilot** — competidor frontal; ya ejecuta tareas agénticas en W/E/P y corre sobre Claude/GPT. Debilidades: requiere licencia M365, caro (~30 USD), sin conocimiento legal colombiano.
2. **Legaltech colombiano** (LEXIUS, Lexivox, Ariel, Sof-IA) — mucho más profundo en lo legal (millones de fuentes con citas verificadas), pero para abogados y caro.
3. **Plantillas gratis** (Bold, GLC, Folou, Canva, MinTrabajo) — estáticas, sin IA; erosionan el "por qué pagar".
4. **SaaS contable** (Alegra, Siigo, World Office) — solapan el lado Excel y tienen **facturación electrónica DIAN válida**, función que WEP AI no posee.

**Hallazgo competitivo crítico:** la "factura" que genera WEP AI en Excel **no es una factura electrónica válida ante la DIAN** (sin CUFE, XML UBL 2.1, Anexo Técnico 1.9), y en Colombia la facturación electrónica es obligatoria. Es un riesgo de expectativa del usuario.

**Nicho defendible:** el dueño de pyme o freelancer colombiano sin licencia de M365, que no es abogado, y quiere generar conversacionalmente cualquier documento con normativa local, barato.

---

## 8. Recomendaciones priorizadas

1. **(Producto)** Aclarar en la UI que la "factura" es una plantilla de control interno, no un comprobante DIAN; o integrar a futuro un proveedor tecnológico DIAN.
2. **(Estrategia)** Priorizar la versión web (ya existe) frente a la dependencia de Office instalado, que es el mayor riesgo arquitectónico vs. la competencia en la nube.
3. **(Calidad)** Validar la verificación legal con una `ANTHROPIC_API_KEY` real (que las búsquedas devuelvan datos correctos en la práctica) — sigue pendiente.
4. **(Mantenibilidad)** Dividir los dos controladores monolíticos por plantilla.
5. **(Legal)** Hacer revisar el disclaimer (v16.2) por un abogado colombiano antes de publicar.

---

*Auditoría realizada sobre el código entregado. Las correcciones de v16.3 están incluidas en el ZIP adjunto. El análisis competitivo se basa en información pública de mayo de 2026.*
