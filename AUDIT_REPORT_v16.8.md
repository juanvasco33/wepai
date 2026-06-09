# Reporte de Auditoría — WEP AI v16.8

**Proyecto:** WEP AI · Asistente de Office con IA para Colombia
**Versión auditada:** v16.7 → entregada como **v16.8**
**Fecha:** 30 de mayo de 2026
**Alcance de esta entrega:** mecanismo de detección y notificación de cambios legales estructurales (fuera del alcance de la auto-actualización), integración con la auto-verificación, visibilidad para desarrollador y usuario, pruebas y no-regresión.

---

## 1. Resumen ejecutivo

Se cerró el eslabón que faltaba en el sistema de auto-actualización. Antes, la app
podía actualizar VALORES sola (un número que cambia) pero los cambios
ESTRUCTURALES (reglas de fondo de una reforma) pasaban desapercibidos hasta que
alguien los notaba. Ahora un mecanismo dedicado los DETECTA y NOTIFICA al
desarrollador, sin intentar aplicarlos a ciegas.

El diseño sigue la filosofía de "fallar seguro" del resto del proyecto: ante un
cambio que no puede aplicar con seguridad, no lo aplica ni bloquea — avisa.

| Métrica | Resultado |
|---|---|
| Archivos Python | 47 (+2 vs v16.7) |
| Pruebas (antes / después) | 347 → 358 ✅ (+11) |
| Errores de sintaxis | 0 |
| Regresiones | 0 |
| Módulos nuevos | 1 (agent/structural_alerts.py) |

---

## 2. Qué se construyó y por qué es correcto

### 2.1 Separación VALOR vs ESTRUCTURAL
El acierto de diseño es no tratar todos los cambios igual. Un cambio de salario
mínimo es un valor: el sistema lo aplica solo. Una reforma que crea una nueva
modalidad de contrato es estructural: requiere reescribir una plantilla, algo que
solo un humano debe hacer. El módulo distingue ambos y solo alerta sobre los
segundos. Esto evita dos fallos opuestos: aplicar a ciegas (peligroso) e ignorar
(silencioso).

### 2.2 Doble vía de detección con degradación
Con API key, una clasificación por IA precisa (VALUE/STRUCTURAL). Sin ella, una
heurística lingüística siempre disponible. El sistema nunca queda ciego por falta
de key — solo menos preciso. Coherente con el resto de la app.

### 2.3 Notificación diferenciada
- Al usuario final: un aviso suave, no técnico, que no alarma ni expone detalles.
- Al desarrollador: un reporte accionable (qué cambió, fuente, qué actualizar).
Esto respeta a ambas audiencias: el usuario no necesita saber de plantillas; el
dueño necesita saberlo todo.

### 2.4 Persistencia robusta
Almacén JSON con escritura atómica (tmp + rename), firma estable para no duplicar,
e historial de resoluciones para auditoría. Mismo patrón que el overlay legal
existente.

---

## 3. Verificación de robustez

| Control | Estado |
|---|---|
| Ninguna función lanza excepciones hacia afuera | ✅ (test) |
| La detección no tumba la generación de documentos | ✅ (envuelta en try/except en los verificadores) |
| Sin API key, degrada a heurística sin fallar | ✅ (test) |
| No duplica alertas en verificaciones repetidas | ✅ (test) |
| Persiste entre reinicios | ✅ (test) |
| Aviso al usuario no expone detalles técnicos | ✅ (test) |
| Endpoint admin protegido por token | ✅ (403 sin token) |

---

## 4. Cómo opera en producción

1. La auto-verificación (que ya corría) confirma un cambio legal desde una fuente
   .gov.co.
2. `detect_and_alert()` clasifica: si es estructural, registra una alerta.
3. El dueño la ve al iniciar la app (consola) o vía `/api/admin/structural-alerts`.
4. Implementa el cambio y lo cierra con `resolve_alert()`.

Mientras tanto, el usuario final ve un aviso suave de prudencia, y el contenido
sigue disponible (no se bloquea nada).

---

## 5. Pendientes (heredados, no técnicos de esta entrega)

1. **(Legal — crítico)** Revisión del contenido jurídico (módulo de diagnóstico
   v16.7 y plantillas laborales) por un abogado colombiano antes de publicar.
2. **(Calidad)** Validar la auto-verificación —y ahora también la clasificación
   de cambios— con una `ANTHROPIC_API_KEY` real.
3. **(Seguridad — heredado)** Forzar fallo de arranque del servidor web si falta
   `WEPAI_SECRET_KEY` en producción.
4. **(Producto — heredado)** Aclarar en la UI que la "factura" no es comprobante
   válido ante la DIAN.

---

## 6. Nota sobre el alcance del mecanismo

Este sistema detecta señales de cambio estructural con buena cobertura, pero no
es infalible: depende de que la auto-verificación traiga el cambio y de que las
señales lingüísticas (o la IA) lo identifiquen. No reemplaza el monitoreo humano
de la normativa; lo complementa, reduciendo la probabilidad de que un cambio de
fondo pase desapercibido. Trátalo como una red de seguridad, no como garantía
absoluta de estar siempre al día.

---

*Auditoría de ingeniería sobre el código entregado en v16.8. La validación
jurídica del contenido sigue siendo un paso humano posterior e indispensable.*
