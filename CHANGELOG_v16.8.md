# CHANGELOG v16.8 — Detección y notificación de cambios estructurales

Fecha: 30 de mayo de 2026
Módulos: agent/structural_alerts.py (NUEVO), agent/fiscal_verifier.py,
main.py, web/server.py, tests/test_structural_alerts.py (NUEVO)

## Resumen

La auto-actualización podía mantener VALORES al día (un número que cambia), pero
no podía aplicar cambios ESTRUCTURALES de fondo (una reforma que cambia reglas,
no cifras) — eso requiere un desarrollador. Esta versión añade el eslabón que
faltaba: un mecanismo que DETECTA señales de cambio estructural y NOTIFICA, en
vez de intentar aplicarlas a ciegas. Es un "detector de humo": no apaga el
incendio, avisa a tiempo.

## 1. Módulo de alertas estructurales (NUEVO) — agent/structural_alerts.py

- DETECCIÓN en dos vías:
  • Heurística (sin API key): señales lingüísticas de reforma / derogación /
    nuevo procedimiento / nueva modalidad / "obligatorio a partir de...".
  • IA (con API key): clasifica un cambio confirmado como VALUE (auto-actualizable)
    o STRUCTURAL (requiere programador). Degrada a la heurística si no hay key.
- PERSISTENCIA: las alertas se guardan en ~/.wepai/structural_alerts.json con
  escritura atómica; no se pierden entre sesiones y no se duplican (firma SHA-256).
- NOTIFICACIÓN en dos niveles:
  • `user_notice()`: aviso SUAVE y no técnico para el usuario final.
  • `developer_report()`: reporte DETALLADO para ti (qué cambió, fuente, qué
    actualizar, cómo resolver).
- RESOLUCIÓN: `resolve_alert(id, nota)` marca una alerta como implementada; deja
  de notificarse pero queda en el historial para auditoría.
- Garantía: ninguna función lanza excepciones; la detección jamás tumba la
  generación de un documento.

## 2. Integración con la auto-verificación — agent/fiscal_verifier.py

- Tras confirmar cambios legales (laborales y fiscales) desde fuentes .gov.co, el
  verificador llama a `detect_and_alert()`. Los cambios de VALOR se aplican solos
  como antes; los ESTRUCTURALES levantan una alerta para el desarrollador.

## 3. Visibilidad para el desarrollador y el usuario

- `main.py`: al iniciar la app de escritorio, si hay alertas estructurales
  pendientes, se imprime el reporte de desarrollador en consola.
- `web/server.py`:
  • `/api/health` ahora reporta `structural_alerts_pending`.
  • Nuevo `/api/admin/structural-alerts` (protegido por `WEPAI_ADMIN_TOKEN` vía
    cabecera `X-Admin-Token`) para revisar las alertas desde la web.

## 4. Pruebas

- tests/test_structural_alerts.py: 11 pruebas nuevas (heurística, persistencia,
  no-duplicación, resolución, avisos diferenciados usuario/dev, robustez).
- Demostración funcional end-to-end ejecutada: una reforma simulada genera 1
  alerta estructural mientras un cambio de salario mínimo se trata como valor
  (sin alerta), el usuario ve el aviso suave, el desarrollador ve el detallado,
  y la resolución limpia la pendiente.
- Suite total: 358/358 ✅ (347 previas + 11 nuevas), sin regresiones.

## Cómo lo usas tú (el dueño/desarrollador)

1. Cuando inicies la app (o consultes `/api/admin/structural-alerts`), verás si
   hay cambios legales de fondo que el sistema no puede aplicar solo.
2. Implementas el cambio en la plantilla/config correspondiente.
3. Llamas a `structural_alerts.resolve_alert(id, "nota")` para cerrarla.

## Configuración nueva

- `WEPAI_ADMIN_TOKEN` (opcional): token para el endpoint de admin de alertas en
  la web. Sin él, el endpoint responde 403.
