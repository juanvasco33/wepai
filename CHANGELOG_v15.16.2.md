# CHANGELOG v15.16.2 — Auditoría PowerPoint: persistencia + orden de verificación

Fecha: 28 de mayo de 2026
Módulo: `office/ppt_controller.py`
Tipo: extensión de la persistencia a PPT + corrección de orden de verificación

---

## Resumen

Auditoría del módulo PowerPoint con la misma metodología de Word y Excel.
Resultado del enrutamiento: **limpio** — los 18 plans están anunciados en
`brain.py`, no hay huérfanos y no se detectaron colisiones de keywords (cada
deck cae a su plan correcto). Es el módulo más sano de los tres en ese aspecto.

Se encontraron y corrigieron dos problemas en la auto-actualización:

---

## 1. La actualización permanente NO llegaba a las presentaciones (CRÍTICO)

### Problema
`generate_powerpoint` resolvía el país leyendo `LEGAL_CONFIG` directamente
(`dict(LEGAL_CONFIG[code])`) en el caso normal — cuando el usuario especifica
país. Solo el caso sin país pasaba por `_get_legal_config`, que aplica overlay.
Consecuencia: en el flujo habitual, las presentaciones usaban los datos base
del código y las actualizaciones verificadas por IA (persistidas en el overlay)
**no llegaban a los slides**. Mismo tipo de brecha que tenía Excel.

### Solución
`generate_powerpoint` ahora aplica `apply_overlay(code, cfg)` siempre que
resuelve un país, igual que Word y Excel. El cfg que reciben los planners ya
trae los valores actualizados (autoridad fiscal, IVA, moneda, etc.).

### Verificación (end-to-end)
Probado en 3 plans financieros (pitch, investor_update, board): tras persistir
un cambio (autoridad fiscal CO → "DIAN-2026", IVA, moneda) y reiniciar, los tres
planners reciben el cfg con `_overlay_applied=['autoridad_fiscal','iva','moneda']`.

---

## 2. Orden frágil: leer variables ANTES de verificar (11 plans)

### Problema
11 planners leían las variables de localización del cfg
(`sym`, `moneda`, `pais_nombre`, `tax_authority`) ANTES de llamar al verifier
en la misma función. Si el verifier detectaba un cambio en vivo durante esa
generación, las variables ya estaban "congeladas" con el valor anterior y el
cambio no se reflejaba en el slide.

### Solución
Se reordenó en los 11 plans (pitch + 10 más: sales, report, company, product,
strategy, board, investor_update, client_onboarding, all_hands, upsell):
el verifier ahora corre PRIMERO y las variables se leen DESPUÉS del cfg ya
actualizado.

### Verificación
0 plans con el patrón frágil restante.

---

## Regresión
- Smoke test PPT completo: **108/108 OK, 0 crashes** (18 plans × 6 países).
- Sin huérfanos, sin colisiones de enrutamiento.
- Todos los `.py` compilan.

## Archivos modificados
- `office/ppt_controller.py` — overlay en la resolución de país + reorden de
  verificación en 11 plans.

## Estado FINAL de los tres módulos
- **Word** (v15.15.1 + v15.16.0): enrutamiento corregido + auto-actualización
  permanente en documentos.
- **Excel** (v15.16.1): enrutamiento corregido + auto-actualización permanente
  en celdas.
- **PowerPoint** (esta versión): enrutamiento ya estaba limpio +
  auto-actualización permanente en slides + orden de verificación corregido.

Los tres módulos comparten ahora el mismo sistema de persistencia
(`agent/legal_overrides.py`): los cambios legales verificados por IA se guardan
en `~/.wepai/legal_overrides.json`, sobreviven reinicios y llegan al documento,
celda o slide final.

## Pendiente para release público
- Validar la calidad de las búsquedas web reales de los 12 verifiers con
  ANTHROPIC_API_KEY en producción (la arquitectura y persistencia ya están
  probadas end-to-end).
