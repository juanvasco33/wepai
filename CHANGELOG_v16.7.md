# CHANGELOG v16.7 — Bache 1: Reforma Laboral (Ley 2466) + Modo Diagnóstico

Fecha: 30 de mayo de 2026
Módulos: office/legal_config.py, agent/labor_reform_diagnostic.py (NUEVO),
agent/brain.py, agent/fiscal_verifier.py, tests/test_labor_reform_diagnostic.py (NUEVO)

## Resumen

Esta versión incorpora el "bache 1" identificado en el análisis de mercado: la
Reforma Laboral colombiana (Ley 2466 de 2025), que cambió las reglas de fondo de
la contratación y a la que el segmento de microempresas y pymes no tiene quién lo
acompañe de forma asequible. Pasa a WEP AI de "solo generar documentos" a también
"diagnosticar contratos existentes frente a la reforma" — en modo alerta
educativa, nunca dictamen legal.

## 1. Conocimiento laboral actualizado a la Ley 2466 de 2025

`office/legal_config.py` (bloque "CO"):
- Jornada actualizada de 47 h a 42 h/semana (Art. 161 CST, mod. Ley 2101/2021 y
  Ley 2466/2025).
- Ley laboral referenciada ahora incluye la modificación por Ley 2466 de 2025.
- Tribunales corregidos a "Jueces Laborales del Circuito".
- Nuevos campos de la reforma: contrato indefinido como regla general, límite de
  4 años al término fijo, jornada nocturna a las 7 p.m., recargo dominical
  escalonado hacia el 100%, procedimiento disciplinario de 7 etapas previo al
  despido, exigencia de forma escrita, cuota de aprendices.

## 2. Módulo de diagnóstico (NUEVO) — agent/labor_reform_diagnostic.py

Punto de entrada `run_diagnostic(datos, texto_contrato)`. Dos modos:
- MODO REGLAS (siempre, sin API key): 8 reglas deterministas que señalan puntos
  a revisar a partir de datos estructurados del contrato.
- MODO IA (si hay ANTHROPIC_API_KEY + SDK): complemento que lee texto libre del
  contrato. Degrada con elegancia a [] si no hay key — el diagnóstico nunca
  depende de él.

Garantías de seguridad legal (verificadas por tests):
- Lenguaje SIEMPRE condicional ("podría", "te recomendamos revisar"); nunca
  afirma que un contrato es ilegal o inválido.
- Cada alerta cita la fuente normativa (Ley 2466 / artículo CST).
- Cada alerta recomienda revisión por un profesional.
- Disclaimer visible al inicio y cierre de todo diagnóstico.
- El resultado se enmarca como material de referencia, no documento final.

## 3. Integración con el motor — agent/brain.py

- Nuevo protocolo en el SYSTEM_PROMPT: cuando el usuario quiere REVISAR un
  contrato existente (no crear uno), el motor recoge los datos y emite
  "##DIAGNOSTICAR##" en vez de "##GENERAR##".
- `chat()` maneja la señal, corre el diagnóstico y devuelve `diagnose` y
  `diagnostic` en su dict de retorno (sin romper el contrato existente de
  `generate`/`gen_data`).

## 4. Auto-actualización extendida — agent/fiscal_verifier.py

- El verificador laboral (`verify_labor_data`) ahora también vigila los campos
  de la reforma (duración máxima del fijo, inicio de jornada nocturna, recargo
  dominical) y busca específicamente "reforma laboral Colombia", "Ley 2466" y
  "recargo dominical" en fuentes .gov.co.
- Mantiene la regla de seguridad: solo acepta cambios con fuente gubernamental
  oficial; rechaza noticias y blogs.

## 5. Pruebas

- tests/test_labor_reform_diagnostic.py: 20 pruebas nuevas (reglas individuales,
  garantías de lenguaje seguro, citación de fuente, degradación sin API key,
  config actualizada).
- Suite total: 347/347 ✅ (327 previas + 20 nuevas), sin regresiones.

## Limitaciones conocidas (lo que falta para producción)

- El contenido jurídico DEBE ser revisado por un abogado laboralista colombiano
  antes de publicarse. Este módulo es una base correcta, no un producto blindado.
- La auto-verificación legal sigue sin validarse con una ANTHROPIC_API_KEY real
  (pendiente desde v16.3).
- El recargo dominical exacto por etapa y el umbral preciso de microempresa
  deben confirmarse contra el texto vigente de la norma en el momento de uso.
- La auto-actualización mantiene VALORES al día, no reescribe la ESTRUCTURA de
  las plantillas ante una nueva reforma de fondo (eso requiere intervención de
  un desarrollador).
