# CHANGELOG v15.16.0 — Auto-actualización legal PERMANENTE por país

Fecha: 28 de mayo de 2026
Módulos: `agent/` (nuevo módulo de persistencia + 12 verifiers), `office/legal_config.py`, `main.py`
Tipo: feature mayor — convierte la verificación efímera en persistencia permanente

---

## Resumen

Hasta v15.15.x, cuando un verifier de IA detectaba un cambio legal (ej. el IVA
subió), lo aplicaba SOLO al documento en curso y lo anotaba en un log de texto
con la nota "actualizar el código manualmente". El cambio **no persistía**: el
siguiente usuario volvía a pagar la búsqueda web y, si nadie editaba el código,
el dato seguía desactualizado para siempre.

Esta versión implementa **persistencia permanente por país**: los cambios
legales confirmados por IA se guardan y sobreviven reinicios, beneficiando a
todos los usuarios siguientes sin intervención de un desarrollador. La app
ahora "aprende" de forma permanente.

---

## 1. Persistencia permanente (lo principal)

### Nuevo módulo `agent/legal_overrides.py`
Capa de persistencia basada en un overlay JSON (`~/.wepai/legal_overrides.json`):

- **Los verifiers ESCRIBEN** cada cambio confirmado, con valor, fuente oficial,
  fecha de verificación y categoría.
- **`legal_config.py` LEE** al resolver cada país, aplicando el overlay sobre
  los valores base del código. Los cambios sobreviven reinicios.

### Por qué overlay JSON y no reescribir `legal_config.py`
Reescribir un `.py` con dicts literales de 20 países es frágil (un fallo a
media escritura tumba la app). El overlay JSON es:
- **Seguro**: escritura atómica (archivo temporal + rename).
- **Auditable**: cada cambio guarda fuente y fecha.
- **Reversible**: `clear_overlay(pais)` restaura los valores base.
- **No invasivo**: nunca toca el código fuente.

### Política de frescura (TTL)
Cada override guarda su fecha. Se considera vigente `OVERRIDE_TTL_DAYS` (30)
días; pasado el plazo, el verifier vuelve a confirmarlo en web. Evita que un
dato verificado hace mucho se asuma correcto indefinidamente.

### Conectado en los 12 verificadores
fiscal, labor, privacy, ip, real_estate, commercial, notarial, corporate,
accounting, depreciation, banking, inventory. Todos persisten al overlay.

### Verificación
Ciclo end-to-end probado: verifier detecta cambio → guarda → reinicio simulado
→ el cambio sigue presente en el cfg, **sin haber tocado `legal_config.py`**.

---

## 2. Aviso visible cuando la verificación está INACTIVA

Antes, sin API key o sin el SDK `anthropic`, los verifiers fallaban en
silencio y la app generaba con datos estáticos sin avisar.

`verification_status()` (en `legal_overrides.py`) reporta si la auto-
actualización está operativa (requiere SDK + API key) y devuelve un mensaje
claro para el usuario. `main.py` ahora lo muestra al arrancar, distinguiendo
"falta API key" de "falta el SDK".

---

## 3. Verificación de MONEDA

El verifier fiscal ahora también verifica `moneda` y `simbolo`. Antes se
excluían por considerarse "estables", pero en países con reformas monetarias
(Venezuela, Argentina) sí cambian. Solo se reporta cambio ante una reforma
monetaria oficial confirmada por fuente del banco central.

### Verificación
Probado con caso Venezuela: persistencia de moneda (VED) y símbolo (Bs.D)
sobrevive reinicio.

---

## Regresión
- Compilación: todos los `.py` del proyecto compilan sin error.
- Smoke test Word: **228/228 OK, 0 crashes** (sin cambios).
- El sistema funciona idéntico cuando la verificación está inactiva (degradación
  elegante a datos estáticos + overlay guardado).

## Archivos nuevos
- `agent/legal_overrides.py` — módulo de persistencia + diagnóstico de estado.

## Archivos modificados
- `office/legal_config.py` — aplica el overlay al resolver cada país.
- `agent/fiscal_verifier.py` — persiste fiscal y labor; verifica moneda/símbolo.
- `agent/corporate_verifier.py` y los otros 9 verifiers — persisten al overlay.
- `main.py` — diagnóstico claro del estado de verificación al arrancar.

## Notas importantes
- La auto-actualización requiere `ANTHROPIC_API_KEY` + SDK `anthropic`. Sin
  ellos la app funciona con datos guardados (ya no falla en silencio: avisa).
- La calidad de las búsquedas web reales aún debe validarse con tráfico real;
  la arquitectura y la persistencia ya están probadas end-to-end.
