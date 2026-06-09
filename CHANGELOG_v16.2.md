# CHANGELOG v16.2 — Actualización por tipo de documento + Disclaimer legal

Fecha: 29 de mayo de 2026
Módulos: agent/*_verifier.py (11), agent/legal_validators.py,
         LEGAL_DISCLAIMER_ES.md (NUEVO), web/static/index.html
Tipo: robustez de actualización legal + protección legal del producto

## 1. La actualización SÍ depende del tipo de documento (verificado)

Se auditó el enrutamiento de verificadores y se confirmó que la actualización
legal es **contextual al documento solicitado**, no genérica:

- Contrato de trabajo / nómina / liquidación → `verify_labor_data`
- Cotización / factura / contrato de servicios → `verify_fiscal_data`
- Acta de constitución / estatutos / SAS → `verify_corporate_data`
- Contrato de cesión de PI / licencia → `verify_ip_data`
- Política de privacidad / tratamiento de datos → `verify_privacy_data`
- Poder / documento notarial → `verify_notarial_data`
- Contrato de arrendamiento → `verify_real_estate_data`
- Distribución / agencia comercial → `verify_commercial_data`
- Estados financieros → `verify_accounting_data`
- Depreciación de activos → `verify_depreciation_data`
- Documentos bancarios → `verify_banking_data`
- Inventarios → `verify_inventory_data`

Cada verificador busca SOLO los datos relevantes a su dominio, lo que reduce el
margen de error: no se "inventa" verificar datos que el documento no necesita.

## 2. Los 11 verificadores ahora exigen fuente oficial

**Brecha corregida:** antes solo el verificador fiscal pasaba `require_official`
y `current_cfg` a `save_overrides`. Los otros 10 podían persistir cambios sin
exigir fuente gubernamental. Ahora **los 11** exigen que el cambio provenga de
una fuente oficial (.gov.co, DIAN, MinTrabajo, etc.) y comparan contra el valor
anterior. Sin fuente oficial, el cambio se rechaza y se conserva el valor base.

## 3. Validación de porcentajes genéricos

`legal_validators.py` ahora valida automáticamente cualquier campo terminado en
`_pct` (depreciación, reserva legal, merma, etc.): debe ser un porcentaje
numérico entre 0 y 100. Esto cubre los campos numéricos de los verificadores
societario, de depreciación e inventario.

## 4. Disclaimer y términos legales (NUEVO)

- `LEGAL_DISCLAIMER_ES.md` — términos de uso completos en español, redactados
  como base profesional. **Requiere revisión de un abogado colombiano** antes de
  publicar (marcado en el propio documento): Ley 1480 (consumidor), Ley 1581
  (habeas data), Ley 527 (comercio electrónico).
- El frontend web ahora muestra:
  - Un **modal de aceptación** en el primer uso, que explica con honestidad qué
    es y qué no es WEP AI antes de permitir generar documentos.
  - La **nota corta** permanente bajo el cuadro de texto.
  - Enlace a los términos completos.

Mensaje central del disclaimer (clave para la protección legal): WEP AI genera
borradores, usa fuentes oficiales cuando existen, NO es asesoría legal, ningún
sistema garantiza exactitud absoluta, y el usuario debe revisar cada documento.

## Pruebas

- `tests/test_legal_validators.py`: 26 tests (4 nuevos de porcentajes).
- Suite completa: 327/327 tests pasan.
- Los 11 verificadores importan sin error tras los cambios.
- Servidor web: sirve el disclaimer (HTTP 200), modal presente en el index,
  flujo de chat intacto.

## Importante (límites honestos)

- Esto reduce el margen de error tanto como es razonablemente posible, pero NO
  garantiza exactitud legal absoluta — ningún sistema de IA con búsqueda web
  puede. El disclaimer refleja esto y es parte central de la protección legal.
- La validación con ANTHROPIC_API_KEY real (que las búsquedas devuelvan datos
  correctos en la práctica) sigue pendiente y es el siguiente paso recomendado.
- El disclaimer debe ser revisado por un abogado colombiano antes de publicarlo.
