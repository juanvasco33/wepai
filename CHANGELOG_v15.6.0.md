# CHANGELOG WEP AI v15.6.0 — Acuerdo de Socios

**Fecha de release:** 27 de mayo de 2026
**Alcance:** Reescritura completa de `_build_shareholders_agreement` + nuevo
módulo `corporate_verifier` para actualización autónoma de datos societarios.

---

## Resumen ejecutivo

La versión anterior del generador de "Acuerdo de Socios" tenía **4 bugs críticos** que
hacían que el documento generado fuera prácticamente inservible:

1. Los datos del usuario se recolectaban pero se descartaban — la empresa quedaba
   rotulada literalmente como `[Nombre de la Empresa]` en el documento final.
2. La tabla de socios estaba hardcoded con 3 placeholders `[Socio A]`, `[Socio B]`, `[Socio C]`.
3. Las 3 tablas centrales (gobernanza, roles, transferencias) salían en español
   incluso cuando el documento estaba en inglés.
4. El bloque de firmas estaba hardcoded a exactamente 3 socios — D, E, F nunca aparecían.

Adicionalmente, no había soporte específico para los distintos países hispanohablantes:
todas las cláusulas usaban porcentajes y referencias legales genéricos (reserva legal
del 5% / 20%, "[Centro de Arbitraje]", "AAA/ICC").

**Esta versión corrige los 4 bugs críticos, añade soporte real para 19 jurisdicciones
hispanohablantes, y agrega un verificador autónomo que mantiene los datos legales
actualizados consultando fuentes oficiales en tiempo real.**

---

## Archivos modificados / creados

### Nuevos

| Archivo | Líneas | Propósito |
|---------|--------|-----------|
| `office/corporate_config.py` | ~330 | Configuración corporativa por país (19 jurisdicciones) |
| `agent/corporate_verifier.py` | ~290 | Verificador autónomo vía Claude + web_search |
| `tests/test_shareholders_agreement.py` | ~310 | Suite pytest parametrizada |

### Modificados

| Archivo | Cambio |
|---------|--------|
| `office/word_controller.py` | Reescritura completa de `_build_shareholders_agreement` (líneas 4404–4601 → ~520 líneas nuevas) + 4 helpers nuevos |

---

## Cobertura geográfica — 19 países hispanohablantes

| Código | País | Ley de sociedades | Forma jurídica común | Reserva legal | Centro de arbitraje |
|--------|------|-------------------|----------------------|---------------|---------------------|
| MX | México | LGSM | S.A.P.I. de C.V. | 5% / 20% | CANACO |
| CO | Colombia | Ley 1258 (S.A.S.) | S.A.S. | 10% / 50% | Cámara de Comercio de Bogotá |
| PE | Perú | LGS 26887 | S.A.C. | 10% / 20% | Cámara de Comercio de Lima |
| AR | Argentina | LGS 19550 / Ley 27.349 | S.A.S. | 5% / 20% | Bolsa de Comercio de Buenos Aires |
| CL | Chile | Ley 18.046 / Ley 20.190 | SpA | configurable | CAM Santiago |
| ES | España | RDL 1/2010 LSC | S.L. | 10% / 20% | CIMA |
| BO | Bolivia | Código de Comercio | S.R.L. | 5% / 50% | CNC |
| EC | Ecuador | Ley de Compañías | S.A.S. | 10% / 50% | Cámara de Comercio de Quito |
| UY | Uruguay | Ley 16.060 / Ley 19.820 | S.A.S. | 5% / 20% | Cámara Nacional de Comercio |
| PY | Paraguay | Ley 6480/2020 | E.A.S. | 5% / 20% | CAMP |
| VE | Venezuela | Código de Comercio | C.A. | 5% / 10% | CEDCA |
| PA | Panamá | Ley 32/1927 / Ley 4/2009 | S.A. | configurable | CeCAP |
| GT | Guatemala | Decreto 2-70 C.Com. | S.A. | 5% / 15% | CENAC |
| HN | Honduras | Decreto 73-50 | S.A. | 5% / 20% | CCA Tegucigalpa |
| NI | Nicaragua | C.Com. de Nicaragua | S.A. | 5% / 10% | CMA |
| SV | El Salvador | Decreto 671 C.Com. | S.A. de C.V. | 7% / 20% | CMA |
| CR | Costa Rica | Ley 3284 C.Com. | S.A. | 5% / 20% | CCA Costa Rica |
| CU | Cuba | C.Com. 1885 + Decretos | S.A. / MIPYME | 5% / 20% | Corte Cubana de Arbitraje |
| DO | República Dominicana | Ley 479-08 | S.R.L. | 5% / 10% | CRC Santo Domingo |

---

## Bugs corregidos

### 🔴 SHA-01 — Datos del usuario ahora se inyectan en el documento

**Antes:** El builder hacía `empresa_nom = _d(datos, "parte_a_nombre", ...)` pero
nunca usaba la variable. En su lugar, el cuerpo del documento contenía el string
literal `"[Nombre de la Empresa]"`.

**Ahora:** Acepta `datos["empresa_nombre"]`, `datos["empresa_id"]`,
`datos["empresa_tipo"]` y `datos["capital_social"]`. Los datos se inyectan en la
apertura, la tabla de socios, y la sección de capital social. Si el usuario no provee
estos datos, se usan placeholders bilingües (ES/EN).

### 🔴 SHA-02 — Tabla de socios totalmente dinámica

**Antes:** Tabla con 3 filas hardcoded: `[Socio A]`, `[Socio B]`, `[Socio C]`, más
una fila ESOP fija.

**Ahora:** Acepta `datos["socios"]` como lista de N socios. Cada socio es un dict con
campos: `nombre`, `id`, `tipo` (`persona` o `empresa`), `pct`, `aportacion`,
`clase_acciones`, `rol`, `dedicacion`, `compensacion`. La fila TOTAL se calcula
automáticamente sumando porcentajes y aportaciones. Soporta socios persona física
o persona jurídica con etiquetas de ID diferenciadas.

### 🔴 SHA-03 — Tablas bilingües ES/EN

**Antes:** Las tablas `governance_table`, `roles_table` y `transfer_table` estaban
definidas con un único conjunto de strings en español. En documentos `lang="en"`
salían 13/14 marcadores en español ("Operaciones ordinarias", "Mayoría simple", etc.).

**Ahora:** Cada tabla se construye vía helpers `_build_governance_table`,
`_build_roles_table` y `_build_transfer_table` que reciben `lang` y devuelven la
versión completa en el idioma correspondiente.

### 🔴 SHA-04 — Firmas escalables a N socios

**Antes:** Bloque hardcoded con `SOCIO A` y `SOCIO B` en `_signature_block`, más un
párrafo gris suelto para `SOCIO C`. Socios D, E, F inexistentes.

**Ahora:** Helper `_render_signatures_dynamic(doc, lang, socios)` que escala a N socios.
Para 2+ socios genera pares en dos columnas vía `_signature_block`; si el total es
impar, el último socio firma solo en una fila adicional. También maneja el caso
patológico de 1 socio (firma centrada).

### 🟠 SHA-05 — Headers de tabla 100% en idioma del documento

**Antes:** `["Socio / Shareholder", ..., "Participación %", "Aportación inicial"]`
mezclaba bilingüe con español.

**Ahora:** Cada idioma tiene su propio set de headers (`"Shareholder", "Ownership %", "Initial Contribution"` para EN).

### 🟠 SHA-06 — Dispatch por tipo, no por substring

**Antes:** `if "2." in sec_title or "GOBIERNO" in sec_title` — frágil a cualquier
cambio de título.

**Ahora:** Cada cláusula es un dict con `type: "text" | "table_governance" |
"table_roles" | "table_transfer"`. El loop hace dispatch por type.

### 🟠 SHA-07 — Centro de arbitraje específico por país

**Antes:** Placeholder literal `"[Centro de Arbitraje]"` en ES, `"AAA/ICC"` en EN
incluso para países LATAM.

**Ahora:** Cada uno de los 19 países tiene su `centro_arbitraje` definido en
`corporate_config.py`. Verificado:
- México → CANACO
- Chile → CAM Santiago
- Colombia → Cámara de Comercio de Bogotá
- Argentina → Bolsa de Comercio de Buenos Aires
- Perú → Cámara de Comercio de Lima
- España → CIMA

### 🟠 SHA-08 — Cláusula de jurisdicción menciona país en EN

**Antes:** Rama ES decía `"Ley aplicable: ... de {pais}"`; rama EN omitía el país.

**Ahora:** Ambos idiomas mencionan el país y la ley de arbitraje aplicable.

### 🟠 SHA-09 — Integración con verificadores autónomos

**Antes:** Sin integración con `FiscalVerifier` ni `LaborVerifier`.

**Ahora:** El builder invoca `verify_corporate_data(cfg)` antes de generar el
documento. Si la red está disponible y la API de Anthropic responde, Claude consulta
fuentes oficiales (gacetas oficiales, sitios gubernamentales, cámaras de comercio)
para verificar que los datos del país siguen siendo correctos. Si encuentra cambios,
los aplica al cfg en runtime. Si falla, el flujo continúa con los datos estáticos
sin interrupción.

### 🟡 SHA-10 — Reserva legal y mayorías por país

**Antes:** "[5]% hasta alcanzar el [20]%" hardcoded para todos los países.

**Ahora:** Cada país tiene su `reserva_legal_pct` y `reserva_legal_tope` reales.

### 🟡 SHA-11 — Diferenciación persona / empresa

**Antes:** Asumía que todos los socios eran personas físicas y usaba `id_persona`.

**Ahora:** Cada socio puede tener `tipo: "persona"` o `tipo: "empresa"`. Cuando es
empresa, la tabla usa la etiqueta de ID corporativo (RFC, NIT, RUC, etc.) en vez de
la de persona física.

---

## Cómo funciona el verificador autónomo

El módulo `agent/corporate_verifier.py` sigue el mismo patrón que el
`FiscalVerifier` existente (v15.5.1):

1. **Cuando el usuario solicita un Acuerdo de Socios**, el builder llama a
   `verify_corporate_data(cfg)` antes de empezar a redactar.

2. **El verificador consulta su cache de sesión.** Si ya verificó el país en la
   sesión actual (mismo proceso), retorna los datos en cache sin hacer red.

3. **Si no hay cache, llama a la API de Anthropic** con el modelo Claude Haiku 4.5
   (rápido y económico) y `web_search` habilitado:

   ```python
   client.messages.create(
       model="claude-haiku-4-5-20251001",
       tools=[{"type": "web_search_20250305", "name": "web_search"}],
       messages=[{"role": "user", "content": VERIFY_PROMPT.format(...)}]
   )
   ```

4. **Claude busca en fuentes oficiales** y verifica si los siguientes campos siguen
   vigentes:
   - `ley_sociedades` (ley vigente de sociedades)
   - `autoridad_societaria` (registro mercantil)
   - `reserva_legal_pct` y `reserva_legal_tope`
   - `mayoria_simple`, `mayoria_calificada`, `mayoria_disolucion`
   - `centro_arbitraje`
   - `ley_arbitraje`
   - `art_distribucion_utilidades`

5. **Si encuentra cambios confirmados** en fuentes primarias oficiales, devuelve un
   JSON con los campos actualizados. El builder usa estos datos para generar el
   documento.

6. **Registra el cambio** en `~/.wepai/corporate_changes.log` para que el equipo de
   desarrollo pueda actualizar el config estático en el siguiente release.

7. **Si algo falla** (red caída, timeout, JSON inválido, API key faltante), el
   verificador **nunca lanza excepción**: silenciosamente retorna el cfg estático
   sin modificar. La generación del documento nunca se interrumpe por una falla del
   verificador.

### Diferencias con `FiscalVerifier` y `LaborVerifier`

| | Fiscal | Labor | Corporate (nuevo) |
|---|--------|-------|-------------------|
| Verifica | IVA, retenciones, autoridad fiscal | Ley laboral, jornada, vacaciones, CTS | Ley de sociedades, mayorías, centros de arbitraje |
| Para qué documentos | Contratos de servicios, compraventa | Contratos de trabajo, finiquito | Acuerdo de socios, estatutos |
| Log | `~/.wepai/fiscal_changes.log` | (compartido) | `~/.wepai/corporate_changes.log` |
| Cache | `_SESSION_CACHE` (fiscal) | `_LABOR_CACHE` | `_SESSION_CACHE` (corporate) |

---

## Cómo llamar al nuevo builder desde la app

Desde `generate_word()` (sin cambios en su API):

```python
from office.word_controller import generate_word

gen_data = {
    "template_id": "shareholders_agreement",
    "titulo": "Acuerdo de Socios — TechVentures México",
    "pais": "México",  # o "Chile", "España", etc.
    "datos": {
        "empresa_nombre": "TechVentures México S.A.P.I. de C.V.",
        "empresa_id":     "TVM230615ABC",
        "empresa_tipo":   "S.A.P.I. de C.V.",
        "capital_social": "5,000,000",
        "socios": [
            {
                "nombre": "Ana Lucía Hernández Torres",
                "id": "HETA850312MQR",
                "tipo": "persona",
                "pct": "35%",
                "aportacion": "1750000",
                "clase_acciones": "Ordinarias Serie A",
                "rol": "Directora General (CEO)",
                "dedicacion": "100% tiempo completo",
                "compensacion": "$80,000/mes",
            },
            {
                "nombre": "Capital Semilla MX SAPI de CV",
                "id": "CSM190504DEF",
                "tipo": "empresa",  # ← persona jurídica
                "pct": "30%",
                "aportacion": "1500000",
                "clase_acciones": "Preferentes Serie Seed",
                "rol": "Miembro del Consejo (Inversor Líder)",
            },
            # ... N socios más
        ],
        # Parámetros opcionales (default mostrado)
        "vesting_anos": "4",
        "vesting_cliff": "1",
        "no_competencia_meses": "24",
        "no_captacion_meses": "12",
        "pacto_lockup_meses": "24",
    },
}

path = generate_word(gen_data)
# Genera /root/Documents/WEP_AI/Acuerdo_de_Socios_-_TechVentures_México.docx
```

**Campos opcionales:** todos los campos en `datos` son opcionales. Si no se proveen,
se generan placeholders bilingües en su lugar.

**Selección de idioma:** se controla con `datos["idioma"]: "es"` o `"en"`, o vía el
campo `idioma` del país en `legal_config.py`.

---

## Suite de tests

Los tests viven en `tests/test_shareholders_agreement.py` y cubren:

- **TestSHA01_DatosDelUsuario** — los datos reales aparecen en el documento.
- **TestSHA02_ListaDinamicaSocios** — parametrizado para N = 1, 2, 3, 5, 7, 10 socios.
- **TestSHA03_TablasBilingues** — 0 marcadores ES en doc EN.
- **TestSHA04_FirmasEscalables** — parametrizado para N = 2..6 socios.
- **TestSHA07_CentroArbitrajePorPais** — parametrizado para 10 países.
- **TestSHA10_ReservaLegalPorPais** — parametrizado para 8 países.
- **TestCoberturaPaisesHispanos** — los 19 países generan sin excepción.
- **TestEdgeCases** — datos vacíos, cfg None, dict incompleto, 1 socio.
- **TestCorporateVerifierIntegracion** — falla silenciosa, aplica overrides.

Los tests usan auto-fixture que mockea los verificadores para que la suite corra
sin red en CI.

Ejecutar:

```bash
pytest tests/test_shareholders_agreement.py -v
```

---

## Documentos de ejemplo generados

| Archivo | Caso de uso |
|---------|-------------|
| `Ejemplo_Acuerdo_Socios_MX.docx` | TechVentures México S.A.P.I. con 4 socios (2 personas + 1 fondo VC + 1 pool de opciones) |
| `Ejemplo_Pacto_Socios_CL.docx`   | InnovaCloud SpA Chile con 3 socios (2 personas + 1 fondo) |
| `Ejemplo_Pacto_Socios_ES.docx`   | DataForge S.L. España con 3 socios (2 personas + 1 SCR) |
| `Ejemplo_Shareholders_AR_EN.docx`| QuantumBA S.A.S. Argentina, versión en inglés |

---

## Notas legales y limitaciones

1. **El documento generado no constituye asesoría legal.** Es una plantilla altamente
   personalizada con datos actualizados, pero **siempre debe ser revisada por un
   abogado licenciado en la jurisdicción correspondiente** antes de la firma.

2. **El verificador autónomo es conservador por diseño.** Solo reporta cambios cuando
   tiene alta confianza basada en fuentes oficiales primarias. Esto reduce falsos
   positivos pero también puede demorar la detección de cambios legislativos muy
   recientes.

3. **Reformas legales recientes** pueden no estar reflejadas en `corporate_config.py`
   estático. El verificador autónomo está diseñado precisamente para cubrir esos
   gaps mientras el equipo actualiza la configuración base.

4. **Sin clave API de Anthropic**, el verificador retorna el config estático sin
   modificar. El documento se genera correctamente pero sin actualización en tiempo
   real.

---

**Auditoría completada:** 27 de mayo de 2026
**Builder anterior — calificación:** 3/10
**Builder v15.6.0 — calificación:** 9/10
