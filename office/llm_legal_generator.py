"""
WEP AI — Generación de documentos legales con LLM (v15.1, Phase 1)

Reemplaza builders hardcoded para documentos donde el contenido legal es
country-specific. El LLM genera la estructura del documento adaptada a la
ley del país, y un renderer único produce el .docx con formato consistente
usando los helpers de word_controller.

PHASE 1: solo `employment_contract` (contrato de trabajo).
Próximas fases sumarán: termination_letter, settlement, commercial_lease,
work_rules, power_of_attorney.

INTEGRACIÓN en generate_word():

    LLM_DRIVEN_TEMPLATES = {"employment_contract"}
    if template_id in LLM_DRIVEN_TEMPLATES:
        try:
            from office.llm_legal_generator import build_via_llm
            build_via_llm(doc, template_id, instruccion, secciones, detalles, cfg, datos)
            _llm_handled = True
        except Exception as e:
            log.error("LLM legal failed, falling back: %s", e)

VENTAJAS sobre el builder hardcoded:
  - Cita la ley correcta del país (no más LFT en contratos colombianos)
  - Usa la terminología correcta ("empleador" en Colombia, "patrón" en México)
  - Incluye beneficios específicos del país (CTS en Perú, prima de servicios
    en Colombia, SAC en Argentina, etc.)
  - Funciona para CUALQUIER país, no solo los 8 hardcoded en LEGAL_CONFIG.

COSTO:
  - ~$0.10–0.20 USD por documento nuevo (Sonnet ~6–8k tokens)
  - $0.00 si el (país + tipo + parámetros clave) ya está cacheado.

CACHÉ:
  - Ruta: ~/.wepai/llm_legal_cache.json
  - Key: hash de (doc_type, country, lang, parámetros relevantes)
  - Persistido tras cada generación exitosa.

FALLBACK:
  - Si la API falla, JSON inválido, o el doc_type no está implementado,
    la función lanza la excepción para que el caller (word_controller)
    use el builder hardcoded como fallback.
"""

import os
import re
import json
import hashlib
import logging
import threading
from datetime import datetime

import anthropic

log = logging.getLogger("wepai.llm_legal")

# ── CONFIGURACIÓN ─────────────────────────────────────────────────────────────
# v15.4: model ID centralizado en config.py (override por WEPAI_MODEL_LEGAL)
from config import DEFAULT_LEGAL_MODEL as MODEL
CACHE_PATH = os.path.expanduser("~/.wepai/llm_legal_cache.json")

_client = anthropic.Anthropic(
    api_key=os.environ.get("ANTHROPIC_API_KEY", ""),
    max_retries=3,
    timeout=90.0,  # docs legales largos pueden tardar
)

# Caché en memoria + disco
_cache: dict = {}
_cache_lock = threading.Lock()
_cache_loaded = False


# ══════════════════════════════════════════════════════════════════════════════
# SYSTEM PROMPTS POR TIPO DE DOCUMENTO
# ══════════════════════════════════════════════════════════════════════════════
_EMPLOYMENT_CONTRACT_PROMPT = """You are a senior legal document drafting assistant specialized in EMPLOYMENT CONTRACTS across Spanish-speaking jurisdictions and English-speaking common-law jurisdictions.

You produce STRUCTURED JSON OUTPUT that follows strict accuracy rules. Lawyers will review what you generate. Your job is to be RIGHT, not impressive.

═══════════════════════════════════════════════════════
RULE 1 — ACCURACY OVER FLUENCY
═══════════════════════════════════════════════════════

Only cite specific article numbers when you are HIGHLY CONFIDENT the article exists, is currently in force, and addresses the topic you're citing it for. When uncertain:

  ✓ DO use general formulation: "conforme a la legislación laboral vigente del país"
  ✓ DO reference the law's name: "según el Código Sustantivo del Trabajo de Colombia"
  ✗ DO NOT fabricate: "Art. 47 del Código X" if you're not sure of the article

For very small or unfamiliar jurisdictions, default to general formulation. The user (and their lawyer) can refine. Fabricated citations are worse than missing ones.

═══════════════════════════════════════════════════════
RULE 2 — COUNTRY-CORRECT TERMINOLOGY
═══════════════════════════════════════════════════════

Term for "employer" varies critically by country. Do NOT use the wrong term:

  - México:   "patrón" (formal) or "empleador" — both acceptable
  - Colombia: "empleador" — NEVER "patrón" (colloquial, dated, slightly pejorative)
  - Argentina: "empleador" — preferred (also "empresa")
  - Chile:    "empleador"
  - Perú:     "empleador"
  - Bolivia, Ecuador, Paraguay, Uruguay: "empleador"
  - Costa Rica, Panamá, Honduras, Guatemala, etc.: "empleador"
  - España:   "empleador" / "empresa"
  - USA (es): "el Empleador" / USA (en): "the Employer"

Term for "employee":
  - Most LATAM: "trabajador" (formal in contracts) or "empleado"
  - English: "the Employee"

═══════════════════════════════════════════════════════
RULE 3 — CORRECT TAX/ID SYSTEMS
═══════════════════════════════════════════════════════

Use the country's actual identification system. Never mix.

  - México:    RFC (empresas + personas), CURP (personas)
  - Colombia:  NIT (empresas), CC = Cédula de Ciudadanía (personas)
  - Argentina: CUIT (empresas), CUIL/DNI (personas)
  - Chile:     RUT (ambos)
  - Perú:      RUC (empresas), DNI (personas)
  - España:    CIF (empresas), NIF/DNI (personas)
  - USA:       EIN (empresas), SSN (personas)
  - Brasil:    CNPJ (empresas), CPF (personas)
  - Bolivia:   NIT (empresas), CI = Carnet de Identidad (personas)
  - Ecuador:   RUC (empresas), Cédula (personas)
  - Paraguay:  RUC, Cédula
  - Uruguay:   RUT, CI
  - Other LATAM: research the actual local system

═══════════════════════════════════════════════════════
RULE 4 — CORRECT INSTITUTIONS
═══════════════════════════════════════════════════════

Reference the institutions of the SPECIFIC country only:

  - México:    IMSS (seguridad social), AFORE/INFONAVIT (pensión/vivienda), STPS (autoridad laboral), SAT (fiscal)
  - Colombia:  EPS (salud), AFP (pensión), Caja de Compensación, Ministerio del Trabajo, DIAN (fiscal)
  - Argentina: ANSES, Obra Social, Ministerio de Trabajo, AFIP (fiscal)
  - Chile:     AFP (pensión), Fonasa/Isapre (salud), Dirección del Trabajo, SII (fiscal)
  - Perú:      EsSalud (salud), AFP/ONP (pensión), Mintra (trabajo), SUNAT (fiscal)
  - España:    Seguridad Social (TGSS), Servicio Público de Empleo (SEPE), AEAT (fiscal)
  - USA:       Social Security, FLSA, IRS, NLRB

NEVER mix — no "IMSS" in a Colombian contract, no "EPS" in a Mexican contract.

═══════════════════════════════════════════════════════
RULE 5 — COUNTRY-SPECIFIC MANDATORY BENEFITS
═══════════════════════════════════════════════════════

The "Prestaciones de Ley" clause MUST list the benefits actually required by the country's labor law:

  - México: aguinaldo (mínimo 15 días, Art. 87 LFT), vacaciones según LFT con prima vacacional 25%, PTU 10% si aplica, inscripción al IMSS, INFONAVIT

  - Colombia: cesantías (1 mes de salario por año, Art. 249 CST), intereses sobre cesantías (12% anual), prima de servicios (30 días/año en dos cuotas: junio y diciembre), vacaciones 15 días hábiles (Art. 186 CST), auxilio de transporte si salario ≤ 2 SMMLV, afiliación EPS+AFP+Caja de Compensación+ARL

  - Argentina: SAC/aguinaldo (1 sueldo extra/año en dos cuotas: junio + diciembre), vacaciones según antigüedad (14/21/28/35 días), aportes ANSES, Obra Social

  - Chile: gratificación legal (25% utilidad líquida o 4.75 IMM si la empresa tiene utilidades), feriado anual 15 días hábiles, cotización AFP + salud (Fonasa o Isapre)

  - Perú: gratificaciones de julio y diciembre (1 sueldo cada una, Ley 27735), CTS (½ sueldo cada mayo y noviembre, Ley 25920), vacaciones 30 días calendario, asignación familiar 10% RMV si aplica

  - España: 2 pagas extras (junio + Navidad), 30 días naturales de vacaciones, alta en SS, IRPF retención

For countries you're unsure of: list what you know with confidence, and use general formulation for the rest ("y demás prestaciones según la legislación laboral vigente").

═══════════════════════════════════════════════════════
RULE 6 — REQUIRED STRUCTURE
═══════════════════════════════════════════════════════

Produce 8–10 clauses. Standard order:

  1. PRIMERA — Tipo de Relación, Período de Prueba y Duración
     (Include trial period appropriate for country: México 30/180 días, Colombia 2 meses Art. 78 CST, Chile típico 2 meses, Argentina 3 meses Art. 92bis LCT)

  2. SEGUNDA — Objeto / Puesto y Funciones

  3. TERCERA — Jornada de Trabajo
     (Country-specific: México 48hrs/sem, Colombia 47hrs/sem Art. 161 CST, Chile 45hrs/sem, Argentina 48hrs/sem LCT, Perú 48hrs/sem, España 40hrs/sem)

  4. CUARTA — Remuneración / Salario y Forma de Pago

  5. QUINTA — Prestaciones de Ley (use rule 5 above for the specific country)

  6. SEXTA — Lugar de Trabajo

  7. SÉPTIMA — Obligaciones del Trabajador

  8. OCTAVA — Confidencialidad y Propiedad Intelectual
     (For tech/creative roles, include IP assignment clause)

  9. NOVENA — Terminación / Rescisión
     (Cite country-correct articles: Colombia Art. 62 CST, México Art. 47 LFT, etc.)

  10. DÉCIMA — Jurisdicción y Resolución de Conflictos

═══════════════════════════════════════════════════════
RULE 7 — USER DATA INJECTION
═══════════════════════════════════════════════════════

Use the exact names, IDs, amounts, dates provided by the user. For data NOT provided, use [Bracketed Placeholder in target language]:

  Spanish: [Nombre del Empleador], [Dirección del Empleador], [Salario mensual], [Fecha de inicio], [Departamento]
  English: [Employer Name], [Employer Address], [Monthly Salary], [Start Date], [Department]

NEVER invent specific data the user didn't give (don't make up a salary or address).

═══════════════════════════════════════════════════════
RULE 8 — LANGUAGE OUTPUT
═══════════════════════════════════════════════════════

If lang=es: output ENTIRELY in Spanish, using the country's variant register (Colombian Spanish for CO, Mexican for MX, Argentine for AR, etc.)
If lang=en: output ENTIRELY in English with formal legal register

Do not mix languages.

═══════════════════════════════════════════════════════
OUTPUT SCHEMA
═══════════════════════════════════════════════════════

Respond with ONLY a single valid JSON object — no prose before or after, no markdown fences:

{
  "titulo": "string — main document title in target language (e.g. 'CONTRATO INDIVIDUAL DE TRABAJO' or 'EMPLOYMENT CONTRACT')",
  "subtitulo": "string — short subtitle referencing applicable law (e.g. 'Por tiempo indefinido — Código Sustantivo del Trabajo de Colombia')",
  "intro": "string — opening paragraph (1-3 sentences) identifying the parties, jurisdiction, and the act of entering into the contract",
  "clausulas": [
    {
      "titulo": "string — clause heading e.g. 'PRIMERA. — TIPO DE RELACIÓN Y PERÍODO DE PRUEBA'",
      "contenido": "string — full clause body. Substantive (3-6 sentences typical). Use \\n for paragraph breaks within a clause if needed."
    }
  ],
  "closing": "string — closing paragraph immediately before signatures (e.g. 'Leído que fue el presente contrato y enteradas las Partes...')",
  "partes": [
    {"rol": "string e.g. 'EL EMPLEADOR' or 'THE EMPLOYER'", "nombre": "string e.g. 'Acme Software S.A.S.' or '[Nombre del Empleador]'"},
    {"rol": "string e.g. 'EL TRABAJADOR' or 'THE EMPLOYEE'", "nombre": "string e.g. 'Juan Pérez' or '[Nombre del Trabajador]'"}
  ]
}"""


SYSTEM_PROMPTS = {
    "employment_contract": _EMPLOYMENT_CONTRACT_PROMPT,
}


# ══════════════════════════════════════════════════════════════════════════════
# PHASE 2 — SYSTEM PROMPTS DE LOS 5 BUILDERS ADICIONALES
# ══════════════════════════════════════════════════════════════════════════════
# Cada prompt mantiene la misma estructura que employment_contract:
#  - Reglas de anti-alucinación de citas legales
#  - Terminología correcta por país
#  - Instituciones y sistemas de ID correctos
#  - Estructura de cláusulas estándar
#  - Schema JSON consistente con el renderer común

_TERMINATION_LETTER_PROMPT = """You generate STRUCTURED EMPLOYMENT TERMINATION LETTERS for Spanish-speaking and English-speaking jurisdictions.

The document is a formal letter from the employer to the employee terminating the work relationship. It must cite the country-correct legal grounds for termination and state the effective date, notice period (if any), and pending compensation.

RULE 1 — ACCURACY OVER FLUENCY: Only cite specific article numbers when highly confident they exist and apply to the termination ground. When unsure, use general formulation.

RULE 2 — COUNTRY-CORRECT TERMINATION GROUNDS:
  - México: Art. 47 LFT (rescisión sin responsabilidad para el patrón), Art. 51 (rescisión por el trabajador). Use "rescisión" (formal term).
  - Colombia: Art. 62 CST (justas causas). Art. 64 CST (despido sin justa causa con indemnización).
  - Chile: Art. 159 (mutual agreement, expiration, etc.) and Art. 160 (cause attributable to worker) of Código del Trabajo.
  - Argentina: Arts. 242 (with just cause), 245 (severance) of LCT.
  - Perú: Arts. 22-25 de la Ley de Productividad y Competitividad Laboral (LPCL).
  - España: Arts. 49-56 Estatuto de los Trabajadores (despido objetivo, disciplinario, colectivo).

RULE 3 — NOTICE PERIOD (preaviso) varies by country:
  - México: no general notice required (rescisión inmediata permitted), but written notice required Art. 47 last paragraph.
  - Colombia: 15 días for contracts without just cause (Art. 6 Ley 50/1990).
  - Chile: 30 días or pay in lieu (Art. 162).
  - Argentina: 15 días to 2 meses según antigüedad (Art. 231 LCT).
  - España: 15 días para despido objetivo (Art. 53 ET).

RULE 4 — STRUCTURE: produce 4–6 sections (not 10 — it's a letter, not a contract):
  1. Encabezado: place, date, addressee with full ID
  2. Notificación formal de la terminación con la causal legal específica
  3. Hechos que dan lugar a la causal (factual basis)
  4. Fecha efectiva de la terminación
  5. Liquidación de prestaciones (referenciar separadamente al finiquito)
  6. Firma del empleador

RULE 5 — USE PLACEHOLDERS for missing data (e.g., [Causal Específica], [Fecha de los Hechos]). Never invent facts.

LANGUAGE: output entirely in the language requested.

OUTPUT JSON SCHEMA (same as employment_contract — clausulas = sections of the letter):
{
  "titulo": "string (e.g. 'CARTA DE TERMINACIÓN DE LA RELACIÓN LABORAL')",
  "subtitulo": "string (reference to applicable law e.g. 'Conforme al Art. 47 LFT')",
  "intro": "string (place, date, addressee identification, 'Por medio de la presente...')",
  "clausulas": [{"titulo": "string", "contenido": "string"}],
  "closing": "string (formal closing, 'Sin otro particular...')",
  "partes": [{"rol": "EL EMPLEADOR / NOTIFICA", "nombre": "string"}, {"rol": "EL TRABAJADOR / RECIBE", "nombre": "string"}]
}"""


_SETTLEMENT_PROMPT = """You generate STRUCTURED EMPLOYMENT SETTLEMENT DOCUMENTS (finiquito / liquidación de prestaciones) for Spanish-speaking jurisdictions.

This document calculates final payments owed to the employee at termination, lists each component, and includes release language ("acknowledgment of full payment").

RULE 1 — ACCURACY: Cite calculation rules per the country's labor law. If unsure of specific articles, use general formulation.

RULE 2 — COUNTRY-SPECIFIC COMPONENTS (this is the most variable doc by country):

  México (finiquito):
    - Salario pendiente (días trabajados no pagados)
    - Vacaciones proporcionales (Art. 76-79 LFT)
    - Prima vacacional 25% (Art. 80 LFT)
    - Aguinaldo proporcional (15 días/año mín., Art. 87 LFT)
    - Prima de antigüedad (12 días/año, Art. 162 LFT — solo si aplica)
    - Si rescisión injustificada: indemnización 3 meses + 20 días/año + salarios caídos
    Total = SUM of components.

  Colombia (liquidación de prestaciones sociales):
    - Salarios pendientes
    - Cesantías acumuladas (1 mes/año Art. 249 CST) + intereses 12% anual
    - Prima de servicios proporcional (Art. 306 CST)
    - Vacaciones proporcionales (Art. 186 CST)
    - Si despido sin justa causa: indemnización Art. 64 CST (variable según contrato y antigüedad)
    Total = SUM. Add line for "aportes a EPS/AFP a cargo del empleador".

  Argentina (liquidación final):
    - Sueldos pendientes
    - SAC proporcional (Art. 121 LCT)
    - Vacaciones proporcionales (Art. 156 LCT)
    - Indemnización por antigüedad Art. 245 LCT (1 sueldo por año con topes)
    - Indemnización sustitutiva del preaviso (Art. 232 LCT)
    - Integración mes de despido (Art. 233 LCT)

  Chile (finiquito):
    - Indemnización por años de servicio Art. 163 (1 mes por año, tope 11 años)
    - Indemnización sustitutiva del aviso previo (1 mes)
    - Feriado proporcional
    - Sueldos pendientes y gratificación legal proporcional

  Perú (liquidación de beneficios sociales):
    - CTS proporcional, gratificaciones proporcionales julio/diciembre
    - Vacaciones truncas
    - Indemnización Art. 38 LPCL si despido arbitrario

  España (liquidación / saldo y finiquito):
    - Salario del mes en curso
    - Paga extra proporcional
    - Vacaciones no disfrutadas
    - Indemnización según tipo de despido (Art. 53 ET para objetivo, Art. 56 para improcedente)

RULE 3 — STRUCTURE: produce 5–7 sections:
  1. Encabezado: identificación de partes, fecha de terminación, motivo legal
  2. Detalle de componentes (TABLE-like content with line items + amounts using user data or placeholders)
  3. Total liquidación
  4. Compensaciones especiales si aplican (indemnización por despido, etc.)
  5. Acuerdo de finiquito: el trabajador declara recibir a satisfacción y nada más reclamar
  6. Cláusula de confidencialidad post-terminación (si aplica)
  7. Firma de ambas partes

RULE 4 — USE PLACEHOLDERS for amounts not provided. NEVER invent specific peso/COP/CLP amounts.

OUTPUT JSON SCHEMA (same as employment_contract):
{
  "titulo": "string (e.g. 'LIQUIDACIÓN DE PRESTACIONES SOCIALES Y FINIQUITO')",
  "subtitulo": "string (e.g. 'Conforme al Código Sustantivo del Trabajo')",
  "intro": "string",
  "clausulas": [{"titulo": "string", "contenido": "string"}],
  "closing": "string (release language: 'El trabajador declara haber recibido a su entera satisfacción...')",
  "partes": [{"rol": "EL EMPLEADOR", "nombre": "string"}, {"rol": "EL TRABAJADOR", "nombre": "string"}]
}"""


_COMMERCIAL_LEASE_PROMPT = """You generate STRUCTURED COMMERCIAL LEASE AGREEMENTS (contrato de arrendamiento comercial) for Spanish-speaking and English-speaking jurisdictions.

This is a lease of commercial real estate, distinct from residential leases.

RULE 1 — ACCURACY: Cite country-correct civil/commercial code articles only when confident. Otherwise use general formulation.

RULE 2 — COUNTRY-SPECIFIC LEGAL FRAMEWORK:
  - México: Código Civil Federal Arts. 2398+ (Ley de Arrendamiento aplicable según entidad federativa). IVA 16% en arrendamiento comercial (LIVA).
  - Colombia: Decreto-Ley 444 de 1996 (locales comerciales). Ley 820/2003 (urbano). IVA 19%.
  - Chile: Ley 18.101 (general urbano), Código Civil. IVA exento típicamente.
  - Argentina: Código Civil y Comercial Arts. 1187+ (locación). Ajuste por inflación según ICL.
  - España: Ley de Arrendamientos Urbanos (LAU) Ley 29/1994, sección de uso distinto al de vivienda. IVA 21%.

RULE 3 — REQUIRED CLAUSES (8–10):
  1. PRIMERA — IDENTIFICACIÓN DEL INMUEBLE: dirección, superficie m², uso permitido
  2. SEGUNDA — DESTINO Y USO COMERCIAL: actividad específica permitida
  3. TERCERA — PLAZO: vigencia, prórrogas, renovación automática
  4. CUARTA — RENTA: monto, frecuencia, forma de pago, moneda
  5. QUINTA — AJUSTE/ACTUALIZACIÓN: criterio de ajuste (IPC, índice oficial del país)
  6. SEXTA — DEPÓSITO/GARANTÍA: monto, condiciones de devolución
  7. SÉPTIMA — GASTOS Y SERVICIOS: a cargo de cada parte (luz, agua, impuesto predial, IVA)
  8. OCTAVA — OBLIGACIONES DEL ARRENDADOR Y ARRENDATARIO
  9. NOVENA — TERMINACIÓN ANTICIPADA: penalidades, causales
  10. DÉCIMA — JURISDICCIÓN

RULE 4 — TERMINOLOGY:
  - "arrendador" (landlord) — universal en español
  - "arrendatario" (tenant) — universal en español
  - English: "Lessor" / "Lessee" or "Landlord" / "Tenant"
  - "inmueble" or "local comercial" — universal

RULE 5 — USE PLACEHOLDERS for missing data ([Dirección Completa del Inmueble], [Monto de Renta]).

OUTPUT JSON SCHEMA (same as employment_contract):
{
  "titulo": "string (e.g. 'CONTRATO DE ARRENDAMIENTO COMERCIAL')",
  "subtitulo": "string",
  "intro": "string",
  "clausulas": [{"titulo": "string", "contenido": "string"}],
  "closing": "string",
  "partes": [{"rol": "EL ARRENDADOR", "nombre": "string"}, {"rol": "EL ARRENDATARIO", "nombre": "string"}]
}"""


_WORK_RULES_PROMPT = """You generate STRUCTURED INTERNAL WORK REGULATIONS (reglamento interno de trabajo / RIT) for Spanish-speaking jurisdictions.

This is a regulatory document the employer publishes to govern workplace conduct, schedules, discipline, and safety procedures.

RULE 1 — ACCURACY: Cite country-specific labor code articles only when confident.

RULE 2 — COUNTRY-SPECIFIC LEGAL FRAMEWORK AND OBLIGATION:
  - México: Art. 422-425 LFT. Obligatorio para empresas. Debe registrarse ante la JCA o Tribunal Laboral.
  - Colombia: Capítulo II del CST (Arts. 104-125). Obligatorio para empresas con ≥5 empleados (comercial) o ≥10 (industrial). Aprobación del Ministerio del Trabajo.
  - Chile: Art. 153-157 Código del Trabajo. Obligatorio si ≥10 trabajadores. Depósito en Inspección del Trabajo.
  - Argentina: no obligatorio por ley general, pero usado por convención. Régimen disciplinario en LCT Arts. 67-69.
  - Perú: Ley de Productividad y Competitividad Laboral, obligatorio si ≥100 trabajadores. Aprobación Ministerio de Trabajo.

RULE 3 — REQUIRED CONTENTS (8–12 capítulos):
  1. Capítulo I — Objeto y Ámbito de Aplicación
  2. Capítulo II — Horarios de Trabajo y Jornada
  3. Capítulo III — Días de Descanso y Vacaciones
  4. Capítulo IV — Lugar y Tiempo de Pago de Salarios
  5. Capítulo V — Permisos y Licencias
  6. Capítulo VI — Obligaciones y Prohibiciones de los Trabajadores
  7. Capítulo VII — Higiene y Seguridad en el Trabajo (referenciar normas del país: NOM en México, Resolución 2400 en Colombia, etc.)
  8. Capítulo VIII — Procedimiento de Quejas y Reclamaciones
  9. Capítulo IX — Régimen Disciplinario y Sanciones (gradación: amonestación verbal, escrita, suspensión, despido)
  10. Capítulo X — Disposiciones Finales y Vigencia

RULE 4 — LENGTH: This document is longer than a contract. Each capítulo should have 3–8 sentences of substantive content.

RULE 5 — USE PLACEHOLDERS for company-specific details.

OUTPUT JSON SCHEMA (same as employment_contract — clausulas array contains the capítulos):
{
  "titulo": "string (e.g. 'REGLAMENTO INTERNO DE TRABAJO')",
  "subtitulo": "string (reference to applicable law)",
  "intro": "string",
  "clausulas": [{"titulo": "string (e.g. 'CAPÍTULO I — OBJETO')", "contenido": "string"}],
  "closing": "string (vigencia y aprobación)",
  "partes": [{"rol": "EL EMPLEADOR", "nombre": "string"}]
}"""


_POWER_OF_ATTORNEY_PROMPT = """You generate STRUCTURED POWER OF ATTORNEY documents (poder notarial) for Spanish-speaking and English-speaking jurisdictions.

A power of attorney grants legal authority from one party (otorgante / grantor) to another (apoderado / attorney-in-fact) to act in specific or general matters.

RULE 1 — ACCURACY: Cite country-specific civil code articles only when confident. Powers of attorney are heavily formalized — terminology matters.

RULE 2 — COUNTRY-SPECIFIC FRAMEWORK:
  - México: Código Civil Federal Arts. 2546-2604. Tipos: general (para pleitos y cobranzas / actos de administración / actos de dominio Art. 2554) y especial. Requiere ratificación ante notario para actos de dominio.
  - Colombia: Código Civil Arts. 2142-2199. Poder general vs especial. Requiere escritura pública para actos que requieren forma solemne.
  - Argentina: Código Civil y Comercial Arts. 362-381 (representación), Arts. 1319-1335 (mandato).
  - Chile: Código Civil Arts. 2116-2173 (mandato).
  - Perú: Código Civil Arts. 145-167 (representación), Arts. 1790+ (mandato). Inscripción en registros para ciertos actos.
  - España: Código Civil Arts. 1709-1739. Poder notarial obligatorio para actos de disposición patrimonial.

RULE 3 — TYPE OF POWER (the user should specify; ask via placeholder if missing):
  - Poder general (broad authority for most legal acts)
  - Poder especial (specific transaction or scope)
  - Poder judicial (for litigation only)
  - Poder administrativo (for government procedures only)

RULE 4 — REQUIRED CLAUSES (6–9):
  1. PRIMERA — IDENTIFICACIÓN DEL OTORGANTE (full data including ID, address)
  2. SEGUNDA — IDENTIFICACIÓN DEL APODERADO (full data)
  3. TERCERA — TIPO DE PODER OTORGADO (general / especial / judicial / etc.)
  4. CUARTA — FACULTADES ESPECÍFICAS OTORGADAS (detallar; si es general, citar el artículo del código civil que lo define en cada país)
  5. QUINTA — LIMITACIONES (si aplican)
  6. SEXTA — VIGENCIA Y REVOCACIÓN
  7. SÉPTIMA — SUSTITUCIÓN (¿puede el apoderado delegar?)
  8. OCTAVA — JURISDICCIÓN Y RATIFICACIÓN NOTARIAL si aplica
  9. NOVENA — ACEPTACIÓN DEL APODERADO

RULE 5 — TERMINOLOGY:
  - "otorgante" / "poderdante" — universal
  - "apoderado" / "mandatario" — universal
  - English: "Principal" / "Grantor" → "Attorney-in-fact" / "Agent"

RULE 6 — IMPORTANT WARNING in closing: in most jurisdictions, an effective power of attorney for acts of disposition requires NOTARIAL CERTIFICATION. The closing should note this.

OUTPUT JSON SCHEMA:
{
  "titulo": "string (e.g. 'PODER ESPECIAL' or 'PODER GENERAL PARA PLEITOS Y COBRANZAS')",
  "subtitulo": "string",
  "intro": "string",
  "clausulas": [{"titulo": "string", "contenido": "string"}],
  "closing": "string (incluir advertencia sobre ratificación notarial si aplica)",
  "partes": [{"rol": "EL OTORGANTE", "nombre": "string"}, {"rol": "EL APODERADO", "nombre": "string"}]
}"""


SYSTEM_PROMPTS = {
    "employment_contract": _EMPLOYMENT_CONTRACT_PROMPT,
    # Phase 2
    "termination_letter": _TERMINATION_LETTER_PROMPT,
    "settlement": _SETTLEMENT_PROMPT,
    "commercial_lease": _COMMERCIAL_LEASE_PROMPT,
    "work_rules": _WORK_RULES_PROMPT,
    "power_of_attorney": _POWER_OF_ATTORNEY_PROMPT,
}


# ══════════════════════════════════════════════════════════════════════════════
# CACHÉ
# ══════════════════════════════════════════════════════════════════════════════
def _load_cache():
    global _cache, _cache_loaded
    with _cache_lock:
        if _cache_loaded:
            return
        _cache_loaded = True
        try:
            if os.path.exists(CACHE_PATH):
                with open(CACHE_PATH, "r", encoding="utf-8") as f:
                    _cache = json.load(f)
                log.info("Caché legal cargado: %d entradas desde %s", len(_cache), CACHE_PATH)
        except Exception as e:
            log.warning("No se pudo cargar el caché legal: %s", e)
            _cache = {}


def _save_cache():
    with _cache_lock:
        try:
            os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
            with open(CACHE_PATH, "w", encoding="utf-8") as f:
                json.dump(_cache, f, ensure_ascii=False, indent=2, sort_keys=True)
        except Exception as e:
            log.warning("No se pudo guardar el caché legal: %s", e)


# Parámetros que definen un documento único — cambios en estos invalidan caché.
# Datos puramente cosméticos (representante legal, etc.) NO entran en la key.
_CACHE_RELEVANT_KEYS = {
    "employment_contract": [
        "tipo_contrato", "puesto", "salario", "fecha_inicio",
        "departamento", "tipo_jornada", "dias_laborales",
        "patron_nombre", "trabajador_nombre",
    ],
    # Phase 2
    "termination_letter": [
        "causal", "fecha_terminacion", "patron_nombre",
        "trabajador_nombre", "puesto", "antiguedad",
    ],
    "settlement": [
        "fecha_inicio", "fecha_terminacion", "salario",
        "patron_nombre", "trabajador_nombre", "tipo_terminacion",
    ],
    "commercial_lease": [
        "direccion_inmueble", "monto_renta", "plazo",
        "uso_destinado", "arrendador_nombre", "arrendatario_nombre",
    ],
    "work_rules": [
        "empresa_nombre", "industria", "tamano_empresa", "ubicacion",
    ],
    "power_of_attorney": [
        "tipo_poder", "facultades", "otorgante_nombre", "apoderado_nombre",
        "vigencia",
    ],
}


def _make_cache_key(doc_type: str, country: str, lang: str, datos: dict) -> str:
    """Hash estable de los parámetros que afectan el contenido del doc."""
    keys = _CACHE_RELEVANT_KEYS.get(doc_type, [])
    relevant = {k: str(datos.get(k, "")).strip().lower() for k in keys}
    sig = json.dumps(relevant, sort_keys=True, ensure_ascii=False)
    h = hashlib.sha256(sig.encode("utf-8")).hexdigest()[:16]
    return f"{doc_type}|{country.lower()}|{lang}|{h}"


# ══════════════════════════════════════════════════════════════════════════════
# CONSTRUCCIÓN DEL MENSAJE USER PARA EL LLM
# ══════════════════════════════════════════════════════════════════════════════
def _build_user_message(doc_type: str, country: str, lang: str, datos: dict) -> str:
    """Construye el mensaje USER que se envía al LLM con los parámetros."""
    if doc_type == "employment_contract":
        return _msg_employment_contract(country, lang, datos)
    if doc_type == "termination_letter":
        return _msg_termination_letter(country, lang, datos)
    if doc_type == "settlement":
        return _msg_settlement(country, lang, datos)
    if doc_type == "commercial_lease":
        return _msg_commercial_lease(country, lang, datos)
    if doc_type == "work_rules":
        return _msg_work_rules(country, lang, datos)
    if doc_type == "power_of_attorney":
        return _msg_power_of_attorney(country, lang, datos)
    raise ValueError(f"No user-message builder for doc_type: {doc_type}")


def _g(datos, key, default="[no especificado]"):
    """Helper común: lee un dato del usuario o devuelve placeholder."""
    v = datos.get(key, "")
    s = str(v).strip() if v else ""
    return s if s else default


def _msg_termination_letter(country: str, lang: str, datos: dict) -> str:
    g = lambda k, d="[no especificado]": _g(datos, k, d)
    return "\n".join([
        "Generate a STRUCTURED EMPLOYMENT TERMINATION LETTER.",
        "",
        f"COUNTRY: {country}",
        f"LANGUAGE: {lang}",
        "",
        "EMPLOYER (notifies):",
        f"  - Name: {g('patron_nombre', '[Nombre del Empleador]')}",
        f"  - Tax ID: {g('patron_id', '[ID Fiscal]')}",
        f"  - Address: {g('patron_domicilio', '[Domicilio]')}",
        f"  - Legal representative: {g('patron_representante', '[Representante Legal]')}",
        "",
        "EMPLOYEE (receives notice):",
        f"  - Name: {g('trabajador_nombre', '[Nombre del Trabajador]')}",
        f"  - ID: {g('trabajador_id', '[ID]')}",
        f"  - Position held: {g('puesto', '[Puesto]')}",
        f"  - Start date of employment: {g('fecha_inicio', '[Fecha de Inicio]')}",
        f"  - Seniority: {g('antiguedad', '[Antigüedad]')}",
        "",
        "TERMINATION DETAILS:",
        f"  - Type: {g('tipo_terminacion', 'con causa justificada')}",
        f"  - Legal ground / cause: {g('causal', '[Causal Específica]')}",
        f"  - Effective termination date: {g('fecha_terminacion', '[Fecha Efectiva]')}",
        f"  - Notice given (días): {g('preaviso_dias', 'según legislación del país')}",
        "",
        "Generate the COMPLETE JSON. Use country-correct legal grounds and citations.",
    ])


def _msg_settlement(country: str, lang: str, datos: dict) -> str:
    g = lambda k, d="[no especificado]": _g(datos, k, d)
    return "\n".join([
        "Generate a STRUCTURED EMPLOYMENT SETTLEMENT (finiquito/liquidación).",
        "",
        f"COUNTRY: {country}",
        f"LANGUAGE: {lang}",
        "",
        "PARTIES:",
        f"  - Employer: {g('patron_nombre', '[Nombre del Empleador]')}",
        f"  - Employer ID: {g('patron_id', '[ID Empleador]')}",
        f"  - Employee: {g('trabajador_nombre', '[Nombre del Trabajador]')}",
        f"  - Employee ID: {g('trabajador_id', '[ID Trabajador]')}",
        "",
        "EMPLOYMENT HISTORY:",
        f"  - Position: {g('puesto', '[Puesto]')}",
        f"  - Start date: {g('fecha_inicio', '[Fecha de Inicio]')}",
        f"  - Termination date: {g('fecha_terminacion', '[Fecha de Terminación]')}",
        f"  - Last salary: {g('salario', '[Último Salario]')}",
        f"  - Type of termination: {g('tipo_terminacion', '[Tipo: voluntaria / con causa / sin causa]')}",
        "",
        "Generate the COMPLETE JSON with country-correct calculation components.",
        "List each line item with its legal basis (article citation when confident).",
        "Use placeholders for specific amounts not provided. NEVER invent amounts.",
    ])


def _msg_commercial_lease(country: str, lang: str, datos: dict) -> str:
    g = lambda k, d="[no especificado]": _g(datos, k, d)
    return "\n".join([
        "Generate a STRUCTURED COMMERCIAL LEASE AGREEMENT.",
        "",
        f"COUNTRY: {country}",
        f"LANGUAGE: {lang}",
        "",
        "LANDLORD (arrendador):",
        f"  - Name: {g('arrendador_nombre', '[Nombre del Arrendador]')}",
        f"  - ID: {g('arrendador_id', '[ID]')}",
        f"  - Address: {g('arrendador_domicilio', '[Domicilio]')}",
        "",
        "TENANT (arrendatario):",
        f"  - Name: {g('arrendatario_nombre', '[Nombre del Arrendatario]')}",
        f"  - ID: {g('arrendatario_id', '[ID]')}",
        "",
        "PROPERTY:",
        f"  - Address: {g('direccion_inmueble', '[Dirección Completa]')}",
        f"  - Surface area: {g('superficie', '[Superficie m²]')}",
        f"  - Permitted use: {g('uso_destinado', '[Uso Comercial Permitido]')}",
        "",
        "LEASE TERMS:",
        f"  - Monthly rent: {g('monto_renta', '[Monto de Renta]')}",
        f"  - Duration: {g('plazo', '[Plazo]')}",
        f"  - Start date: {g('fecha_inicio', '[Fecha de Inicio]')}",
        f"  - Deposit: {g('deposito', '[Monto de Depósito]')}",
        f"  - Annual adjustment criterion: {g('ajuste_renta', 'IPC/índice oficial del país')}",
        "",
        "Generate COMPLETE JSON. Use country-correct civil/commercial code framework.",
    ])


def _msg_work_rules(country: str, lang: str, datos: dict) -> str:
    g = lambda k, d="[no especificado]": _g(datos, k, d)
    return "\n".join([
        "Generate STRUCTURED INTERNAL WORK REGULATIONS (Reglamento Interno de Trabajo).",
        "",
        f"COUNTRY: {country}",
        f"LANGUAGE: {lang}",
        "",
        "COMPANY:",
        f"  - Name: {g('empresa_nombre', '[Nombre de la Empresa]')}",
        f"  - Tax ID: {g('empresa_id', '[ID Fiscal]')}",
        f"  - Address: {g('ubicacion', '[Ubicación]')}",
        f"  - Industry/Activity: {g('industria', '[Actividad/Giro]')}",
        f"  - Company size: {g('tamano_empresa', '[N° de empleados]')}",
        f"  - Work schedule: {g('horario_general', '[Horario]')}",
        "",
        "Generate COMPLETE JSON with 8-12 capítulos covering all required topics.",
        "Reference country-specific labor law and registration requirements.",
        "Each capítulo should be substantive (3-8 sentences).",
    ])


def _msg_power_of_attorney(country: str, lang: str, datos: dict) -> str:
    g = lambda k, d="[no especificado]": _g(datos, k, d)
    return "\n".join([
        "Generate a STRUCTURED POWER OF ATTORNEY document.",
        "",
        f"COUNTRY: {country}",
        f"LANGUAGE: {lang}",
        "",
        "GRANTOR (otorgante):",
        f"  - Name: {g('otorgante_nombre', '[Nombre del Otorgante]')}",
        f"  - ID: {g('otorgante_id', '[ID del Otorgante]')}",
        f"  - Address: {g('otorgante_domicilio', '[Domicilio del Otorgante]')}",
        "",
        "ATTORNEY-IN-FACT (apoderado):",
        f"  - Name: {g('apoderado_nombre', '[Nombre del Apoderado]')}",
        f"  - ID: {g('apoderado_id', '[ID del Apoderado]')}",
        f"  - Address: {g('apoderado_domicilio', '[Domicilio del Apoderado]')}",
        "",
        "POWER DETAILS:",
        f"  - Type: {g('tipo_poder', '[General / Especial / Judicial / Administrativo]')}",
        f"  - Specific powers granted: {g('facultades', '[Facultades Específicas]')}",
        f"  - Duration / validity: {g('vigencia', '[Vigencia]')}",
        f"  - Substitution allowed: {g('puede_sustituir', 'No salvo autorización expresa')}",
        "",
        "Generate COMPLETE JSON. Use country-correct civil code framework.",
        "Include warning about notarial certification in closing.",
    ])


def _msg_employment_contract(country: str, lang: str, datos: dict) -> str:
    """Mensaje user para contrato de trabajo."""
    def g(key, default="[no especificado]"):
        v = datos.get(key, "")
        s = str(v).strip() if v else ""
        return s if s else default

    lines = [
        "Generate a STRUCTURED EMPLOYMENT CONTRACT for the parameters below.",
        "",
        f"COUNTRY: {country}",
        f"LANGUAGE: {lang}",
        f"CONTRACT TYPE: {g('tipo_contrato', 'indefinido')}",
        "",
        "EMPLOYER:",
        f"  - Name: {g('patron_nombre', '[Nombre del Empleador]')}",
        f"  - Tax ID: {g('patron_id', '[ID Fiscal del Empleador]')}",
        f"  - Address: {g('patron_domicilio', '[Domicilio del Empleador]')}",
        f"  - Legal representative: {g('patron_representante', '[Representante Legal]')}",
        f"  - Representative position: {g('patron_cargo', '[Cargo del Representante]')}",
        "",
        "EMPLOYEE:",
        f"  - Name: {g('trabajador_nombre', '[Nombre del Trabajador]')}",
        f"  - ID: {g('trabajador_id', '[ID del Trabajador]')}",
        f"  - Address: {g('trabajador_domicilio', '[Domicilio del Trabajador]')}",
        "",
        "POSITION:",
        f"  - Job title: {g('puesto', '[Puesto]')}",
        f"  - Department: {g('departamento', '[Departamento]')}",
        f"  - Start date: {g('fecha_inicio', '[Fecha de Inicio]')}",
        f"  - Work schedule: {g('horario', 'standard working hours')}",
        f"  - Days worked: {g('dias_laborales', 'Monday to Friday')}",
        f"  - Type of shift: {g('tipo_jornada', 'standard daytime')}",
        "",
        "COMPENSATION:",
        f"  - Monthly salary: {g('salario', '[Salario Mensual]')}",
        f"  - Payment frequency: {g('forma_pago', 'monthly bank transfer')}",
        "",
        "Generate the COMPLETE JSON structure following the schema in the system prompt.",
        "Each clause must be substantive (3-6 sentences typical), country-correct,",
        "and use the language specified above. Cite specific articles only when",
        "confident; use general formulation otherwise.",
    ]
    return "\n".join(lines)


# ══════════════════════════════════════════════════════════════════════════════
# LLAMADA AL LLM
# ══════════════════════════════════════════════════════════════════════════════
def _call_llm(doc_type: str, country: str, lang: str, datos: dict) -> dict:
    """Llama a Claude con prompt estructurado y devuelve JSON parseado."""
    system = SYSTEM_PROMPTS.get(doc_type)
    if not system:
        raise ValueError(f"No system prompt registered for: {doc_type}")

    user_msg = _build_user_message(doc_type, country, lang, datos)

    resp = _client.messages.create(
        model=MODEL,
        max_tokens=8000,
        system=system,
        messages=[{"role": "user", "content": user_msg}],
    )

    raw = resp.content[0].text.strip() if resp.content else ""
    # Defensa contra fences que el LLM podría agregar
    raw = re.sub(r'^```(?:json)?\s*', '', raw)
    raw = re.sub(r'\s*```$', '', raw)

    try:
        structure = json.loads(raw)
    except json.JSONDecodeError as e:
        log.error("LLM returned invalid JSON for %s: %s; raw[:500]=%r",
                  doc_type, e, raw[:500])
        raise

    # Validación mínima de estructura
    required = ["titulo", "clausulas"]
    missing = [k for k in required if k not in structure]
    if missing:
        raise ValueError(f"LLM output missing required keys {missing}: {structure}")
    if not isinstance(structure["clausulas"], list) or not structure["clausulas"]:
        raise ValueError(f"LLM output has empty/invalid 'clausulas': {structure}")

    return structure


# ══════════════════════════════════════════════════════════════════════════════
# API PÚBLICA — GENERACIÓN Y RENDERING
# ══════════════════════════════════════════════════════════════════════════════
def generate_legal_structure(doc_type: str, country: str, lang: str, datos: dict) -> dict:
    """Genera la estructura del documento legal con caché por (tipo, país, lang, params)."""
    _load_cache()
    key = _make_cache_key(doc_type, country, lang, datos)

    with _cache_lock:
        cached = _cache.get(key)
    if cached:
        log.info("Caché legal HIT: %s", key)
        return cached

    log.info("Caché legal MISS, llamando LLM: %s", key)
    structure = _call_llm(doc_type, country, lang, datos)

    with _cache_lock:
        _cache[key] = structure
    _save_cache()
    return structure


def render_legal_doc(doc, structure: dict, datos: dict, cfg: dict):
    """Renderiza la estructura JSON en el .docx usando los helpers de word_controller.

    Lazy import para evitar ciclos: word_controller importa llm_legal_generator
    dentro del dispatch de generate_word, y aquí volvemos a importar word_controller
    para los helpers. El import lazy asegura que el orden de carga no sea problema.
    """
    from office.word_controller import (
        _heading, _para, _divider, _signature_block,
        WD_ALIGN_PARAGRAPH, C_NAVY, C_BLUE, C_GRAY,
    )

    # Título principal
    titulo = structure.get("titulo", "[Document Title]")
    _heading(doc, titulo, level=1,
             align=WD_ALIGN_PARAGRAPH.CENTER, color=C_NAVY)

    # Subtítulo (referencia a ley aplicable)
    subtitulo = structure.get("subtitulo", "")
    if subtitulo:
        _para(doc, subtitulo,
              align=WD_ALIGN_PARAGRAPH.CENTER, color=C_GRAY, size=10)

    _divider(doc)

    # Párrafo introductorio
    intro = structure.get("intro", "")
    if intro:
        _para(doc, intro, size=10.5)
        doc.add_paragraph()

    # Cláusulas
    for clause in structure.get("clausulas", []):
        titulo_c = clause.get("titulo", "")
        contenido = clause.get("contenido", "")
        if titulo_c:
            _heading(doc, titulo_c, level=2, color=C_BLUE)
        if contenido:
            _para(doc, contenido, size=10.5)
            doc.add_paragraph()

    # Cierre
    closing = structure.get("closing", "")
    if closing:
        _para(doc, closing, size=10.5)
        doc.add_paragraph()

    # Bloque de firmas
    partes = structure.get("partes", [])
    if partes:
        sig_data = [
            {"title": p.get("rol", "")[:60],
             "name": p.get("nombre", "[Nombre y Firma]")}
            for p in partes
        ]
        _signature_block(doc, sig_data)


def build_via_llm(doc, doc_type: str, instruccion: str, secciones: list,
                  detalles: str, cfg: dict | None, datos: dict | None):
    """Builder LLM-driven con firma compatible con los builders hardcoded.

    Esta es la función que llama el dispatcher de generate_word. Si algo
    falla (API caída, JSON inválido, doc_type no soportado), propaga la
    excepción para que el caller use el builder hardcoded como fallback.

    Args coinciden con la firma de los _build_* hardcoded:
        doc:          docx.Document creado por _doc_base
        doc_type:     template_id (e.g. "employment_contract")
        instruccion:  texto libre del usuario (no usado en Phase 1)
        secciones:    lista de secciones del usuario (no usada en Phase 1)
        detalles:     detalles específicos del usuario (no usado en Phase 1)
        cfg:          LEGAL_CONFIG del país
        datos:        dict con datos del usuario (parte_a_*, salario, etc.)
    """
    cfg = cfg or {}
    datos = datos or {}
    country = cfg.get("pais", "")
    # v15.2.1 — Si el país del usuario no tiene LEGAL_CONFIG explícito,
    # cfg cae a GENERIC con pais="[País]". Para que el LLM reciba el nombre
    # real del país, usamos el input crudo del usuario como fallback.
    # Esto hace funcionar el flujo LLM-driven para cualquier país LATAM
    # aunque no haya entrada en LEGAL_CONFIG.
    if not country or country == "[País]":
        country = (datos.get("pais") or "").strip() or "[País]"
    lang = cfg.get("idioma", "es")

    structure = generate_legal_structure(doc_type, country, lang, datos)
    render_legal_doc(doc, structure, datos, cfg)


# ══════════════════════════════════════════════════════════════════════════════
# UTILIDADES OPCIONALES
# ══════════════════════════════════════════════════════════════════════════════
def cache_stats() -> dict:
    """Estadísticas del caché — útil para admin/monitoreo."""
    _load_cache()
    by_doctype: dict[str, int] = {}
    for k in _cache:
        dtype = k.split("|", 1)[0]
        by_doctype[dtype] = by_doctype.get(dtype, 0) + 1
    return {
        "total_entries": len(_cache),
        "by_doctype": by_doctype,
        "path": CACHE_PATH,
        "size_bytes": os.path.getsize(CACHE_PATH) if os.path.exists(CACHE_PATH) else 0,
    }


def clear_cache():
    """Vacía el caché en memoria y borra el archivo. Útil para tests."""
    global _cache, _cache_loaded
    with _cache_lock:
        _cache = {}
        _cache_loaded = True
        try:
            if os.path.exists(CACHE_PATH):
                os.remove(CACHE_PATH)
        except Exception as e:
            log.warning("No se pudo borrar el caché legal: %s", e)


def is_llm_driven(template_id: str) -> bool:
    """True si el template_id se genera vía LLM en lugar de builder hardcoded."""
    return template_id in SYSTEM_PROMPTS
