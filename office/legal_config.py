"""
WEP AI — Configuración legal y fiscal compartida (v15.5).

Módulo extraído de word_controller.py para que Word *y* Excel consuman la
misma fuente de datos por país. Resuelve el problema histórico de que Excel
asumiera México implícitamente mientras Word ya soportaba 20 países.

Estructura:
─────────────
  • LEGAL_CONFIG       — diccionario por código de país (MX, CO, PE, …)
                         con datos legales (leyes, artículos, instituciones).
                         Usado por word_controller para los builders legales.
  • COUNTRY_ALIASES    — mapa de variaciones de nombre ("méxico", "cdmx",
                         "mexico", "mx") → código ISO ("MX").
  • EXCEL_FISCAL       — overlay numérico por país para Excel:
                         tasa IVA decimal, formato de moneda, lista de
                         deducciones de nómina con sus tasas por defecto,
                         retención de honorarios para freelance.
  • _get_legal_config  — resolver código país → cfg, con fallback a GENERIC.
  • _detect_country    — detectar país desde gen_data o texto libre.
  • get_excel_fiscal   — devolver EXCEL_FISCAL para un cfg dado.

Cada builder de Excel que reciba `cfg` lo pasa por `get_excel_fiscal(cfg)`
para obtener tasas y etiquetas locales en lugar de hardcodear 16% IVA o
IMSS/INFONAVIT.

NOTA SOBRE LAS TASAS DE EXCEL_FISCAL:
Las tasas de deducción de nómina e impuestos a honorarios son valores de
referencia comunes en cada país a mayo 2026. NO son asesoría fiscal ni
laboral. Los builders generan estas tasas en celdas EDITABLES de una hoja
"Parámetros" — el usuario final ajusta según su caso. Cuando el régimen
fiscal de un país es muy variado (escalas progresivas, retenciones
condicionales), la tasa por defecto es 0% y la nota explica que debe
consultarse la tabla oficial.
"""

import os


# ─────────────────────────────────────────────────────────────────────────────
LEGAL_CONFIG = {
    # ── LATINOAMÉRICA ─────────────────────────────────────────────────────────
    "MX": {
        "pais": "México", "ciudad_default": "Ciudad de México",
        "idioma": "es", "moneda": "MXN", "simbolo": "$",
        "id_empresa": "RFC", "id_persona": "RFC / CURP",
        "ley_laboral": "Ley Federal del Trabajo (LFT)",
        "ley_servicios": "Código Civil Federal",
        "ley_compraventa": "Código Civil Federal",
        "art_objeto": "Art. 2248 CCF",
        "seguridad_social": "IMSS", "tasa_ss_patron": "9%",
        "pension": "INFONAVIT / AFORE", "tasa_pension": "5%",
        "iva": "16%", "iva_label": "IVA", "autoridad_fiscal": "SAT",
        "vacaciones": "según tabla Art. 76 LFT (6-24 días/año)",
        "aguinaldo": "15 días de salario (Art. 87 LFT)",
        "jornada": "8 hrs/día, 48 hrs/semana (Art. 61 LFT)",
        "rescision_patron": "Art. 47 LFT", "rescision_trabajador": "Art. 51 LFT",
        "tribunales": "Tribunales del Poder Judicial de la Federación",
        "gratificaciones": "No aplica (se sustituye por aguinaldo)",
        "cts": "No aplica",
        "retencion_honorarios": "10% ISR (Art. 106 LISR)",
        "registro_empresa": "SAT / Registro Público de Comercio",
    },
    "CO": {
        "pais": "Colombia", "ciudad_default": "Bogotá, D.C.",
        "idioma": "es", "moneda": "COP", "simbolo": "$",
        "id_empresa": "NIT", "id_persona": "CC (Cédula de Ciudadanía)",
        "ley_laboral": "Código Sustantivo del Trabajo (CST), modificado por la Ley 2466 de 2025 (Reforma Laboral)",
        "ley_servicios": "Código Civil y Código de Comercio colombiano",
        "ley_compraventa": "Código de Comercio de Colombia",
        "art_objeto": "Art. 905 C. de Co.",
        "seguridad_social": "EPS (Entidad Promotora de Salud)", "tasa_ss_patron": "12.5%",
        "pension": "AFP (Administradora de Fondos de Pensiones)", "tasa_pension": "12%",
        "iva": "19%", "iva_label": "IVA", "autoridad_fiscal": "DIAN",
        "vacaciones": "15 días hábiles por año (Art. 186 CST)",
        "aguinaldo": "Prima de servicios: 30 días de salario/año (Art. 306 CST)",
        # CORREGIDO v16.9: la reducción de jornada (Ley 2101/2021) es gradual.
        # 44 h/semana desde 15-jul-2025; 42 h/semana desde 15-jul-2026.
        "jornada": "Hasta 8 hrs/día; 44 hrs/semana hasta el 14-jul-2026 y 42 hrs/semana desde el 15-jul-2026 (Art. 161 CST, mod. Ley 2101/2021 y Ley 2466/2025)",
        "rescision_patron": "Art. 62 CST (justa causa) + procedimiento disciplinario previo (Ley 2466/2025)",
        "rescision_trabajador": "Art. 62 CST",
        "tribunales": "Jueces Laborales del Circuito de [ciudad]",
        "gratificaciones": "Prima de navidad: 15 días de salario (Art. 306 CST)",
        "cts": "Cesantías: 1 mes de salario/año (Art. 249 CST)",
        "retencion_honorarios": "Retención en la fuente según tabla DIAN",
        "registro_empresa": "Cámara de Comercio",
        # ── v16.7 — Datos clave de la Reforma Laboral (Ley 2466 de 2025) ──────
        # Fuente normativa: Ley 2466 de 2025 (D.O. 53.160), vigente desde el
        # 25 de junio de 2025. Estos valores alimentan el contrato de trabajo
        # y el módulo de diagnóstico (agent/labor_reform_diagnostic.py).
        "reforma_laboral": "Ley 2466 de 2025",
        "reforma_vigencia": "25 de junio de 2025",
        "contrato_regla_general": "término indefinido (Art. 5 Ley 2466/2025; el indefinido es la regla, las demás modalidades son la excepción y deben justificarse)",
        "fijo_duracion_max": "4 años, incluidas prórrogas (Ley 2466/2025)",
        # CORREGIDO v16.9: la jornada nocturna a las 7:00 p.m. aplica a TODAS las
        # empresas sin excepción por tamaño (Art. 11 Ley 2466/2025, vigente desde
        # el 25-dic-2025). No existe excepción para microempresas en este punto.
        "jornada_nocturna_inicio": "7:00 p.m. a 6:00 a.m. para TODAS las empresas, sin excepción por tamaño (Art. 11 Ley 2466/2025, vigente desde el 25 de diciembre de 2025)",
        # CORREGIDO v16.9: el recargo dominical es escalonado con fechas exactas
        # (parágrafo transitorio, Art. 14 Ley 2466/2025), no un genérico "hacia 100%".
        "recargo_dominical": "Escalonado: 80% (1-jul-2025 a 30-jun-2026), 90% (1-jul-2026 a 30-jun-2027), 100% (desde 1-jul-2027). Antes era 75% (Art. 14 Ley 2466/2025)",
        # CORREGIDO v16.9: el procedimiento disciplinario tiene un régimen
        # simplificado para empresas con menos de 10 empleados.
        "procedimiento_disciplinario": "Procedimiento mínimo previo al despido/sanción con justa causa: notificación escrita de los hechos, garantía de defensa, investigación, análisis de descargos, decisión fundamentada, notificación formal y recurso (Ley 2466/2025). Para empresas con MENOS de 10 empleados el proceso es simplificado: basta escuchar previamente al trabajador.",
        "contrato_escrito": "El contrato a término fijo y por obra/labor debe constar por escrito; en su defecto se entiende a término indefinido (Ley 2466/2025)",
        "cuota_aprendices": "1 aprendiz por cada 20 trabajadores en empresas con 15+ empleados; monetización de 1,5 SMMLV por aprendiz no contratado (Ley 2466/2025)",
    },
    "PE": {
        "pais": "Perú", "ciudad_default": "Lima",
        "idioma": "es", "moneda": "PEN", "simbolo": "S/",
        "id_empresa": "RUC", "id_persona": "DNI",
        "ley_laboral": "Decreto Legislativo N° 728 (sector privado)",
        "ley_servicios": "Código Civil peruano (Art. 1764 y ss.)",
        "ley_compraventa": "Código Civil peruano (Art. 1529 y ss.)",
        "art_objeto": "Art. 1529 CC peruano",
        "seguridad_social": "EsSalud", "tasa_ss_patron": "9%",
        "pension": "ONP o AFP (a elección del trabajador)", "tasa_pension": "13% ONP",
        "iva": "18%", "iva_label": "IGV", "autoridad_fiscal": "SUNAT",
        "vacaciones": "30 días calendario por año (DL 713)",
        "aguinaldo": "No aplica",
        "jornada": "8 hrs/día, 48 hrs/semana (Art. 25 DL 854)",
        "rescision_patron": "Art. 25 DL 728 (falta grave)",
        "rescision_trabajador": "Art. 35 DL 728",
        "tribunales": "Jueces Especializados de Trabajo de [ciudad]",
        "gratificaciones": "2 gratificaciones/año: julio y diciembre (1 sueldo c/u — Ley 27735)",
        "cts": "½ sueldo cada mayo y noviembre (Ley 25920)",
        "retencion_honorarios": "8% retención 4ta categoría (Art. 74 LISR)",
        "registro_empresa": "SUNARP / SUNAT",
    },
    "AR": {
        "pais": "Argentina", "ciudad_default": "Ciudad Autónoma de Buenos Aires",
        "idioma": "es", "moneda": "ARS", "simbolo": "$",
        "id_empresa": "CUIT", "id_persona": "CUIL / DNI",
        "ley_laboral": "Ley de Contrato de Trabajo (LCT — Ley 20.744)",
        "ley_servicios": "Código Civil y Comercial de la Nación (CCyCN)",
        "ley_compraventa": "CCyCN Art. 1123 y ss.",
        "art_objeto": "Art. 1123 CCyCN",
        "seguridad_social": "ANSES / Obra Social", "tasa_ss_patron": "24%",
        "pension": "SIPA (ANSES)", "tasa_pension": "11%",
        "iva": "21%", "iva_label": "IVA", "autoridad_fiscal": "AFIP",
        "vacaciones": "14-35 días según antigüedad (Art. 150 LCT)",
        "aguinaldo": "SAC: ½ mejor remuneración, en junio y diciembre (Art. 121 LCT)",
        "jornada": "8 hrs/día, 48 hrs/semana (Ley 11.544)",
        "rescision_patron": "Art. 242 LCT (justa causa)",
        "rescision_trabajador": "Art. 242 LCT",
        "tribunales": "Tribunales del Trabajo de [ciudad]",
        "gratificaciones": "SAC (aguinaldo) — 2 cuotas/año",
        "cts": "No aplica (se reemplaza por indemnización LCT)",
        "retencion_honorarios": "Retención según escala AFIP (Ganancias 4ta categoría)",
        "registro_empresa": "IGJ (CABA) / Registro Público de Comercio",
    },
    "CL": {
        "pais": "Chile", "ciudad_default": "Santiago de Chile",
        "idioma": "es", "moneda": "CLP", "simbolo": "$",
        "id_empresa": "RUT (empresa)", "id_persona": "RUT",
        "ley_laboral": "Código del Trabajo de Chile",
        "ley_servicios": "Código Civil de Chile (Art. 2006 y ss.)",
        "ley_compraventa": "Código Civil de Chile (Art. 1793 y ss.)",
        "art_objeto": "Art. 1793 CC chileno",
        "seguridad_social": "Isapre o Fonasa", "tasa_ss_patron": "No aplica (trabajador paga)",
        "pension": "AFP (Administradora de Fondos de Pensiones)", "tasa_pension": "~10% trabajador",
        "iva": "19%", "iva_label": "IVA", "autoridad_fiscal": "SII",
        "vacaciones": "15 días hábiles por año (Art. 67 CT)",
        "aguinaldo": "No aplica",
        "jornada": "40 hrs/semana (Ley 21.561 desde 2023)",
        "rescision_patron": "Art. 160 CT (causal) / Art. 161 CT (necesidades empresa)",
        "rescision_trabajador": "Art. 171 CT",
        "tribunales": "Juzgado de Letras del Trabajo de [ciudad]",
        "gratificaciones": "25% de remuneraciones anuales o 4.75 IMM (Art. 47 CT)",
        "cts": "Indemnización por años de servicio (Art. 163 CT)",
        "retencion_honorarios": "Retención 10.75% (honorarios en boleta)",
        "registro_empresa": "Registro de Empresas y Sociedades / SII",
    },
    "ES": {
        "pais": "España", "ciudad_default": "Madrid",
        "idioma": "es", "moneda": "EUR", "simbolo": "€",
        "id_empresa": "CIF", "id_persona": "NIF / NIE",
        "ley_laboral": "Estatuto de los Trabajadores (ET — RDL 2/2015)",
        "ley_servicios": "Código Civil español (Art. 1544 y ss.)",
        "ley_compraventa": "Código de Comercio español (Art. 325 y ss.)",
        "art_objeto": "Art. 1254 CC español",
        "seguridad_social": "Seguridad Social (TGSS)", "tasa_ss_patron": "23.6%",
        "pension": "Seguridad Social — Régimen General", "tasa_pension": "4.7% trabajador",
        "iva": "21%", "iva_label": "IVA", "autoridad_fiscal": "AEAT (Agencia Tributaria)",
        "vacaciones": "30 días naturales por año (Art. 38 ET)",
        "aguinaldo": "Pagas extraordinarias: mínimo 2/año (Art. 31 ET)",
        "jornada": "40 hrs/semana, 9 hrs/día máx. (Art. 34 ET)",
        "rescision_patron": "Art. 54 ET (causas disciplinarias)",
        "rescision_trabajador": "Art. 50 ET (resolución causal)",
        "tribunales": "Juzgados de lo Social de [ciudad]",
        "gratificaciones": "Pagas extras (junio y diciembre mínimo)",
        "cts": "No aplica (liquidación al finalizar contrato)",
        "retencion_honorarios": "IRPF 15% (7% primeros 2 años de actividad)",
        "registro_empresa": "Registro Mercantil",
    },
    "US": {
        "pais": "United States", "ciudad_default": "[City, State]",
        "idioma": "en", "moneda": "USD", "simbolo": "$",
        "id_empresa": "EIN (Employer Identification Number)",
        "id_persona": "SSN (Social Security Number)",
        "ley_laboral": "Fair Labor Standards Act (FLSA) + applicable State Law",
        "ley_servicios": "Applicable State Contract Law",
        "ley_compraventa": "Uniform Commercial Code (UCC)",
        "art_objeto": "UCC § 2-201",
        "seguridad_social": "Social Security & Medicare", "tasa_ss_patron": "7.65% (FICA)",
        "pension": "401(k) — employer discretionary", "tasa_pension": "Varies",
        "iva": "Varies by state", "iva_label": "Sales Tax", "autoridad_fiscal": "IRS",
        "vacaciones": "No federal mandate — per company policy",
        "aguinaldo": "No legal requirement — discretionary",
        "jornada": "40 hrs/week (overtime over 40 hrs at 1.5x — FLSA)",
        "rescision_patron": "At-Will Employment (most states) with applicable notice",
        "rescision_trabajador": "At-Will — resignation with [notice period] notice",
        "tribunales": "Courts of [State] — [County] jurisdiction",
        "gratificaciones": "Discretionary bonus — not legally mandated",
        "cts": "Not applicable (severance per company policy)",
        "retencion_honorarios": "Form 1099-NEC for payments over $600",
        "registro_empresa": "Secretary of State — [State]",
    },
    "BO": {
        "pais": "Bolivia", "ciudad_default": "La Paz",
        "idioma": "es", "moneda": "BOB", "simbolo": "Bs",
        "id_empresa": "NIT", "id_persona": "CI (Carnet de Identidad)",
        "ley_laboral": "Ley General del Trabajo (LGT)",
        "ley_servicios": "Código Civil boliviano + Código de Comercio",
        "ley_compraventa": "Código de Comercio boliviano",
        "art_objeto": "Art. 822 Código de Comercio boliviano (compraventa mercantil)",
        "seguridad_social": "Caja Nacional de Salud (CNS)",
        "tasa_ss_patron": "16.71% aportes patronales (incluye CNS, AFP-Riesgo Profesional)",
        "pension": "Sistema Integral de Pensiones (Gestora Pública)",
        "tasa_pension": "12.71% aporte laboral",
        "iva": "13%", "iva_label": "IVA",
        "autoridad_fiscal": "Servicio de Impuestos Nacionales (SIN)",
        "vacaciones": "15 a 30 días por año según antigüedad (Art. 33 LGT)",
        "aguinaldo": "1 sueldo mensual antes del 20 de diciembre (Ley General del Aguinaldo)",
        "jornada": "8 hrs/día, 48 hrs/semana (Art. 46 LGT)",
        "rescision_patron": "Art. 16 LGT (causales de despido con justa causa)",
        "rescision_trabajador": "Art. 2 LGT (retiro voluntario)",
        "tribunales": "Juzgados de Trabajo y Seguridad Social de [Ciudad]",
        "gratificaciones": "2do aguinaldo 'Esfuerzo por Bolivia' si crecimiento PIB >4.5% (DS 1802)",
        "cts": "Indemnización por tiempo de servicios: 1 sueldo por año (Art. 13 LGT)",
        "retencion_honorarios": "RC-IVA 13% y IUE según corresponda",
        "registro_empresa": "FUNDEMPRESA y SIN",
    },
    "EC": {
        "pais": "Ecuador", "ciudad_default": "Quito",
        "idioma": "es", "moneda": "USD", "simbolo": "$",
        "id_empresa": "RUC (13 dígitos)", "id_persona": "Cédula",
        "ley_laboral": "Código del Trabajo del Ecuador",
        "ley_servicios": "Código Civil ecuatoriano (Art. 1454 y ss.)",
        "ley_compraventa": "Código de Comercio ecuatoriano",
        "art_objeto": "Art. 1454 CC ecuatoriano",
        "seguridad_social": "Instituto Ecuatoriano de Seguridad Social (IESS)",
        "tasa_ss_patron": "12.15% aporte patronal",
        "pension": "IESS (Seguro General Obligatorio)",
        "tasa_pension": "9.45% aporte personal",
        "iva": "15%", "iva_label": "IVA",
        "autoridad_fiscal": "Servicio de Rentas Internas (SRI)",
        "vacaciones": "15 días por año (Art. 69 Código del Trabajo)",
        "aguinaldo": "Décimo Tercer Sueldo (Navidad) + Décimo Cuarto Sueldo (Escolar)",
        "jornada": "8 hrs/día, 40 hrs/semana (Art. 47 Código del Trabajo)",
        "rescision_patron": "Art. 172 Código del Trabajo (causales de despido)",
        "rescision_trabajador": "Art. 173 Código del Trabajo (visto bueno)",
        "tribunales": "Juzgados de Trabajo de [Ciudad]",
        "gratificaciones": "Décimo tercero (1 sueldo / dic) + Décimo cuarto (1 SBU / región)",
        "cts": "Fondos de reserva (1 sueldo por año, Art. 196 Código del Trabajo)",
        "retencion_honorarios": "Retención IR según tabla SRI (8-10% honorarios profesionales)",
        "registro_empresa": "Superintendencia de Compañías y SRI",
    },
    "UY": {
        "pais": "Uruguay", "ciudad_default": "Montevideo",
        "idioma": "es", "moneda": "UYU", "simbolo": "$U",
        "id_empresa": "RUT (DGI)", "id_persona": "CI (Cédula de Identidad)",
        "ley_laboral": "Conjunto de leyes laborales uruguayas (no hay Código único)",
        "ley_servicios": "Código Civil uruguayo + Ley 12.030",
        "ley_compraventa": "Código de Comercio uruguayo",
        "art_objeto": "Art. 1247 CC uruguayo",
        "seguridad_social": "Banco de Previsión Social (BPS)",
        "tasa_ss_patron": "12.625% aportes patronales (jubilatorio + FRL)",
        "pension": "BPS + AFAP (sistema mixto)",
        "tasa_pension": "15% aporte personal (jubilatorio)",
        "iva": "22%", "iva_label": "IVA",
        "autoridad_fiscal": "Dirección General Impositiva (DGI)",
        "vacaciones": "20 días por año + Salario Vacacional (Ley 14.522)",
        "aguinaldo": "Sueldo Anual Complementario — medio sueldo en junio + medio en diciembre",
        "jornada": "8 hrs/día, 48 hrs/semana (Ley 5.350)",
        "rescision_patron": "Indemnización por despido — 1 mes por año (Ley 10.489)",
        "rescision_trabajador": "Renuncia con preaviso según antigüedad",
        "tribunales": "Juzgados Letrados de Trabajo de [Ciudad]",
        "gratificaciones": "Salario Vacacional + Aguinaldo (SAC)",
        "cts": "Indemnización por despido (Ley 10.489)",
        "retencion_honorarios": "IRPF según tabla DGI",
        "registro_empresa": "DGI + BPS + Registro de Comercio",
    },
    "PY": {
        "pais": "Paraguay", "ciudad_default": "Asunción",
        "idioma": "es", "moneda": "PYG", "simbolo": "₲",
        "id_empresa": "RUC", "id_persona": "CI (Cédula de Identidad)",
        "ley_laboral": "Código del Trabajo (Ley 213/93)",
        "ley_servicios": "Código Civil paraguayo (Ley 1183/85)",
        "ley_compraventa": "Código de Comercio paraguayo",
        "art_objeto": "Art. 1453 CC paraguayo",
        "seguridad_social": "Instituto de Previsión Social (IPS)",
        "tasa_ss_patron": "16.5% aporte patronal",
        "pension": "IPS (jubilatorio incluido)",
        "tasa_pension": "9% aporte personal",
        "iva": "10%", "iva_label": "IVA",
        "autoridad_fiscal": "Subsecretaría de Estado de Tributación (SET)",
        "vacaciones": "12 a 30 días por año según antigüedad (Art. 218 CT)",
        "aguinaldo": "1 sueldo antes del 31 de diciembre (Decreto-Ley 18.071/50)",
        "jornada": "8 hrs/día, 48 hrs/semana (Art. 194 Código del Trabajo)",
        "rescision_patron": "Art. 81 Código del Trabajo (causales)",
        "rescision_trabajador": "Art. 84 Código del Trabajo",
        "tribunales": "Juzgados de Primera Instancia del Trabajo de [Ciudad]",
        "gratificaciones": "Aguinaldo anual obligatorio",
        "cts": "Indemnización por despido (Art. 91 Código del Trabajo)",
        "retencion_honorarios": "IRP según tabla SET",
        "registro_empresa": "SET + IPS + Registro Público de Comercio",
    },
    "VE": {
        "pais": "Venezuela", "ciudad_default": "Caracas",
        "idioma": "es", "moneda": "VED", "simbolo": "Bs.",
        "id_empresa": "RIF (Registro de Información Fiscal)",
        "id_persona": "Cédula de Identidad",
        "ley_laboral": "Ley Orgánica del Trabajo, los Trabajadores y las Trabajadoras (LOTTT)",
        "ley_servicios": "Código Civil venezolano",
        "ley_compraventa": "Código de Comercio venezolano",
        "art_objeto": "Art. 1474 CC venezolano",
        "seguridad_social": "Instituto Venezolano de los Seguros Sociales (IVSS)",
        "tasa_ss_patron": "Variable según riesgo de la empresa (LOTTT y RGSS)",
        "pension": "IVSS",
        "tasa_pension": "4% aporte personal",
        "iva": "16%", "iva_label": "IVA",
        "autoridad_fiscal": "SENIAT (Servicio Nacional Integrado de Administración Aduanera y Tributaria)",
        "vacaciones": "15 días iniciales + 1 día adicional por año (Art. 190 LOTTT)",
        "aguinaldo": "Utilidades anuales (mín. 30 días, Art. 132 LOTTT) + Bono de Fin de Año",
        "jornada": "8 hrs/día, 40 hrs/semana (Art. 173 LOTTT)",
        "rescision_patron": "Art. 79 LOTTT (causales justificadas)",
        "rescision_trabajador": "Art. 80 LOTTT",
        "tribunales": "Tribunales del Trabajo del Circuito Judicial de [Ciudad]",
        "gratificaciones": "Utilidades + Bono de Fin de Año",
        "cts": "Prestaciones Sociales (Art. 142 LOTTT)",
        "retencion_honorarios": "Retención ISLR según tabla SENIAT",
        "registro_empresa": "Registro Mercantil + SENIAT",
    },
    "PA": {
        "pais": "Panamá", "ciudad_default": "Ciudad de Panamá",
        "idioma": "es", "moneda": "PAB / USD", "simbolo": "B/.",
        "id_empresa": "RUC", "id_persona": "Cédula",
        "ley_laboral": "Código de Trabajo de Panamá",
        "ley_servicios": "Código Civil panameño",
        "ley_compraventa": "Código de Comercio panameño",
        "art_objeto": "Art. 1215 CC panameño",
        "seguridad_social": "Caja de Seguro Social (CSS)",
        "tasa_ss_patron": "12.25% aporte patronal",
        "pension": "CSS (Sistema Mixto)",
        "tasa_pension": "9.75% aporte personal",
        "iva": "7%", "iva_label": "ITBMS",
        "autoridad_fiscal": "Dirección General de Ingresos (DGI)",
        "vacaciones": "30 días por cada 11 meses trabajados (Art. 54 CT)",
        "aguinaldo": "Décimo Tercer Mes (XIII Mes) en 3 cuotas: abril, agosto y diciembre",
        "jornada": "8 hrs/día, 48 hrs/semana (Art. 33 CT)",
        "rescision_patron": "Art. 213 CT (causales justificadas)",
        "rescision_trabajador": "Art. 214 CT",
        "tribunales": "Juntas de Conciliación y Decisión + Juzgados Seccionales de Trabajo",
        "gratificaciones": "XIII Mes obligatorio",
        "cts": "Prima de antigüedad (Art. 224 CT — 1 semana por año)",
        "retencion_honorarios": "Retención ISR según tabla DGI",
        "registro_empresa": "Registro Público de Panamá + DGI",
    },
    "GT": {
        "pais": "Guatemala", "ciudad_default": "Ciudad de Guatemala",
        "idioma": "es", "moneda": "GTQ", "simbolo": "Q",
        "id_empresa": "NIT", "id_persona": "DPI (Documento Personal de Identificación)",
        "ley_laboral": "Código de Trabajo (Decreto 1441)",
        "ley_servicios": "Código Civil guatemalteco (Decreto-Ley 106)",
        "ley_compraventa": "Código de Comercio guatemalteco",
        "art_objeto": "Art. 1517 CC guatemalteco",
        "seguridad_social": "Instituto Guatemalteco de Seguridad Social (IGSS)",
        "tasa_ss_patron": "12.67% aporte patronal",
        "pension": "IGSS (Programa de Invalidez, Vejez y Sobrevivencia)",
        "tasa_pension": "4.83% aporte personal",
        "iva": "12%", "iva_label": "IVA",
        "autoridad_fiscal": "Superintendencia de Administración Tributaria (SAT)",
        "vacaciones": "15 días por año (Art. 130 Código de Trabajo)",
        "aguinaldo": "Bono 14 (julio, Decreto 42-92) + Aguinaldo (diciembre, Art. 102 Const.)",
        "jornada": "8 hrs/día, 44 hrs/semana (Art. 116 Código de Trabajo)",
        "rescision_patron": "Art. 76-78 Código de Trabajo (causales)",
        "rescision_trabajador": "Art. 79 Código de Trabajo",
        "tribunales": "Juzgados de Trabajo y Previsión Social",
        "gratificaciones": "Bono 14 + Aguinaldo (2 sueldos extras/año)",
        "cts": "Indemnización por tiempo de servicio (Art. 82 Código de Trabajo)",
        "retencion_honorarios": "Retención ISR según régimen SAT",
        "registro_empresa": "Registro Mercantil + SAT",
    },
    "HN": {
        "pais": "Honduras", "ciudad_default": "Tegucigalpa",
        "idioma": "es", "moneda": "HNL", "simbolo": "L",
        "id_empresa": "RTN (Registro Tributario Nacional)",
        "id_persona": "DNI / Tarjeta de Identidad",
        "ley_laboral": "Código del Trabajo de Honduras",
        "ley_servicios": "Código Civil hondureño",
        "ley_compraventa": "Código de Comercio hondureño",
        "art_objeto": "Art. 1605 CC hondureño",
        "seguridad_social": "Instituto Hondureño de Seguridad Social (IHSS)",
        "tasa_ss_patron": "Variable según régimen (EM ~5%, IVM ~3.5%)",
        "pension": "IHSS + RAP (Régimen de Aportaciones Privadas)",
        "tasa_pension": "2.5% aporte personal (IVM)",
        "iva": "15%", "iva_label": "ISV (Impuesto Sobre Ventas)",
        "autoridad_fiscal": "Servicio de Administración de Rentas (SAR)",
        "vacaciones": "10 a 20 días por año según antigüedad (Art. 346 CT)",
        "aguinaldo": "Décimo Tercer Mes (Navidad) + Décimo Cuarto Mes (julio)",
        "jornada": "8 hrs/día, 44 hrs/semana (Art. 322 Código del Trabajo)",
        "rescision_patron": "Art. 112 Código del Trabajo (causas justas)",
        "rescision_trabajador": "Art. 114 Código del Trabajo",
        "tribunales": "Juzgados de Trabajo de [Ciudad]",
        "gratificaciones": "Décimo Tercer + Décimo Cuarto Mes (2 sueldos extras/año)",
        "cts": "Auxilio de cesantía (Art. 120 Código del Trabajo)",
        "retencion_honorarios": "Retención ISR según tabla SAR",
        "registro_empresa": "Cámara de Comercio + SAR",
    },
    "NI": {
        "pais": "Nicaragua", "ciudad_default": "Managua",
        "idioma": "es", "moneda": "NIO", "simbolo": "C$",
        "id_empresa": "RUC", "id_persona": "Cédula",
        "ley_laboral": "Código del Trabajo (Ley 185)",
        "ley_servicios": "Código Civil nicaragüense",
        "ley_compraventa": "Código de Comercio nicaragüense",
        "art_objeto": "Art. 2436 CC nicaragüense",
        "seguridad_social": "Instituto Nicaragüense de Seguridad Social (INSS)",
        "tasa_ss_patron": "22.5% aporte patronal (régimen integral)",
        "pension": "INSS",
        "tasa_pension": "7% aporte personal",
        "iva": "15%", "iva_label": "IVA",
        "autoridad_fiscal": "Dirección General de Ingresos (DGI)",
        "vacaciones": "30 días por año (Art. 76 Código del Trabajo)",
        "aguinaldo": "Décimo Tercer Mes antes del 10 de diciembre (Art. 93 CT)",
        "jornada": "8 hrs/día, 48 hrs/semana (Art. 51 Código del Trabajo)",
        "rescision_patron": "Art. 48 Código del Trabajo (causales justas)",
        "rescision_trabajador": "Art. 41 Código del Trabajo",
        "tribunales": "Juzgados del Trabajo y Seguridad Social",
        "gratificaciones": "Décimo Tercer Mes obligatorio",
        "cts": "Indemnización por antigüedad (Art. 45 Código del Trabajo)",
        "retencion_honorarios": "Retención IR según tabla DGI",
        "registro_empresa": "DGI + INSS + Registro Público Mercantil",
    },
    "SV": {
        "pais": "El Salvador", "ciudad_default": "San Salvador",
        "idioma": "es", "moneda": "USD", "simbolo": "$",
        "id_empresa": "NIT", "id_persona": "DUI (Documento Único de Identidad)",
        "ley_laboral": "Código de Trabajo de El Salvador",
        "ley_servicios": "Código Civil salvadoreño",
        "ley_compraventa": "Código de Comercio salvadoreño",
        "art_objeto": "Art. 1597 CC salvadoreño",
        "seguridad_social": "Instituto Salvadoreño del Seguro Social (ISSS)",
        "tasa_ss_patron": "7.5% ISSS + 7.75% AFP aporte patronal",
        "pension": "AFP (Sistema de Ahorro para Pensiones)",
        "tasa_pension": "7.25% aporte personal AFP",
        "iva": "13%", "iva_label": "IVA",
        "autoridad_fiscal": "Ministerio de Hacienda — DGII",
        "vacaciones": "15 días por año con prima vacacional 30% (Art. 177 CT)",
        "aguinaldo": "Después del 12 de diciembre, según antigüedad (15-21 días)",
        "jornada": "8 hrs/día, 44 hrs/semana (Art. 161 Código de Trabajo)",
        "rescision_patron": "Art. 50 Código de Trabajo (causales)",
        "rescision_trabajador": "Art. 53 Código de Trabajo",
        "tribunales": "Juzgados de lo Laboral de [Ciudad]",
        "gratificaciones": "Aguinaldo navideño obligatorio",
        "cts": "Indemnización por despido injusto (Art. 58 Código de Trabajo)",
        "retencion_honorarios": "Retención ISR según tabla DGII",
        "registro_empresa": "Centro Nacional de Registros (CNR) + Ministerio de Hacienda",
    },
    "CR": {
        "pais": "Costa Rica", "ciudad_default": "San José",
        "idioma": "es", "moneda": "CRC", "simbolo": "₡",
        "id_empresa": "Cédula Jurídica", "id_persona": "Cédula",
        "ley_laboral": "Código de Trabajo de Costa Rica",
        "ley_servicios": "Código Civil costarricense",
        "ley_compraventa": "Código de Comercio costarricense",
        "art_objeto": "Art. 1049 CC costarricense",
        "seguridad_social": "Caja Costarricense de Seguro Social (CCSS)",
        "tasa_ss_patron": "26.67% cargas sociales patronales (incl. CCSS, INS, BPDC, FODESAF)",
        "pension": "CCSS (IVM) + ROP (Régimen Obligatorio de Pensiones Complementarias)",
        "tasa_pension": "10.67% aporte personal",
        "iva": "13%", "iva_label": "IVA",
        "autoridad_fiscal": "Ministerio de Hacienda — Dirección General de Tributación",
        "vacaciones": "2 semanas por cada 50 semanas trabajadas (Art. 153 CT)",
        "aguinaldo": "1 mes de salario en diciembre (Ley de Aguinaldo)",
        "jornada": "8 hrs/día, 48 hrs/semana (Art. 136 Código de Trabajo)",
        "rescision_patron": "Art. 81 Código de Trabajo (causales justificadas)",
        "rescision_trabajador": "Art. 83 Código de Trabajo",
        "tribunales": "Juzgados de Trabajo de [Ciudad]",
        "gratificaciones": "Aguinaldo obligatorio (mes 13)",
        "cts": "Cesantía (Art. 29 CT — 1 mes por año hasta máx. 8 años)",
        "retencion_honorarios": "Retención ISR según régimen de Hacienda",
        "registro_empresa": "Registro Nacional + Ministerio de Hacienda",
    },
    "CU": {
        "pais": "Cuba", "ciudad_default": "La Habana",
        "idioma": "es", "moneda": "CUP", "simbolo": "$",
        "id_empresa": "NIT (Cuba)", "id_persona": "Carné de Identidad",
        "ley_laboral": "Código de Trabajo (Ley 116/2013)",
        "ley_servicios": "Código Civil cubano (Ley 59)",
        "ley_compraventa": "Código de Comercio cubano",
        "art_objeto": "Art. 334 CC cubano",
        "seguridad_social": "Sistema de Seguridad Social cubano (Ley 105/2008)",
        "tasa_ss_patron": "Variable según régimen estatal/no estatal",
        "pension": "Sistema de Seguridad Social cubano",
        "tasa_pension": "Variable según régimen",
        "iva": "No aplica IVA tradicional — sistema tributario especial",
        "iva_label": "Sistema tributario cubano",
        "autoridad_fiscal": "Oficina Nacional de Administración Tributaria (ONAT)",
        "vacaciones": "30 días por año (Art. 99 Código de Trabajo)",
        "aguinaldo": "Conforme a normativa laboral vigente",
        "jornada": "8 hrs/día, 44 hrs/semana (Art. 89 Código de Trabajo)",
        "rescision_patron": "Conforme a Ley 116/2013",
        "rescision_trabajador": "Conforme a Ley 116/2013",
        "tribunales": "Tribunales Municipales Populares — Sala de lo Económico",
        "gratificaciones": "Conforme a normativa vigente",
        "cts": "Conforme a normativa vigente",
        "retencion_honorarios": "Conforme a normativa ONAT",
        "registro_empresa": "Registro Estatal + ONAT",
    },
    "DO": {
        "pais": "República Dominicana", "ciudad_default": "Santo Domingo",
        "idioma": "es", "moneda": "DOP", "simbolo": "RD$",
        "id_empresa": "RNC (Registro Nacional del Contribuyente)",
        "id_persona": "Cédula",
        "ley_laboral": "Código de Trabajo (Ley 16-92)",
        "ley_servicios": "Código Civil dominicano",
        "ley_compraventa": "Código de Comercio dominicano",
        "art_objeto": "Art. 1582 CC dominicano",
        "seguridad_social": "Tesorería de Seguridad Social (TSS) - SDSS",
        "tasa_ss_patron": "16.04% aportes patronales (SFS, SVDS, ARL)",
        "pension": "AFP (Sistema Dominicano de Seguridad Social)",
        "tasa_pension": "2.87% aporte personal",
        "iva": "18%", "iva_label": "ITBIS",
        "autoridad_fiscal": "Dirección General de Impuestos Internos (DGII)",
        "vacaciones": "14 días por año (Art. 177 Código de Trabajo)",
        "aguinaldo": "Salario de Navidad / Regalía Pascual (mes 13) en diciembre",
        "jornada": "8 hrs/día, 44 hrs/semana (Art. 147 Código de Trabajo)",
        "rescision_patron": "Art. 88 Código de Trabajo (causales de despido)",
        "rescision_trabajador": "Art. 96 Código de Trabajo (dimisión)",
        "tribunales": "Tribunales de Trabajo de [Ciudad]",
        "gratificaciones": "Regalía Pascual + bonificación por utilidades (10%)",
        "cts": "Cesantía (Art. 80 Código de Trabajo)",
        "retencion_honorarios": "Retención ISR según tabla DGII",
        "registro_empresa": "Cámara de Comercio + DGII",
    },
    "GENERIC": {
        "pais": "[País]", "ciudad_default": "[Ciudad]",
        "idioma": "es", "moneda": "[Moneda]", "simbolo": "$",
        "id_empresa": "[ID Fiscal Empresa]", "id_persona": "[ID Personal]",
        "ley_laboral": "[Ley Laboral aplicable del país]",
        "ley_servicios": "[Código Civil / Ley de Contratos aplicable]",
        "ley_compraventa": "[Ley de Compraventa aplicable]",
        "art_objeto": "[Artículo aplicable]",
        "seguridad_social": "[Entidad de Seguridad Social]", "tasa_ss_patron": "[%]",
        "pension": "[Sistema de Pensiones]", "tasa_pension": "[%]",
        "iva": "[%]", "iva_label": "IVA/IGV", "autoridad_fiscal": "[Autoridad Fiscal]",
        "vacaciones": "[N] días por año según ley aplicable",
        "aguinaldo": "[Según ley local]",
        "jornada": "[Según ley laboral local]",
        "rescision_patron": "[Artículo aplicable]",
        "rescision_trabajador": "[Artículo aplicable]",
        "tribunales": "Tribunales competentes de [Ciudad, País]",
        "gratificaciones": "[Según ley local]",
        "cts": "[Según ley local]",
        "retencion_honorarios": "[Según ley fiscal local]",
        "registro_empresa": "[Registro Público de Comercio / equivalente]",
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# Aliases — variaciones de nombres de país
# ─────────────────────────────────────────────────────────────────────────────
# Aliases — variaciones de nombres de país
COUNTRY_ALIASES = {
    "mexico":"MX","méxico":"MX","mex":"MX","mx":"MX",
    "colombia":"CO","col":"CO","co":"CO",
    "peru":"PE","perú":"PE","pe":"PE",
    "argentina":"AR","arg":"AR","ar":"AR",
    "chile":"CL","cl":"CL",
    "españa":"ES","spain":"ES","es":"ES","esp":"ES",
    "usa":"US","united states":"US","estados unidos":"US","us":"US","eeuu":"US",
    "bolivia":"BO","bo":"BO","la paz":"BO","santa cruz":"BO",
    "ecuador":"EC","ec":"EC","quito":"EC","guayaquil":"EC",
    "uruguay":"UY","uy":"UY","montevideo":"UY",
    "paraguay":"PY","py":"PY","asunción":"PY","asuncion":"PY",
    "venezuela":"VE","ve":"VE","caracas":"VE","maracaibo":"VE",
    "panama":"PA","panamá":"PA","pa":"PA",
    "guatemala":"GT","gt":"GT","gua":"GT",
    "honduras":"HN","hn":"HN","tegucigalpa":"HN",
    "nicaragua":"NI","ni":"NI","managua":"NI",
    "el salvador":"SV","salvador":"SV","sv":"SV","san salvador":"SV",
    "costa rica":"CR","cr":"CR","san josé":"CR","san jose":"CR",
    "cuba":"CU","cu":"CU","habana":"CU","la habana":"CU",
    "republica dominicana":"DO","república dominicana":"DO",
    "dominicana":"DO","do":"DO","santo domingo":"DO",
}

# ─────────────────────────────────────────────────────────────────────────────
# Funciones de detección
# ─────────────────────────────────────────────────────────────────────────────
def _apply_persistent_overlay(code: str, cfg: dict) -> dict:
    """
    v15.16.0 — Aplica el overlay persistente de cambios verificados por IA.
    Los verifiers guardan en ~/.wepai/legal_overrides.json los cambios legales
    confirmados; aquí se aplican sobre los valores base del código para que
    sobrevivan reinicios. Defensivo: si el módulo falla, devuelve el cfg base.
    """
    try:
        from agent.legal_overrides import apply_overlay
        return apply_overlay(code, cfg)
    except Exception:
        return cfg


def _get_legal_config(pais_raw: str = "") -> dict:
    """Return country-specific legal config. Falls back to Colombia (v16.0)."""
    key = COUNTRY_ALIASES.get(pais_raw.lower().strip(), "CO")
    cfg = LEGAL_CONFIG.get(key, LEGAL_CONFIG["CO"])
    return _apply_persistent_overlay(key, cfg)


def _detect_country(gen_data: dict, instruccion: str = "", detalles: str = "") -> dict:
    """Detect country from gen_data fields or text context."""
    # Priority 1: explicit field in gen_data
    pais = gen_data.get("datos", {}).get("pais", "") or gen_data.get("pais", "")
    if pais:
        return _get_legal_config(pais)
    # Priority 2: detect from text keywords
    full = f"{instruccion} {detalles}".lower()
    country_keywords = {
        "MX": ["mexico","méxico","cdmx","monterrey","guadalajara","lft","rfc","imss","sat mexico","infonavit"],
        "CO": ["colombia","bogotá","bogota","medellín","medellin","cali","cst","nit","dian","eps","colombia"],
        "PE": ["peru","perú","lima","arequipa","dl 728","sunat","ruc","essalud","sunarp","perú"],
        "AR": ["argentina","buenos aires","córdoba","rosario","lct","afip","cuit","anses"],
        "CL": ["chile","santiago","valparaíso","valparaiso","código del trabajo","sii","rut","afp chile"],
        "ES": ["españa","spain","madrid","barcelona","estatuto trabajadores","aeat","cif","nif"],
        "US": ["at-will","united states","estados unidos","eeuu","flsa","irs","ein","ssn","us startup","llc","inc.","corporation"],
        "BO": ["bolivia","la paz","santa cruz","cochabamba","sucre","lgt","ley general del trabajo","caja nacional de salud","cns","fundempresa","servicio de impuestos nacionales"],
        "EC": ["ecuador","quito","guayaquil","cuenca","sri","iess","ruc ecuador","sbu","décimo tercero","décimo cuarto"],
        "UY": ["uruguay","montevideo","punta del este","bps","afap","dgi uruguay","salario vacacional","sac uruguay"],
        "PY": ["paraguay","asunción","asuncion","ips paraguay","set paraguay","guaraní","guarani"],
        "VE": ["venezuela","caracas","maracaibo","valencia venezuela","lottt","seniat","ivss","rif","bolívar","bolivar digital","ved"],
        "PA": ["panamá","panama","ciudad de panamá","ciudad de panama","css","itbms","balboa","b/.","décimo tercer mes","caja de seguro social"],
        "GT": ["guatemala","ciudad de guatemala","quetzaltenango","igss","sat guatemala","bono 14","quetzal","dpi guatemala"],
        "HN": ["honduras","tegucigalpa","san pedro sula","ihss","sar honduras","rtn","lempira","décimo cuarto"],
        "NI": ["nicaragua","managua","león","leon","inss nicaragua","dgi nicaragua","córdoba nicaragua"],
        "SV": ["el salvador","san salvador","isss","dgii","dui","salvador","afp el salvador"],
        "CR": ["costa rica","san josé","san jose","ccss","caja costarricense","cédula jurídica","colón costarricense"],
        "CU": ["cuba","la habana","habana","santiago de cuba","onat","ley 116","cup","peso cubano"],
        "DO": ["república dominicana","republica dominicana","dominicana","santo domingo","santiago de los caballeros","dgii","rnc","itbis","regalía pascual","tss"],
    }
    for code, keywords in country_keywords.items():
        if any(kw in full for kw in keywords):
            return _apply_persistent_overlay(code, LEGAL_CONFIG[code])
    # v16.0 — ENFOQUE COLOMBIA: si no se detecta país, el default es Colombia
    # (antes era GENERIC). La arquitectura multipaís sigue intacta: si el texto
    # menciona otro país de la lista, se respeta. Solo cambia el fallback.
    return _apply_persistent_overlay("CO", LEGAL_CONFIG["CO"])



# ─────────────────────────────────────────────────────────────────────────────
# EXCEL_FISCAL — overlay numérico para builders de Excel (v15.5).
# Cada país tiene:
#   iva_rate:          tasa IVA/IGV/VAT como decimal (0.16 = 16%)
#   iva_label:         "IVA" / "IGV" / "ITBMS" / "ITBIS" / "Sales Tax" / …
#   currency_format:   formato de openpyxl para celdas de moneda
#   payroll_deductions: lista de (label, default_rate, note) para builder de nómina.
#                       Las tasas 0.00 indican "depende de tabla" — el usuario las
#                       edita en la hoja "Parámetros".
#   income_tax_freelance:        tasa default de retención sobre honorarios
#   income_tax_freelance_label:  cómo aparece esa retención en el doc
#   tax_authority:     "SAT", "DIAN", "SUNAT", "AFIP", … (para referencias)
# ─────────────────────────────────────────────────────────────────────────────
EXCEL_FISCAL = {
    "MX": {
        "iva_rate": 0.16, "iva_label": "IVA",
        "currency_format": "$#,##0.00",
        "payroll_deductions": [
            ("IMSS",      0.0287, "Cuotas obrero IMSS (enfermedad + invalidez + cesantía)"),
            ("ISR",       0.10,   "Estimación — usa tabla LISR Art. 96 vigente"),
            ("INFONAVIT", 0.05,   "Aplica solo si el trabajador tiene crédito"),
        ],
        "payroll_benefits": [
            ("Aguinaldo",     "15 días de salario mínimo por año (Art. 87 LFT) — pagadero antes del 20 de diciembre"),
            ("Prima Vacacional","25% sobre días de vacaciones (Art. 80 LFT)"),
            ("PTU",           "10% de la utilidad fiscal del año (Art. 117 LFT) — pagadero en mayo/junio"),
            ("Vacaciones",    "12 días primer año, +2 por año hasta 20 días (reforma 2023, Art. 76 LFT)"),
            ("Prima Dominical","25% adicional al salario por trabajar en domingo (Art. 71 LFT)"),
        ],
        "income_tax_freelance": 0.10,
        "income_tax_freelance_label": "ISR retención honorarios (Art. 106 LISR)",
        "tax_authority": "SAT",
    },
    "CO": {
        "iva_rate": 0.19, "iva_label": "IVA",
        "currency_format": "$#,##0",
        "payroll_deductions": [
            ("Salud EPS",         0.04, "Aporte personal a salud (4%)"),
            ("Pensión AFP",       0.04, "Aporte personal a pensión (4%)"),
            ("Retefuente",        0.00, "Según tabla DIAN — depende del salario"),
            ("Solidaridad",       0.00, "1-2% si salario > 4 SMMLV (Art. 27 Ley 100)"),
        ],
        "payroll_benefits": [
            ("Prima de Servicios",    "1 salario mensual al año (pagadero 15 jun + 20 dic) — Art. 306 CST"),
            ("Cesantías",             "1 salario mensual por año — depósito antes del 14 feb del año siguiente (Art. 249 CST)"),
            ("Intereses sobre Cesantías","12% anual sobre las cesantías — pagadero antes del 31 ene (Art. 1 Ley 52/1975)"),
            ("Vacaciones",            "15 días hábiles por año de servicio (Art. 186 CST)"),
            ("Auxilio de Transporte", "Subsidio mensual para salarios hasta 2 SMMLV — fijado por decreto anual"),
            ("Dotación",              "3 mudas de ropa/calzado al año para salarios hasta 2 SMMLV (Art. 230 CST)"),
        ],
        "income_tax_freelance": 0.11,
        "income_tax_freelance_label": "Retención en la fuente — honorarios (10-11%)",
        "tax_authority": "DIAN",
    },
    "PE": {
        "iva_rate": 0.18, "iva_label": "IGV",
        "currency_format": "S/. #,##0.00",
        "payroll_deductions": [
            ("AFP / ONP",   0.10,  "AFP ~12.7% o ONP 13% (a elección del trabajador)"),
            ("IR 5ta cat.", 0.00,  "Según escala SUNAT — depende del salario"),
            ("CTS (acumula)", 0.00, "Beneficio — 1/2 sueldo mayo + noviembre (no descuento)"),
        ],
        "payroll_benefits": [
            ("EsSalud",          "9% sobre el salario — aporte del EMPLEADOR (Ley 26790, no es deducción)"),
            ("Gratificación Julio","1 remuneración íntegra — pagadera en julio (Ley 27735)"),
            ("Gratificación Diciembre","1 remuneración íntegra — pagadera en diciembre (Ley 27735)"),
            ("CTS",              "1 remuneración por año — depositada en mayo y noviembre (TUO DL 650)"),
            ("Vacaciones",       "30 días por año de servicio (Art. 10 DL 713)"),
            ("Bonificación Extraordinaria","9% sobre la gratificación (lo que el empleador ahorra de EsSalud)"),
            ("Asignación Familiar","10% de RMV si tiene hijos menores de 18 (Ley 25129)"),
        ],
        "income_tax_freelance": 0.08,
        "income_tax_freelance_label": "Retención IR 4ta categoría (8% honorarios)",
        "tax_authority": "SUNAT",
    },
    "AR": {
        "iva_rate": 0.21, "iva_label": "IVA",
        "currency_format": "$ #,##0.00",
        "payroll_deductions": [
            ("Jubilación SIPA",       0.11,  "Aporte personal jubilatorio (11%)"),
            ("Obra Social",           0.03,  "Aporte personal obra social (3%)"),
            ("PAMI / INSSJP",         0.03,  "Aporte personal Ley 19.032"),
            ("Ganancias",             0.00,  "Según tabla AFIP — depende del salario"),
        ],
        "payroll_benefits": [
            ("Aguinaldo (SAC)",      "1 sueldo dividido en 2 cuotas — junio y diciembre (Ley 27073, Art. 121-122 LCT)"),
            ("Vacaciones",           "14 días < 5 años; 21 días 5-10; 28 días 10-20; 35 días > 20 (Art. 150 LCT)"),
            ("Asignaciones Familiares","Por hijo menor 18 (Ley 24.714) — pagadas por ANSES"),
            ("Plus por Antigüedad",  "1% del salario por año trabajado (en algunos convenios)"),
            ("ART",                  "Cobertura riesgos del trabajo — paga el EMPLEADOR (Ley 24.557, no deducción)"),
        ],
        "income_tax_freelance": 0.06,
        "income_tax_freelance_label": "Retención Ganancias 4ta categoría (RG 830)",
        "tax_authority": "AFIP",
    },
    "CL": {
        "iva_rate": 0.19, "iva_label": "IVA",
        "currency_format": "$#,##0",
        "payroll_deductions": [
            ("AFP",                   0.10,   "Aporte obligatorio AFP (10%)"),
            ("Salud Isapre/Fonasa",   0.07,   "Aporte personal salud (7% mínimo)"),
            ("AFC seguro cesantía",   0.006,  "Aporte personal cesantía (0.6%)"),
            ("Impuesto único",        0.00,   "Escala progresiva SII"),
        ],
        "payroll_benefits": [
            ("Gratificación Legal","25% de utilidades con tope 4.75 IMM anual (Art. 50 CT) — opción A o B según empresa"),
            ("Vacaciones",         "15 días hábiles por año (Art. 67 CT)"),
            ("Indemnización por Años","1 mes de salario por año trabajado (tope 11 años) en despido sin causa (Art. 163 CT)"),
            ("Bono de Escolaridad","Negociado por contrato — no es legal obligatorio"),
            ("Asignación Familiar","Si renta < tope INE — pagada por IPS"),
        ],
        "income_tax_freelance": 0.1375,
        "income_tax_freelance_label": "Retención boleta de honorarios (13.75% 2024+)",
        "tax_authority": "SII",
    },
    "ES": {
        "iva_rate": 0.21, "iva_label": "IVA",
        "currency_format": "#,##0.00 €",
        "payroll_deductions": [
            ("Seguridad Social", 0.0635, "Aporte personal régimen general (6.35%)"),
            ("IRPF",             0.00,   "Según tabla AEAT — depende del salario"),
        ],
        "income_tax_freelance": 0.15,
        "income_tax_freelance_label": "Retención IRPF (15% / 7% primeros 2 años)",
        "tax_authority": "AEAT",
    },
    "US": {
        "iva_rate": 0.00, "iva_label": "Sales Tax",
        "currency_format": "$#,##0.00",
        "payroll_deductions": [
            ("FICA Social Security", 0.062,  "6.2% withheld (employee share)"),
            ("Medicare",             0.0145, "1.45% withheld (employee share)"),
            ("Federal Income Tax",   0.00,   "Per W-4 / IRS withholding tables"),
            ("State Income Tax",     0.00,   "Varies by state"),
        ],
        "income_tax_freelance": 0.00,
        "income_tax_freelance_label": "Form 1099-NEC — contractor pays own taxes",
        "tax_authority": "IRS",
    },
    "BO": {
        "iva_rate": 0.13, "iva_label": "IVA",
        "currency_format": "Bs #,##0.00",
        "payroll_deductions": [
            ("AFP aporte personal",   0.1071, "Jubilación + comisión + riesgo (10.71%)"),
            ("Aporte Solidario",      0.005,  "0.5% adicional Ley 065"),
            ("RC-IVA",                0.13,   "13% sobre saldo después de mínimo exento"),
        ],
        "income_tax_freelance": 0.13,
        "income_tax_freelance_label": "RC-IVA / IUE según corresponda (SIN)",
        "tax_authority": "SIN",
    },
    "EC": {
        "iva_rate": 0.15, "iva_label": "IVA",
        "currency_format": "$#,##0.00",
        "payroll_deductions": [
            ("IESS aporte personal", 0.0945, "9.45% al IESS"),
            ("Retención IR",         0.00,   "Según tabla SRI — depende del salario"),
        ],
        "income_tax_freelance": 0.10,
        "income_tax_freelance_label": "Retención IR honorarios profesionales (8-10%)",
        "tax_authority": "SRI",
    },
    "UY": {
        "iva_rate": 0.22, "iva_label": "IVA",
        "currency_format": "$U #,##0.00",
        "payroll_deductions": [
            ("BPS jubilatorio",  0.15,  "15% aporte personal"),
            ("FONASA salud",     0.045, "4.5% (variable según ingreso y familia)"),
            ("FRL",              0.001, "0.1% Fondo de Reconversión Laboral"),
            ("IRPF",             0.00,  "Según tabla DGI"),
        ],
        "income_tax_freelance": 0.07,
        "income_tax_freelance_label": "IRPF servicios personales (escala 7-36%)",
        "tax_authority": "DGI",
    },
    "PY": {
        "iva_rate": 0.10, "iva_label": "IVA",
        "currency_format": "₲ #,##0",
        "payroll_deductions": [
            ("IPS aporte personal", 0.09, "9% al Instituto de Previsión Social"),
            ("IRP",                 0.00, "Según escala SET (8% o 10%)"),
        ],
        "income_tax_freelance": 0.085,
        "income_tax_freelance_label": "IRP-RSP (8% o 10% según escala SET)",
        "tax_authority": "SET",
    },
    "VE": {
        "iva_rate": 0.16, "iva_label": "IVA",
        "currency_format": "Bs. #,##0.00",
        "payroll_deductions": [
            ("IVSS",                          0.04,  "4% aporte personal Seguro Social"),
            ("Régimen Prestacional Empleo",   0.005, "0.5%"),
            ("Política Habitacional BANAVIH", 0.01,  "1%"),
            ("ISLR",                          0.00,  "Retención según tabla SENIAT"),
        ],
        "income_tax_freelance": 0.03,
        "income_tax_freelance_label": "Retención ISLR honorarios (3%)",
        "tax_authority": "SENIAT",
    },
    "PA": {
        "iva_rate": 0.07, "iva_label": "ITBMS",
        "currency_format": "B/. #,##0.00",
        "payroll_deductions": [
            ("CSS aporte personal", 0.0975, "9.75% al Seguro Social"),
            ("Seguro Educativo",    0.0125, "1.25%"),
            ("ISR",                 0.00,   "Escala 0 / 15 / 25% según tramo"),
        ],
        "income_tax_freelance": 0.07,
        "income_tax_freelance_label": "Retención ISR según tabla DGI",
        "tax_authority": "DGI Panamá",
    },
    "GT": {
        "iva_rate": 0.12, "iva_label": "IVA",
        "currency_format": "Q #,##0.00",
        "payroll_deductions": [
            ("IGSS aporte personal", 0.0483, "4.83% al IGSS"),
            ("ISR",                  0.00,   "Régimen opcional: 5-7% s/ingresos o 25% s/renta neta"),
        ],
        "income_tax_freelance": 0.05,
        "income_tax_freelance_label": "Retención ISR (5% pequeño contribuyente)",
        "tax_authority": "SAT Guatemala",
    },
    "HN": {
        "iva_rate": 0.15, "iva_label": "ISV",
        "currency_format": "L #,##0.00",
        "payroll_deductions": [
            ("IHSS EM",  0.025, "Enfermedad-Maternidad (2.5%)"),
            ("IHSS IVM", 0.025, "Invalidez-Vejez-Muerte (2.5%)"),
            ("RAP",      0.015, "Régimen de Aportaciones Privadas (1.5%)"),
            ("ISR",      0.00,  "Escala 0/15/20/25% — tabla SAR"),
        ],
        "income_tax_freelance": 0.125,
        "income_tax_freelance_label": "Retención ISR honorarios (12.5%)",
        "tax_authority": "SAR",
    },
    "NI": {
        "iva_rate": 0.15, "iva_label": "IVA",
        "currency_format": "C$ #,##0.00",
        "payroll_deductions": [
            ("INSS Laboral", 0.07, "7% aporte personal (régimen integral)"),
            ("IR",           0.00, "Escala 15-30% — tabla DGI"),
        ],
        "income_tax_freelance": 0.10,
        "income_tax_freelance_label": "Retención IR honorarios (10%)",
        "tax_authority": "DGI Nicaragua",
    },
    "SV": {
        "iva_rate": 0.13, "iva_label": "IVA",
        "currency_format": "$#,##0.00",
        "payroll_deductions": [
            ("ISSS", 0.03,   "3% aporte personal al ISSS"),
            ("AFP",  0.0725, "7.25% aporte personal AFP"),
            ("ISR",  0.00,   "Escala 10/20/30% — tabla DGII"),
        ],
        "income_tax_freelance": 0.10,
        "income_tax_freelance_label": "Retención ISR servicios profesionales (10%)",
        "tax_authority": "DGII El Salvador",
    },
    "CR": {
        "iva_rate": 0.13, "iva_label": "IVA",
        "currency_format": "₡#,##0.00",
        "payroll_deductions": [
            ("CCSS aporte personal", 0.1067, "10.67% — IVM + SEM + ROP"),
            ("Impuesto único",       0.00,   "Escala 0/10/15/20/25%"),
        ],
        "income_tax_freelance": 0.10,
        "income_tax_freelance_label": "Retención ISR servicios (10%)",
        "tax_authority": "Hacienda CR",
    },
    "CU": {
        "iva_rate": 0.10, "iva_label": "Impuesto sobre Ventas",
        "currency_format": "$#,##0.00",
        "payroll_deductions": [
            ("Seg. Social",       0.05, "5% aporte personal"),
            ("Imp. Personal",     0.00, "Según régimen ONAT"),
        ],
        "income_tax_freelance": 0.15,
        "income_tax_freelance_label": "Impuesto sobre Ingresos Personales — escala ONAT",
        "tax_authority": "ONAT",
    },
    "DO": {
        "iva_rate": 0.18, "iva_label": "ITBIS",
        "currency_format": "RD$#,##0.00",
        "payroll_deductions": [
            ("TSS SFS",         0.0304, "Seguro Familiar de Salud (3.04%)"),
            ("AFP",             0.0287, "Fondo Pensiones (2.87%)"),
            ("ISR",             0.00,   "Escala — tabla DGII"),
        ],
        "income_tax_freelance": 0.10,
        "income_tax_freelance_label": "Retención ISR honorarios (10%)",
        "tax_authority": "DGII Dominicana",
    },
    "GENERIC": {
        "iva_rate": 0.16, "iva_label": "IVA / Tax",
        "currency_format": "$#,##0.00",
        "payroll_deductions": [
            ("Seguridad Social",         0.07, "Tasa estándar — consulta tu país"),
            ("Pensión / Jubilación",     0.10, "Aporte personal típico"),
            ("Impuesto sobre la renta",  0.10, "Estimación — usa tabla local"),
        ],
        "income_tax_freelance": 0.10,
        "income_tax_freelance_label": "Retención honorarios (configurable)",
        "tax_authority": "Autoridad fiscal local",
    },
}


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS para builders de Excel
# ─────────────────────────────────────────────────────────────────────────────
def _country_code_from_cfg(cfg: dict) -> str:
    """Devuelve el código ISO del país a partir de cfg. Si no se encuentra,
    devuelve 'GENERIC'."""
    if not cfg:
        return "GENERIC"
    pais = cfg.get("pais", "")
    # cfg["pais"] es un string como "México" — invertir la búsqueda
    for code, entry in LEGAL_CONFIG.items():
        if entry.get("pais") == pais:
            return code
    return "GENERIC"


def get_excel_fiscal(cfg: dict) -> dict:
    """Devuelve el EXCEL_FISCAL apropiado para un cfg dado. Fallback a GENERIC.

    v15.16.0 — Aplica el overlay persistente verificado por IA sobre las tasas
    numéricas de Excel. El verifier fiscal guarda el IVA como string ("17%");
    aquí se convierte al decimal (0.17) que consumen las hojas de cálculo, para
    que las actualizaciones autónomas también lleguen a Excel, no solo a Word.
    """
    code = _country_code_from_cfg(cfg)
    base = dict(EXCEL_FISCAL.get(code, EXCEL_FISCAL["GENERIC"]))
    try:
        from agent.legal_overrides import get_overrides
        ov = get_overrides(code)
        if ov:
            # IVA: el verifier lo guarda como "17%" -> convertir a 0.17 decimal
            if "iva" in ov:
                raw = str(ov["iva"]).replace("%", "").strip()
                try:
                    base["iva_rate"] = float(raw) / 100.0
                except ValueError:
                    pass
            if "iva_label" in ov:
                base["iva_label"] = ov["iva_label"]
            if "autoridad_fiscal" in ov:
                base["tax_authority"] = ov["autoridad_fiscal"]
            base["_overlay_applied"] = [k for k in ("iva", "iva_label", "autoridad_fiscal") if k in ov]
    except Exception:
        pass
    return base


def iva_rate(cfg: dict) -> float:
    """Atajo: tasa IVA decimal del país en cfg."""
    return get_excel_fiscal(cfg).get("iva_rate", 0.16)


def iva_label(cfg: dict) -> str:
    """Atajo: cómo se llama el IVA en el país (IVA / IGV / ITBMS / …)."""
    return get_excel_fiscal(cfg).get("iva_label", "IVA")


def currency_format(cfg: dict) -> str:
    """Atajo: formato openpyxl de moneda para el país."""
    return get_excel_fiscal(cfg).get("currency_format", "$#,##0.00")


def currency_symbol(cfg: dict) -> str:
    """Atajo: símbolo de moneda ($, S/, ₡, Bs, Q, L, …)."""
    return (cfg or {}).get("simbolo", "$")


def currency_code(cfg: dict) -> str:
    """Atajo: código ISO de moneda (MXN, COP, PEN, …)."""
    return (cfg or {}).get("moneda", "USD")
