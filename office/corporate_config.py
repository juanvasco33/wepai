"""
WEP AI — Corporate Config (v15.6.0)
═══════════════════════════════════════════════════════════════════════════

Configuración corporativa por país para acuerdos de socios, estatutos
sociales y documentos societarios. Complementa a `legal_config.py` con
datos específicos del derecho de sociedades de cada jurisdicción
hispanohablante (19 países: 18 LATAM + España).

ESTRUCTURA POR PAÍS
    ley_sociedades       — Ley específica vigente
    autoridad_societaria — Organismo registrador (Registro Mercantil, SUNARP…)
    tipo_sociedad_default — La forma jurídica más común para startups
    tipos_sociedad        — Lista de formas jurídicas reconocidas
    reserva_legal_pct     — % anual a reservar de utilidades
    reserva_legal_tope    — % del capital social hasta el que se reserva
    mayoria_simple        — Mayoría para decisiones ordinarias
    mayoria_calificada    — Mayoría para decisiones extraordinarias
    mayoria_disolucion    — Mayoría para venta/disolución
    centro_arbitraje      — Centro de arbitraje recomendado del país
    ley_arbitraje         — Ley de arbitraje aplicable
    plazo_pago_dividendos — Plazo legal para pago de dividendos
    art_distribucion_utilidades — Artículo legal de distribución
    permite_acciones_preferentes — Boolean: si la jurisdicción reconoce acciones preferentes
    notas_vesting         — Notas específicas sobre cláusulas de vesting/cliff

USO
    from office.corporate_config import get_corporate_config
    corp = get_corporate_config(cfg)
    cfg.update(corp)

DATOS VERIFICADOS
    Los valores aquí son la *base estática*. Antes de generar el documento,
    `agent/corporate_verifier.py` consulta a Claude con web_search para
    verificar si la ley de sociedades, reservas, mayorías, etc. han cambiado.
    Si encuentra actualizaciones, las aplica al cfg en runtime.

NOTA LEGAL
    Estos datos son la mejor aproximación a la fecha del release. La ley
    corporativa de cada país puede cambiar. SIEMPRE revise el documento
    generado con un abogado licenciado en la jurisdicción correspondiente
    antes de firmarlo.
"""

# ── Configuración corporativa por código de país ────────────────────────────
CORPORATE_CONFIG = {
    # ── MÉXICO ────────────────────────────────────────────────────────────
    "MX": {
        "ley_sociedades": "Ley General de Sociedades Mercantiles (LGSM)",
        "autoridad_societaria": "Registro Público de Comercio (RPC)",
        "tipo_sociedad_default": "S.A. de C.V.",
        "tipos_sociedad": ["S.A. de C.V.", "S. de R.L. de C.V.", "S.A.P.I. de C.V.", "S.A.S."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "más del 50% del capital social",
        "mayoria_calificada": "75% del capital social (Art. 182 LGSM)",
        "mayoria_disolucion": "75% del capital social",
        "centro_arbitraje": "Centro de Arbitraje de México (CAM) o Centro de Mediación y Arbitraje de la CANACO",
        "ley_arbitraje": "Título Cuarto del Libro Quinto del Código de Comercio",
        "plazo_pago_dividendos": "dentro de los 90 días siguientes al acuerdo de la asamblea",
        "art_distribucion_utilidades": "Art. 16 LGSM",
        "permite_acciones_preferentes": True,
        "notas_vesting": "El vesting no está regulado expresamente en LGSM; se implementa vía pactos parasociales y cláusulas estatutarias en S.A.P.I.",
    },
    # ── COLOMBIA ──────────────────────────────────────────────────────────
    "CO": {
        "ley_sociedades": "Código de Comercio (Libro II) y Ley 1258 de 2008 para S.A.S.",
        "autoridad_societaria": "Cámara de Comercio / Superintendencia de Sociedades",
        "tipo_sociedad_default": "S.A.S. (Sociedad por Acciones Simplificada)",
        "tipos_sociedad": ["S.A.S.", "S.A.", "Ltda.", "S. en C.", "S. en C. por A."],
        "reserva_legal_pct": "10",
        "reserva_legal_tope": "50",
        "mayoria_simple": "más del 50% del capital social",
        "mayoria_calificada": "70% de las acciones suscritas (Art. 68 Ley 222 de 1995)",
        "mayoria_disolucion": "70% de las acciones suscritas",
        "centro_arbitraje": "Centro de Arbitraje y Conciliación de la Cámara de Comercio de Bogotá",
        "ley_arbitraje": "Ley 1563 de 2012 (Estatuto de Arbitraje Nacional e Internacional)",
        "plazo_pago_dividendos": "dentro del año siguiente a la aprobación del balance",
        "art_distribucion_utilidades": "Art. 451 C. de Co.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "La Ley 1258 (S.A.S.) permite amplia libertad estatutaria, ideal para cláusulas de vesting.",
    },
    # ── PERÚ ──────────────────────────────────────────────────────────────
    "PE": {
        "ley_sociedades": "Ley General de Sociedades N° 26887 (LGS)",
        "autoridad_societaria": "SUNARP (Superintendencia Nacional de los Registros Públicos)",
        "tipo_sociedad_default": "S.A.C. (Sociedad Anónima Cerrada)",
        "tipos_sociedad": ["S.A.C.", "S.A.A.", "S.A.", "S.R.L.", "S.C.R.L."],
        "reserva_legal_pct": "10",
        "reserva_legal_tope": "20",
        "mayoria_simple": "mayoría absoluta de las acciones representadas (Art. 127 LGS)",
        "mayoria_calificada": "dos tercios (2/3) de las acciones suscritas con derecho a voto (Art. 126 LGS)",
        "mayoria_disolucion": "2/3 de las acciones suscritas con derecho a voto",
        "centro_arbitraje": "Centro de Arbitraje de la Cámara de Comercio de Lima (CCL)",
        "ley_arbitraje": "Decreto Legislativo N° 1071 (Ley de Arbitraje)",
        "plazo_pago_dividendos": "dentro de los 30 días siguientes al acuerdo de la junta",
        "art_distribucion_utilidades": "Art. 230 LGS",
        "permite_acciones_preferentes": True,
        "notas_vesting": "S.A.C. permite restricciones a la libre transferencia y pactos sociales con vesting (Art. 237 LGS).",
    },
    # ── ARGENTINA ─────────────────────────────────────────────────────────
    "AR": {
        "ley_sociedades": "Ley General de Sociedades N° 19.550 (LGS)",
        "autoridad_societaria": "Inspección General de Justicia (IGJ) / Registro Público de Comercio",
        "tipo_sociedad_default": "S.A. (Sociedad Anónima) o S.A.S.",
        "tipos_sociedad": ["S.A.", "S.A.S.", "S.R.L.", "S. en C.", "S. de C. e I."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "mayoría absoluta de los votos presentes (Art. 243 LGS)",
        "mayoria_calificada": "mayoría de acciones con derecho a voto, sin pluralidad (Art. 244 LGS)",
        "mayoria_disolucion": "mayoría de acciones con derecho a voto",
        "centro_arbitraje": "Tribunal de Arbitraje General de la Bolsa de Comercio de Buenos Aires (TAG BCBA)",
        "ley_arbitraje": "Ley 27.449 de Arbitraje Comercial Internacional / Código Procesal Civil y Comercial",
        "plazo_pago_dividendos": "según lo establezca la asamblea, conforme Art. 224 LGS",
        "art_distribucion_utilidades": "Art. 68 y 224 LGS",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Ley 27.349 (Apoyo al Capital Emprendedor) introdujo la S.A.S. con flexibilidad para vesting.",
    },
    # ── CHILE ─────────────────────────────────────────────────────────────
    "CL": {
        "ley_sociedades": "Ley 18.046 sobre Sociedades Anónimas / Ley 20.190 (SpA)",
        "autoridad_societaria": "Comisión para el Mercado Financiero (CMF) / Conservador de Comercio",
        "tipo_sociedad_default": "SpA (Sociedad por Acciones)",
        "tipos_sociedad": ["SpA", "S.A. Cerrada", "S.A. Abierta", "Ltda.", "E.I.R.L."],
        "reserva_legal_pct": "no obligatoria",
        "reserva_legal_tope": "definida por estatutos",
        "mayoria_simple": "mayoría absoluta de las acciones presentes (Art. 67 Ley 18.046)",
        "mayoria_calificada": "dos tercios (2/3) de las acciones emitidas con derecho a voto (Art. 67)",
        "mayoria_disolucion": "2/3 de las acciones emitidas con derecho a voto",
        "centro_arbitraje": "Centro de Arbitraje y Mediación de Santiago (CAM Santiago) de la Cámara de Comercio",
        "ley_arbitraje": "Ley 19.971 de Arbitraje Comercial Internacional / Código Orgánico de Tribunales",
        "plazo_pago_dividendos": "dentro del año siguiente a la junta que los acordó",
        "art_distribucion_utilidades": "Art. 78 Ley 18.046",
        "permite_acciones_preferentes": True,
        "notas_vesting": "La SpA (Ley 20.190) permite máxima libertad estatutaria, ideal para vesting y cap tables complejas.",
    },
    # ── ESPAÑA ────────────────────────────────────────────────────────────
    "ES": {
        "ley_sociedades": "Real Decreto Legislativo 1/2010 — Ley de Sociedades de Capital (LSC)",
        "autoridad_societaria": "Registro Mercantil",
        "tipo_sociedad_default": "S.L. (Sociedad Limitada)",
        "tipos_sociedad": ["S.L.", "S.A.", "S.L.U.", "S.A.U.", "S.L.N.E.", "S. Coop."],
        "reserva_legal_pct": "10",
        "reserva_legal_tope": "20",
        "mayoria_simple": "más de la mitad de los votos válidamente emitidos (Art. 198 LSC)",
        "mayoria_calificada": "dos tercios (2/3) de los votos correspondientes a las participaciones (Art. 199 LSC)",
        "mayoria_disolucion": "2/3 del capital con derecho a voto",
        "centro_arbitraje": "Corte Civil y Mercantil de Arbitraje (CIMA) / Corte Española de Arbitraje",
        "ley_arbitraje": "Ley 60/2003 de Arbitraje",
        "plazo_pago_dividendos": "fecha acordada por la junta (por defecto, día siguiente al acuerdo)",
        "art_distribucion_utilidades": "Art. 273-277 LSC",
        "permite_acciones_preferentes": True,
        "notas_vesting": "S.L. permite restricciones estatutarias amplias; recomendable pacto de socios separado para vesting.",
    },
    # ── BOLIVIA ───────────────────────────────────────────────────────────
    "BO": {
        "ley_sociedades": "Código de Comercio (Libro Primero, Título III)",
        "autoridad_societaria": "Registro de Comercio (FUNDEMPRESA)",
        "tipo_sociedad_default": "S.R.L. (Sociedad de Responsabilidad Limitada)",
        "tipos_sociedad": ["S.R.L.", "S.A.", "S. en C.", "S. en N.C."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "50",
        "mayoria_simple": "más del 50% del capital social (Art. 199 C.Com.)",
        "mayoria_calificada": "2/3 del capital social para modificaciones estatutarias",
        "mayoria_disolucion": "2/3 del capital social",
        "centro_arbitraje": "Centro de Conciliación y Arbitraje de la Cámara Nacional de Comercio (CNC)",
        "ley_arbitraje": "Ley 708 de Conciliación y Arbitraje (2015)",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 168 C.Com.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "No regulado expresamente; se implementa vía pactos estatutarios y de socios.",
    },
    # ── ECUADOR ───────────────────────────────────────────────────────────
    "EC": {
        "ley_sociedades": "Ley de Compañías (Codificación 2014, reformas 2020 con SAS)",
        "autoridad_societaria": "Superintendencia de Compañías, Valores y Seguros",
        "tipo_sociedad_default": "S.A.S. (Sociedad por Acciones Simplificada)",
        "tipos_sociedad": ["S.A.S.", "S.A.", "Cía. Ltda.", "S. en C."],
        "reserva_legal_pct": "10",
        "reserva_legal_tope": "50",
        "mayoria_simple": "más del 50% del capital pagado",
        "mayoria_calificada": "dos tercios (2/3) del capital pagado (Art. 240 Ley de Compañías)",
        "mayoria_disolucion": "2/3 del capital pagado",
        "centro_arbitraje": "Centro de Arbitraje y Mediación de la Cámara de Comercio de Quito",
        "ley_arbitraje": "Ley de Arbitraje y Mediación (Registro Oficial 145 de 1997)",
        "plazo_pago_dividendos": "según resolución de la junta general",
        "art_distribucion_utilidades": "Art. 297 Ley de Compañías",
        "permite_acciones_preferentes": True,
        "notas_vesting": "La S.A.S. (Ley Orgánica de Emprendimiento e Innovación 2020) permite estatutos flexibles para vesting.",
    },
    # ── URUGUAY ───────────────────────────────────────────────────────────
    "UY": {
        "ley_sociedades": "Ley 16.060 de Sociedades Comerciales / Ley 19.820 (SAS)",
        "autoridad_societaria": "Registro Nacional de Comercio",
        "tipo_sociedad_default": "S.A.S. (Sociedad por Acciones Simplificada)",
        "tipos_sociedad": ["S.A.S.", "S.A.", "S.R.L.", "S. en C."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "mayoría absoluta de capital con derecho a voto",
        "mayoria_calificada": "mayoría absoluta del capital integrado con derecho a voto (Art. 356 Ley 16.060)",
        "mayoria_disolucion": "mayoría absoluta del capital integrado",
        "centro_arbitraje": "Centro de Conciliación y Arbitraje de la Cámara Nacional de Comercio y Servicios del Uruguay",
        "ley_arbitraje": "Ley 19.636 de Arbitraje Comercial Internacional / Código General del Proceso",
        "plazo_pago_dividendos": "según lo determine la asamblea",
        "art_distribucion_utilidades": "Art. 93 y 320 Ley 16.060",
        "permite_acciones_preferentes": True,
        "notas_vesting": "La Ley 19.820 (S.A.S.) permite máxima libertad estatutaria.",
    },
    # ── PARAGUAY ──────────────────────────────────────────────────────────
    "PY": {
        "ley_sociedades": "Código Civil (Libro III, Título VIII) / Ley 388/94 / Ley 6480/2020 (EAS)",
        "autoridad_societaria": "Dirección General de los Registros Públicos",
        "tipo_sociedad_default": "E.A.S. (Empresa por Acciones Simplificada)",
        "tipos_sociedad": ["E.A.S.", "S.A.", "S.R.L.", "S. en C."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "más del 50% del capital integrado",
        "mayoria_calificada": "dos tercios (2/3) del capital integrado",
        "mayoria_disolucion": "2/3 del capital integrado",
        "centro_arbitraje": "Centro de Arbitraje y Mediación del Paraguay (CAMP) — Cámara Nacional de Comercio",
        "ley_arbitraje": "Ley 1879/02 de Arbitraje y Mediación",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 1080 Código Civil",
        "permite_acciones_preferentes": True,
        "notas_vesting": "La EAS (Ley 6480/2020) permite estatutos flexibles, ideal para startups con vesting.",
    },
    # ── VENEZUELA ─────────────────────────────────────────────────────────
    "VE": {
        "ley_sociedades": "Código de Comercio (Libro Primero, Título VII)",
        "autoridad_societaria": "Registro Mercantil",
        "tipo_sociedad_default": "C.A. (Compañía Anónima)",
        "tipos_sociedad": ["C.A.", "S.A.", "S.R.L.", "S. en C.A."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "10",
        "mayoria_simple": "más del 50% del capital social",
        "mayoria_calificada": "tres cuartos (3/4) del capital social (Art. 280 C.Com.)",
        "mayoria_disolucion": "3/4 del capital social",
        "centro_arbitraje": "Centro Empresarial de Conciliación y Arbitraje (CEDCA) o Centro de Arbitraje de la Cámara de Caracas",
        "ley_arbitraje": "Ley de Arbitraje Comercial (Gaceta Oficial 36.430 de 1998)",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 307 C.Com.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Por inflación VED se recomienda fijar montos en USD equivalente o UVR; cláusulas de ajuste cambiario obligatorias.",
    },
    # ── PANAMÁ ────────────────────────────────────────────────────────────
    "PA": {
        "ley_sociedades": "Ley 32 de 1927 (S.A.) / Ley 4 de 2009 (S.A.S.) / Ley 4 de 2009 (E.I.P.)",
        "autoridad_societaria": "Registro Público de Panamá",
        "tipo_sociedad_default": "S.A. (Sociedad Anónima)",
        "tipos_sociedad": ["S.A.", "S.R.L.", "S.A.S.", "E.I.P."],
        "reserva_legal_pct": "no obligatoria por ley general",
        "reserva_legal_tope": "definida por estatutos",
        "mayoria_simple": "mayoría de los votos presentes",
        "mayoria_calificada": "mayoría establecida en el pacto social (típicamente 2/3)",
        "mayoria_disolucion": "mayoría establecida en el pacto social",
        "centro_arbitraje": "Centro de Conciliación y Arbitraje de Panamá (CeCAP) de la Cámara de Comercio",
        "ley_arbitraje": "Ley 131 de 2013 de Arbitraje Comercial Nacional e Internacional",
        "plazo_pago_dividendos": "según resolución de la junta",
        "art_distribucion_utilidades": "Art. 47-49 Ley 32 de 1927",
        "permite_acciones_preferentes": True,
        "notas_vesting": "La Ley 32 ofrece máxima flexibilidad estatutaria; vesting se implementa vía pactos accesorios.",
    },
    # ── GUATEMALA ─────────────────────────────────────────────────────────
    "GT": {
        "ley_sociedades": "Código de Comercio de Guatemala (Decreto 2-70)",
        "autoridad_societaria": "Registro Mercantil",
        "tipo_sociedad_default": "S.A. (Sociedad Anónima)",
        "tipos_sociedad": ["S.A.", "S.R.L.", "S. en C.", "S. en N.C."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "15",
        "mayoria_simple": "mayoría de las acciones presentes",
        "mayoria_calificada": "dos tercios (2/3) de las acciones emitidas (Art. 149 C.Com.)",
        "mayoria_disolucion": "2/3 de las acciones emitidas",
        "centro_arbitraje": "Centro de Arbitraje y Conciliación de la Cámara de Comercio de Guatemala (CENAC)",
        "ley_arbitraje": "Ley de Arbitraje (Decreto 67-95)",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 33 y 150 C.Com.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "No regulado expresamente; se implementa vía estatutos y convenios accesorios.",
    },
    # ── HONDURAS ──────────────────────────────────────────────────────────
    "HN": {
        "ley_sociedades": "Código de Comercio (Decreto 73-50)",
        "autoridad_societaria": "Registro Mercantil de la Cámara de Comercio",
        "tipo_sociedad_default": "S.A. (Sociedad Anónima)",
        "tipos_sociedad": ["S.A.", "S. de R.L.", "S. en C.", "S. en N.C."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "más del 50% del capital social",
        "mayoria_calificada": "tres cuartos (3/4) del capital social para modificaciones",
        "mayoria_disolucion": "3/4 del capital social",
        "centro_arbitraje": "Centro de Conciliación y Arbitraje (CCA) de la Cámara de Comercio e Industrias de Tegucigalpa",
        "ley_arbitraje": "Ley de Conciliación y Arbitraje (Decreto 161-2000)",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 32 C.Com.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Se recomienda pacto de socios separado para cláusulas de vesting.",
    },
    # ── NICARAGUA ─────────────────────────────────────────────────────────
    "NI": {
        "ley_sociedades": "Código de Comercio de la República de Nicaragua",
        "autoridad_societaria": "Registro Público Mercantil",
        "tipo_sociedad_default": "S.A. (Sociedad Anónima)",
        "tipos_sociedad": ["S.A.", "S. en C.", "S. en N.C."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "10",
        "mayoria_simple": "más del 50% del capital social",
        "mayoria_calificada": "tres cuartos (3/4) del capital social",
        "mayoria_disolucion": "3/4 del capital social",
        "centro_arbitraje": "Centro de Mediación y Arbitraje (CMA) de la Cámara de Comercio y Servicios de Nicaragua",
        "ley_arbitraje": "Ley 540 de Mediación y Arbitraje",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 257 C.Com.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "No regulado expresamente; implementar vía estatutos y pactos parasociales.",
    },
    # ── EL SALVADOR ───────────────────────────────────────────────────────
    "SV": {
        "ley_sociedades": "Código de Comercio (Decreto 671)",
        "autoridad_societaria": "Centro Nacional de Registros (CNR) — Registro de Comercio",
        "tipo_sociedad_default": "S.A. de C.V.",
        "tipos_sociedad": ["S.A. de C.V.", "S. de R.L. de C.V.", "S. en C.", "S. en N.C."],
        "reserva_legal_pct": "7",
        "reserva_legal_tope": "20",
        "mayoria_simple": "más del 50% del capital social",
        "mayoria_calificada": "tres cuartos (3/4) de las acciones (Art. 248 C.Com.)",
        "mayoria_disolucion": "3/4 de las acciones",
        "centro_arbitraje": "Centro de Mediación y Arbitraje (CMA) de la Cámara de Comercio e Industria de El Salvador",
        "ley_arbitraje": "Ley de Mediación, Conciliación y Arbitraje (Decreto 914 de 2002)",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 122 C.Com.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Tras dolarización (2001) los montos se expresan en USD; cláusulas de vesting vía estatutos.",
    },
    # ── COSTA RICA ────────────────────────────────────────────────────────
    "CR": {
        "ley_sociedades": "Código de Comercio (Ley 3284)",
        "autoridad_societaria": "Registro Nacional — Registro de Personas Jurídicas",
        "tipo_sociedad_default": "S.A. (Sociedad Anónima)",
        "tipos_sociedad": ["S.A.", "S.R.L.", "S. en C.", "S. en N.C."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "más del 50% del capital social",
        "mayoria_calificada": "tres cuartos (3/4) del capital social",
        "mayoria_disolucion": "3/4 del capital social",
        "centro_arbitraje": "Centro de Conciliación y Arbitraje (CCA) de la Cámara de Comercio de Costa Rica",
        "ley_arbitraje": "Ley 7727 sobre Resolución Alterna de Conflictos / Ley 8937 de Arbitraje Comercial Internacional",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 143 C.Com.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "S.A. permite múltiples clases de acciones; recomendable pacto parasocial complementario.",
    },
    # ── CUBA ──────────────────────────────────────────────────────────────
    "CU": {
        "ley_sociedades": "Código de Comercio (1885 vigente con reformas) / Decreto-Ley 304/2012 / Decreto-Ley 46/2021 (MIPYMES)",
        "autoridad_societaria": "Registro Mercantil del Ministerio de Justicia",
        "tipo_sociedad_default": "S.A. (Sociedad Anónima) / MIPYME",
        "tipos_sociedad": ["S.A.", "S.R.L.", "Cooperativa No Agropecuaria", "MIPYME"],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "más del 50% del capital social",
        "mayoria_calificada": "dos tercios (2/3) del capital social",
        "mayoria_disolucion": "2/3 del capital social",
        "centro_arbitraje": "Corte Cubana de Arbitraje Comercial Internacional (Cámara de Comercio de la República de Cuba)",
        "ley_arbitraje": "Decreto-Ley 250 de 2007 — Corte Cubana de Arbitraje",
        "plazo_pago_dividendos": "según resolución del órgano de gobierno",
        "art_distribucion_utilidades": "Art. 165 C.Com.",
        "permite_acciones_preferentes": False,  # restringido en el marco actual
        "notas_vesting": "Marco corporativo restringido; aplicable principalmente a empresas mixtas y MIPYMES. Consultar regulación específica.",
    },
    # ── REPÚBLICA DOMINICANA ──────────────────────────────────────────────
    "DO": {
        "ley_sociedades": "Ley 479-08 General de las Sociedades Comerciales y Empresas Individuales (modificada por Ley 31-11)",
        "autoridad_societaria": "Cámara de Comercio y Producción / Registro Mercantil",
        "tipo_sociedad_default": "S.R.L. (Sociedad de Responsabilidad Limitada)",
        "tipos_sociedad": ["S.R.L.", "S.A.", "S.A.S.", "E.I.R.L."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "10",
        "mayoria_simple": "más del 50% de las cuotas/acciones",
        "mayoria_calificada": "tres cuartos (3/4) del capital social (Art. 122 Ley 479-08)",
        "mayoria_disolucion": "3/4 del capital social",
        "centro_arbitraje": "Centro de Resolución Alternativa de Controversias (CRC) de la Cámara de Comercio de Santo Domingo",
        "ley_arbitraje": "Ley 489-08 de Arbitraje Comercial",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "Art. 44 y 47 Ley 479-08",
        "permite_acciones_preferentes": True,
        "notas_vesting": "La Ley 479-08 permite múltiples categorías de acciones y restricciones a la transferencia.",
    },
    # ── ESTADOS UNIDOS ────────────────────────────────────────────────────
    "US": {
        "ley_sociedades": "Delaware General Corporation Law (DGCL) / Model Business Corporation Act (MBCA)",
        "autoridad_societaria": "Delaware Division of Corporations / Secretary of State (by state of incorporation)",
        "tipo_sociedad_default": "Delaware C-Corporation",
        "tipos_sociedad": ["C-Corp", "S-Corp", "LLC", "LP", "LLP", "PBC (Public Benefit Corp)"],
        "reserva_legal_pct": "not required by federal law",
        "reserva_legal_tope": "as defined in bylaws",
        "mayoria_simple": "majority of shares present and entitled to vote",
        "mayoria_calificada": "two-thirds (2/3) of outstanding shares (DGCL §242)",
        "mayoria_disolucion": "majority of outstanding shares (DGCL §275)",
        "centro_arbitraje": "American Arbitration Association (AAA) / JAMS",
        "ley_arbitraje": "Federal Arbitration Act (9 U.S.C. §§ 1-16)",
        "plazo_pago_dividendos": "as declared by the Board of Directors",
        "art_distribucion_utilidades": "DGCL §170",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Standard practice: 4-year vesting with 1-year cliff. 83(b) election available for early-exercised stock options. ISO/NSO distinction relevant.",
    },
    # ── BRASIL ────────────────────────────────────────────────────────────
    "BR": {
        "ley_sociedades": "Lei 6.404/76 (Lei das Sociedades por Ações) / Lei 10.406/02 (Código Civil — Sociedades Limitadas)",
        "autoridad_societaria": "Junta Comercial estadual (DREI federal)",
        "tipo_sociedad_default": "Ltda. (Sociedade Limitada)",
        "tipos_sociedad": ["Ltda.", "S.A.", "EIRELI", "S.A. de Capital Fechado", "S.A. de Capital Aberto"],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "maioria absoluta do capital social",
        "mayoria_calificada": "três quartos (3/4) do capital social (Lei 6.404/76 art. 136)",
        "mayoria_disolucion": "3/4 do capital social",
        "centro_arbitraje": "Câmara de Arbitragem do Mercado (CAM) — B3 / CAM-CCBC",
        "ley_arbitraje": "Lei 9.307/96 (Lei de Arbitragem)",
        "plazo_pago_dividendos": "no prazo de 60 dias da deliberação, salvo disposição estatutária",
        "art_distribucion_utilidades": "Lei 6.404/76 art. 201-205",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Vesting via acordo de acionistas; opções de compra de ações regulamentadas pela CVM.",
    },
    # ── FRANCIA ───────────────────────────────────────────────────────────
    "FR": {
        "ley_sociedades": "Code de Commerce (Livre II) / Loi PACTE 2019",
        "autoridad_societaria": "Registre du Commerce et des Sociétés (RCS)",
        "tipo_sociedad_default": "SAS (Société par Actions Simplifiée)",
        "tipos_sociedad": ["SAS", "SA", "SARL", "SASU", "SCI", "SNC"],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "10",
        "mayoria_simple": "majorité simple (50% + 1)",
        "mayoria_calificada": "deux tiers (2/3) des actions (Art. L. 225-96 C.com.)",
        "mayoria_disolucion": "2/3 des actions",
        "centro_arbitraje": "Chambre de Commerce Internationale (CCI/ICC) Paris",
        "ley_arbitraje": "Code de Procédure Civile (Livre IV) — Décret 2011-48",
        "plazo_pago_dividendos": "dans les 9 mois suivant la clôture de l'exercice",
        "art_distribucion_utilidades": "Art. L. 232-12 C.com.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "La SAS offre la plus grande liberté statutaire pour clauses de vesting et BSPCE (Bons de Souscription de Parts de Créateur d'Entreprise).",
    },
    # ── ITALIA ────────────────────────────────────────────────────────────
    "IT": {
        "ley_sociedades": "Codice Civile (Libro V, Titolo V) / D.Lgs. 6/2003",
        "autoridad_societaria": "Registro delle Imprese (Camera di Commercio)",
        "tipo_sociedad_default": "S.r.l. (Società a responsabilità limitata)",
        "tipos_sociedad": ["S.r.l.", "S.p.A.", "S.r.l.s.", "S.a.p.a.", "S.n.c.", "S.a.s."],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "maggioranza assoluta del capitale presente",
        "mayoria_calificada": "almeno due terzi (2/3) del capitale sociale (Art. 2479-bis C.C.)",
        "mayoria_disolucion": "2/3 del capitale sociale",
        "centro_arbitraje": "Camera Arbitrale di Milano (CAM) / Camera Arbitrale Nazionale e Internazionale",
        "ley_arbitraje": "Codice di Procedura Civile (artt. 806-840)",
        "plazo_pago_dividendos": "secondo deliberazione assembleare",
        "art_distribucion_utilidades": "Art. 2433 C.C.",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Vesting implementato via patti parasociali; S.r.l. innovativa offre maggiore flessibilità.",
    },
    # ── PORTUGAL ──────────────────────────────────────────────────────────
    "PT": {
        "ley_sociedades": "Código das Sociedades Comerciais (Decreto-Lei 262/86)",
        "autoridad_societaria": "Registo Comercial / Instituto dos Registos e do Notariado (IRN)",
        "tipo_sociedad_default": "Lda. (Sociedade por Quotas)",
        "tipos_sociedad": ["Lda.", "S.A.", "Unipessoal Lda.", "S.G.P.S.", "S. em comandita"],
        "reserva_legal_pct": "5",
        "reserva_legal_tope": "20",
        "mayoria_simple": "maioria simples dos votos emitidos",
        "mayoria_calificada": "três quartos (3/4) dos votos correspondentes ao capital social (CSC art. 265)",
        "mayoria_disolucion": "3/4 do capital social",
        "centro_arbitraje": "Centro de Arbitragem Comercial (CAC) — Câmara de Comércio e Indústria Portuguesa",
        "ley_arbitraje": "Lei 63/2011 — Lei da Arbitragem Voluntária",
        "plazo_pago_dividendos": "no prazo fixado pela assembleia geral",
        "art_distribucion_utilidades": "CSC art. 217 e 294",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Vesting via acordos parassociais; regime das stock options regulado pelo CIRS.",
    },
    # ── GENERIC (fallback) ────────────────────────────────────────────────
    "GENERIC": {
        "ley_sociedades": "[ley de sociedades aplicable]",
        "autoridad_societaria": "[Registro Mercantil correspondiente]",
        "tipo_sociedad_default": "[forma jurídica]",
        "tipos_sociedad": ["S.A.", "S.R.L."],
        "reserva_legal_pct": "[%]",
        "reserva_legal_tope": "[%]",
        "mayoria_simple": "mayoría simple según la ley aplicable",
        "mayoria_calificada": "mayoría calificada según la ley aplicable",
        "mayoria_disolucion": "mayoría calificada según la ley aplicable",
        "centro_arbitraje": "[Centro de arbitraje del país]",
        "ley_arbitraje": "[ley de arbitraje aplicable]",
        "plazo_pago_dividendos": "según resolución de la asamblea",
        "art_distribucion_utilidades": "[artículo aplicable]",
        "permite_acciones_preferentes": True,
        "notas_vesting": "Implementar vía estatutos y/o pacto de socios separado.",
    },
}


def get_corporate_config(cfg: dict) -> dict:
    """Obtiene la configuración corporativa para el país detectado en cfg.

    Args:
        cfg: configuración base (de legal_config.py) con campo 'pais'

    Returns:
        dict con campos corporativos del país, o el de GENERIC si no se encuentra.
    """
    try:
        from office.legal_config import LEGAL_CONFIG
        pais = cfg.get("pais", "")
        for code, entry in LEGAL_CONFIG.items():
            if entry.get("pais") == pais:
                return CORPORATE_CONFIG.get(code, CORPORATE_CONFIG["GENERIC"]).copy()
    except Exception:
        pass
    return CORPORATE_CONFIG["GENERIC"].copy()


def merge_corporate(cfg: dict) -> dict:
    """Devuelve un nuevo cfg que combina los campos legales base con
    los campos corporativos del país. No muta el cfg original.
    """
    merged = dict(cfg)
    corp   = get_corporate_config(cfg)
    # Solo añadir claves que no existan ya en cfg (legal_config tiene prioridad)
    for k, v in corp.items():
        merged.setdefault(k, v)
    return merged
