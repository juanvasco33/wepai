"""
WEP AI — LaborReformDiagnostic (v16.7)

Módulo de DIAGNÓSTICO de contratos frente a la Reforma Laboral colombiana
(Ley 2466 de 2025, D.O. 53.160, vigente desde el 25 de junio de 2025).

═══════════════════════════════════════════════════════════════════════════════
QUÉ ES Y QUÉ NO ES — LÉASE ANTES DE MODIFICAR
═══════════════════════════════════════════════════════════════════════════════

Este módulo NO emite dictámenes legales. NO afirma que un contrato sea "ilegal"
ni indica qué hacer en un caso concreto. Eso constituiría ejercicio del derecho
y expone a WEP AI y al usuario a responsabilidad.

Lo que SÍ hace: a partir de la descripción de un contrato existente, SEÑALA
puntos que PODRÍAN no alinearse con la Ley 2466 de 2025, cita el artículo o la
regla pública correspondiente, y RECOMIENDA revisarlos con un profesional. Es una
herramienta educativa y de alerta, no asesoría jurídica.

Toda salida de este módulo:
  • Usa lenguaje condicional ("podría", "te recomendamos revisar"), nunca
    afirmativo sobre el estatus legal.
  • Cita la fuente normativa (artículo / regla de la Ley 2466) de cada alerta.
  • Incluye un disclaimer visible.
  • Marca el resultado como material de referencia, no como documento final.

═══════════════════════════════════════════════════════════════════════════════
DOS MODOS DE OPERACIÓN
═══════════════════════════════════════════════════════════════════════════════

1. MODO REGLAS (siempre disponible, sin API key): aplica un conjunto de reglas
   deterministas derivadas de la Ley 2466 sobre los datos estructurados que el
   usuario describe (tipo de contrato, duración, si es escrito, jornada nocturna,
   procedimiento de despido, etc.). Es rápido, gratis y predecible.

2. MODO IA (si hay ANTHROPIC_API_KEY + SDK): además de las reglas, usa el modelo
   para leer texto libre de un contrato pegado por el usuario y detectar señales
   que las reglas no capturan. SIEMPRE pasa por el mismo filtro de lenguaje
   seguro y citación.

El MODO REGLAS es la base de confianza. El MODO IA es un complemento que nunca
sustituye el disclaimer ni inventa estatus legal.
"""

import os
import json
from datetime import date

try:
    from config import DEFAULT_CHAT_MODEL as _MODEL
except Exception:
    _MODEL = "claude-sonnet-4-6"


# ─────────────────────────────────────────────────────────────────────────────
# VALORES ESCALONADOS POR FECHA (Ley 2466 de 2025)
# ─────────────────────────────────────────────────────────────────────────────
# Varios valores de la reforma cambian en fechas concretas. En vez de comparar
# contra una constante (que queda obsoleta y genera falsos positivos), estas
# funciones devuelven el valor VIGENTE según la fecha actual. Así el diagnóstico
# es correcto hoy y sigue siéndolo cuando cambien las etapas (sin tocar código).
# Fuentes oficiales: texto de la Ley 2466 (Función Pública) y parágrafos
# transitorios de los Art. 11 y 14.

def recargo_dominical_vigente(hoy: date = None) -> int:
    """
    % de recargo dominical/festivo vigente según la fecha.
    Escalonamiento (Art. 14 Ley 2466/2025, parágrafo transitorio):
      • 1-jul-2025 a 30-jun-2026: 80%
      • 1-jul-2026 a 30-jun-2027: 90%
      • desde 1-jul-2027:         100%
      • antes del 1-jul-2025:     75% (régimen anterior)
    """
    hoy = hoy or date.today()
    if hoy >= date(2027, 7, 1):
        return 100
    if hoy >= date(2026, 7, 1):
        return 90
    if hoy >= date(2025, 7, 1):
        return 80
    return 75


def jornada_maxima_vigente(hoy: date = None) -> int:
    """
    Horas/semana máximas de la jornada ordinaria vigentes según la fecha.
    Reducción gradual (Ley 2101/2021, ratificada por Ley 2466/2025):
      • hasta 14-jul-2026: 44 h/semana
      • desde 15-jul-2026:  42 h/semana
    """
    hoy = hoy or date.today()
    if hoy >= date(2026, 7, 15):
        return 42
    return 44

# Disclaimer único, reutilizado en toda salida del módulo.
DISCLAIMER = (
    "⚖️ AVISO IMPORTANTE: Esto es una herramienta educativa, NO asesoría "
    "jurídica. WEP AI no es un abogado y no puede confirmar el estatus legal de "
    "tu contrato. Las observaciones señalan puntos que PODRÍAN requerir revisión "
    "frente a la Ley 2466 de 2025; verifícalos con un profesional del derecho "
    "antes de tomar cualquier decisión, firmar o modificar un documento."
)

FUENTE = "Ley 2466 de 2025 (Reforma Laboral, D.O. 53.160, vigente desde el 25 de junio de 2025)"


# ─────────────────────────────────────────────────────────────────────────────
# REGLAS DETERMINISTAS — el núcleo confiable del diagnóstico
# ─────────────────────────────────────────────────────────────────────────────
# Cada regla recibe el dict de datos del contrato y devuelve None (sin alerta)
# o una alerta dict: {punto, observacion, fuente, sugerencia}.
#
# Las claves esperadas en `datos` (todas opcionales; la regla se salta si falta
# el dato necesario):
#   tipo_contrato        : "indefinido" | "fijo" | "obra_labor" | "prestacion_servicios" | "aprendizaje" | str
#   por_escrito          : bool
#   duracion_meses       : int | float   (para contratos a término fijo)
#   num_prorrogas        : int
#   jornada_nocturna     : bool          (¿el cargo trabaja de noche?)
#   hora_inicio_nocturna : str           (ej. "21:00", "19:00")
#   es_microempresa      : bool
#   recargo_dominical_pct: int           (ej. 75, 100)
#   tiene_procedimiento_disciplinario : bool  (¿existe procedimiento previo al despido?)
#   num_empleados        : int
#   num_aprendices       : int
#   horas_semanales      : int | float


def _r_fijo_sin_escrito(d):
    if d.get("tipo_contrato") in ("fijo", "obra_labor") and d.get("por_escrito") is False:
        return {
            "punto": "Contrato a término fijo u obra/labor sin constar por escrito",
            "observacion": (
                "Un contrato a término fijo o por obra/labor que no consta por "
                "escrito PODRÍA entenderse celebrado a término indefinido desde "
                "el inicio."
            ),
            "fuente": f"{FUENTE} — exigencia de forma escrita para modalidades distintas al indefinido",
            "sugerencia": "Revisa con un profesional si conviene formalizarlo por escrito o reconocerlo como indefinido.",
        }
    return None


def _r_fijo_excede_4_anios(d):
    dur = d.get("duracion_meses")
    if d.get("tipo_contrato") == "fijo" and isinstance(dur, (int, float)) and dur > 48:
        return {
            "punto": "Contrato a término fijo con duración superior a 4 años",
            "observacion": (
                "La duración total de un contrato a término fijo, incluidas sus "
                "prórrogas, PODRÍA estar limitada a un máximo de 4 años."
            ),
            "fuente": f"{FUENTE} — límite de duración del contrato a término fijo",
            "sugerencia": "Revisa con un profesional la duración acumulada y si corresponde pasar a término indefinido.",
        }
    return None


def _r_fijo_como_norma(d):
    # Señal blanda: usar fijo cuando la labor parece permanente.
    if d.get("tipo_contrato") == "fijo" and d.get("labor_permanente") is True:
        return {
            "punto": "Uso de contrato a término fijo para una labor de carácter permanente",
            "observacion": (
                "Desde la reforma, el contrato a término indefinido es la regla "
                "general y las demás modalidades son la excepción, que debe "
                "justificarse por la naturaleza temporal de la labor. Usar un "
                "fijo para una labor permanente PODRÍA ser cuestionado."
            ),
            "fuente": f"{FUENTE}, Art. 5 — el indefinido como modalidad preferente",
            "sugerencia": "Revisa con un profesional si la naturaleza de la labor justifica una modalidad distinta al indefinido.",
        }
    return None


def _r_jornada_nocturna(d):
    # CORREGIDO v16.9: la jornada nocturna a las 7:00 p.m. aplica a TODAS las
    # empresas sin excepción por tamaño (Art. 11 Ley 2466/2025). Se eliminó la
    # falsa excepción de microempresa que existía en v16.7-16.8.
    if d.get("jornada_nocturna") is True:
        hora = d.get("hora_inicio_nocturna")
        if hora and ("21" in str(hora) or "9" in str(hora).lower().replace("pm", "").replace("p.m.", "")):
            return {
                "punto": "Inicio de la jornada nocturna calculado desde las 9:00 p.m.",
                "observacion": (
                    "Desde el 25 de diciembre de 2025, el trabajo nocturno se "
                    "cuenta a partir de las 7:00 p.m. (antes 9:00 p.m.), para "
                    "TODAS las empresas sin importar su tamaño. Calcular el "
                    "recargo nocturno (35%) desde las 9:00 p.m. PODRÍA dejar sin "
                    "pagar las horas entre las 7:00 y las 9:00 p.m."
                ),
                "fuente": f"{FUENTE}, Art. 11 — trabajo nocturno de 7:00 p.m. a 6:00 a.m., vigente desde el 25-dic-2025",
                "sugerencia": "Revisa con un profesional el cálculo del recargo nocturno desde las 7:00 p.m. (aplica a empresas de cualquier tamaño).",
            }
    return None


def _r_recargo_dominical(d):
    # CORREGIDO v16.9: compara contra el % VIGENTE según la fecha, no contra 100.
    # Antes alertaba siempre que fuera < 100%, generando falsos positivos
    # (ej. un 90% correcto en 2026 se marcaba como error).
    pct = d.get("recargo_dominical_pct")
    if isinstance(pct, (int, float)):
        vigente = recargo_dominical_vigente()
        if pct < vigente:
            return {
                "punto": f"Recargo dominical/festivo pactado en {pct}% (vigente actual: {vigente}%)",
                "observacion": (
                    f"El recargo por trabajo en dominicales y festivos aumenta de "
                    f"forma escalonada. En la etapa vigente corresponde {vigente}% "
                    f"(80% desde jul-2025, 90% desde jul-2026, 100% desde jul-2027). "
                    f"Un {pct}% PODRÍA quedar por debajo de lo exigido hoy."
                ),
                "fuente": f"{FUENTE}, Art. 14 — incremento escalonado del recargo dominical/festivo (parágrafo transitorio)",
                "sugerencia": f"Revisa con un profesional el ajuste del recargo al {vigente}% vigente y recalcula la nómina afectada.",
            }
    return None


def _r_procedimiento_disciplinario(d):
    if d.get("tiene_procedimiento_disciplinario") is False:
        n = d.get("num_empleados")
        # Régimen simplificado para empresas con menos de 10 empleados.
        if isinstance(n, (int, float)) and n < 10:
            return {
                "punto": "Sin procedimiento previo al despido/sanción con justa causa (empresa pequeña)",
                "observacion": (
                    "Aun en empresas con menos de 10 empleados, antes de un "
                    "despido o sanción con justa causa se debe al menos ESCUCHAR "
                    "previamente al trabajador (procedimiento simplificado). "
                    "Omitirlo PODRÍA viciar la decisión."
                ),
                "fuente": f"{FUENTE} — procedimiento simplificado para empresas de menos de 10 empleados",
                "sugerencia": "Revisa con un profesional cómo dejar constancia de haber escuchado al trabajador antes de decidir.",
            }
        return {
            "punto": "Ausencia de procedimiento disciplinario previo al despido con justa causa",
            "observacion": (
                "La reforma establece un procedimiento mínimo previo al despido "
                "con justa causa (notificación de los hechos, derecho de defensa, "
                "investigación, descargos, decisión fundamentada, notificación y "
                "recurso). Un despido sin este procedimiento PODRÍA ser declarado "
                "nulo y dar lugar a reintegro o indemnización."
            ),
            "fuente": f"{FUENTE} — procedimiento mínimo previo al despido con justa causa",
            "sugerencia": "Revisa con un profesional cómo documentar este procedimiento en tu reglamento interno y en la práctica.",
        }
    return None


def _r_cuota_aprendices(d):
    n = d.get("num_empleados")
    apr = d.get("num_aprendices")
    if isinstance(n, (int, float)) and n >= 15 and isinstance(apr, (int, float)):
        requeridos = int(n // 20)
        if apr < requeridos:
            return {
                "punto": "Posible incumplimiento de la cuota de aprendices",
                "observacion": (
                    f"Con {int(n)} empleados, la cuota orientativa sería de "
                    f"aproximadamente {requeridos} aprendiz(es) (1 por cada 20 "
                    f"trabajadores en empresas de 15+). Tienes {int(apr)}. La "
                    "diferencia PODRÍA implicar monetización ante el SENA."
                ),
                "fuente": f"{FUENTE} — cuota de aprendizaje y monetización (1,5 SMMLV por aprendiz no contratado)",
                "sugerencia": "Revisa con un profesional tu obligación exacta y la conveniencia de contratar vs. monetizar.",
            }
    return None


def _r_jornada_semanal(d):
    # CORREGIDO v16.9: compara contra el máximo VIGENTE según la fecha (44h hasta
    # 14-jul-2026, 42h después), no contra una constante de 42 que generaría
    # falsos positivos durante la transición.
    h = d.get("horas_semanales")
    if isinstance(h, (int, float)):
        maximo = jornada_maxima_vigente()
        if h > maximo:
            return {
                "punto": f"Jornada semanal pactada en {h} horas (máximo vigente: {maximo}h)",
                "observacion": (
                    f"La jornada ordinaria máxima se reduce de forma gradual: {maximo} "
                    f"horas semanales en la etapa vigente (44h hasta el 14-jul-2026, "
                    f"42h desde el 15-jul-2026). Una jornada de {h}h PODRÍA requerir "
                    f"ajuste o el pago de horas extra."
                ),
                "fuente": "Art. 161 CST (mod. Ley 2101 de 2021 y Ley 2466 de 2025) — reducción gradual de la jornada",
                "sugerencia": "Revisa con un profesional el ajuste de la jornada al máximo vigente o la liquidación de horas extra.",
            }
    return None


_REGLAS = [
    _r_fijo_sin_escrito,
    _r_fijo_excede_4_anios,
    _r_fijo_como_norma,
    _r_jornada_nocturna,
    _r_recargo_dominical,
    _r_procedimiento_disciplinario,
    _r_cuota_aprendices,
    _r_jornada_semanal,
]


def diagnose_rules(datos: dict) -> list:
    """
    MODO REGLAS. Aplica las reglas deterministas y devuelve la lista de alertas.
    No requiere API key. Nunca afirma estatus legal; solo señala con lenguaje
    condicional y cita la fuente.
    """
    datos = datos or {}
    alertas = []
    for regla in _REGLAS:
        try:
            a = regla(datos)
            if a:
                alertas.append(a)
        except Exception:
            # Una regla con datos inesperados nunca debe tumbar el diagnóstico.
            continue
    return alertas


# ─────────────────────────────────────────────────────────────────────────────
# MODO IA — complemento opcional para texto libre de contrato
# ─────────────────────────────────────────────────────────────────────────────

_IA_PROMPT = """Eres un asistente educativo que ayuda a microempresarios colombianos a IDENTIFICAR puntos de sus contratos laborales que PODRÍAN requerir revisión frente a la Ley 2466 de 2025 (Reforma Laboral).

REGLAS ABSOLUTAS QUE NO PUEDES VIOLAR:
1. NUNCA afirmes que algo es "ilegal", "inválido" o "incumple". Usa SIEMPRE lenguaje condicional: "podría", "te recomendamos revisar".
2. NUNCA des instrucciones sobre qué hacer en el caso concreto (no digas "despide así", "cambia esto"). Solo SEÑALA el punto y recomienda revisar con un profesional.
3. Cada observación DEBE citar la regla o artículo de la Ley 2466 que la motiva.
4. NO inventes cláusulas que no estén en el texto. Si el texto no permite evaluar un punto, omítelo.
5. NO eres abogado y no emites dictámenes.

Cambios clave de la Ley 2466 de 2025 que debes vigilar:
- El contrato a término indefinido es la regla general; las demás modalidades son excepción y deben justificarse.
- Contrato a término fijo: máximo 4 años incluidas prórrogas; debe constar por escrito (si no, se entiende indefinido).
- Jornada nocturna: inicia a las 7:00 p.m. (antes 9:00 p.m.) para TODAS las empresas sin excepción por tamaño, vigente desde el 25-dic-2025.
- Recargo dominical/festivo escalonado: 80% (jul-2025 a jun-2026), 90% (jul-2026 a jun-2027), 100% (desde jul-2027).
- Jornada ordinaria máxima: 44 horas semanales hasta el 14-jul-2026; 42 horas desde el 15-jul-2026.
- Procedimiento disciplinario mínimo previo al despido con justa causa; para empresas de menos de 10 empleados es simplificado (basta escuchar al trabajador).
- Cuota de aprendices: 1 por cada 20 trabajadores en empresas de 15+; monetización de 1,5 SMMLV por aprendiz no contratado.

Analiza el siguiente texto de contrato y responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional ni ```:
{{"alertas": [{{"punto": "...", "observacion": "... (condicional)", "fuente": "Ley 2466 de 2025 — ...", "sugerencia": "Revisa con un profesional ..."}}]}}
Si no detectas nada que señalar, responde {{"alertas": []}}.

TEXTO DEL CONTRATO:
---
{texto}
---"""


def diagnose_text_ai(texto_contrato: str) -> list:
    """
    MODO IA. Lee texto libre de un contrato y devuelve alertas adicionales.
    Devuelve [] si no hay API key/SDK o si algo falla (degradación elegante):
    el diagnóstico NUNCA depende de este modo para funcionar.
    """
    if not texto_contrato or not texto_contrato.strip():
        return []
    if not os.environ.get("ANTHROPIC_API_KEY", "").strip():
        return []
    try:
        import anthropic
    except ImportError:
        return []

    try:
        client = anthropic.Anthropic(
            api_key=os.environ.get("ANTHROPIC_API_KEY", ""),
            max_retries=2,
            timeout=60.0,
        )
        resp = client.messages.create(
            model=_MODEL,
            max_tokens=1500,
            messages=[{"role": "user", "content": _IA_PROMPT.format(texto=texto_contrato[:12000])}],
        )
        # Extraer texto de forma segura (puede haber bloques no-texto).
        out = ""
        for b in (resp.content or []):
            if getattr(b, "type", None) == "text":
                out += b.text
        out = out.strip().replace("```json", "").replace("```", "").strip()
        data = json.loads(out)
        alertas = data.get("alertas", [])
        # Saneamiento mínimo: cada alerta debe tener los campos esperados.
        limpias = []
        for a in alertas:
            if isinstance(a, dict) and a.get("punto") and a.get("observacion"):
                limpias.append({
                    "punto": str(a.get("punto", "")),
                    "observacion": str(a.get("observacion", "")),
                    "fuente": str(a.get("fuente", FUENTE)),
                    "sugerencia": str(a.get("sugerencia", "Revisa este punto con un profesional del derecho.")),
                })
        return limpias
    except Exception:
        return []


# ─────────────────────────────────────────────────────────────────────────────
# API PÚBLICA — lo que el resto de la app llama
# ─────────────────────────────────────────────────────────────────────────────

def run_diagnostic(datos: dict = None, texto_contrato: str = None) -> dict:
    """
    Punto de entrada único del diagnóstico Ley 2466.

    Args:
        datos: dict estructurado del contrato (ver claves arriba). Modo REGLAS.
        texto_contrato: texto libre opcional del contrato. Modo IA (si disponible).

    Returns dict:
        alertas (list): lista combinada de alertas (reglas + IA), deduplicada.
        total (int): número de alertas.
        disclaimer (str): aviso legal a mostrar SIEMPRE.
        fuente (str): norma de referencia.
        modo_ia_activo (bool): si el complemento IA se ejecutó.
        resumen (str): texto listo para mostrar en el chat.
    """
    alertas = diagnose_rules(datos or {})
    ia = diagnose_text_ai(texto_contrato or "")
    modo_ia = bool(ia)

    # Deduplicar por 'punto' (las reglas tienen prioridad sobre la IA).
    vistos = {a["punto"].lower() for a in alertas}
    for a in ia:
        if a["punto"].lower() not in vistos:
            alertas.append(a)
            vistos.add(a["punto"].lower())

    return {
        "alertas": alertas,
        "total": len(alertas),
        "disclaimer": DISCLAIMER,
        "fuente": FUENTE,
        "modo_ia_activo": modo_ia,
        "resumen": format_diagnostic(alertas, modo_ia),
    }


def format_diagnostic(alertas: list, modo_ia: bool = False) -> str:
    """
    Formatea el diagnóstico como texto seguro para mostrar en el chat.
    Siempre empieza y termina con disclaimer/recomendación de revisión.
    """
    lines = []
    lines.append("📋 DIAGNÓSTICO DE CONTRATO — Reforma Laboral (Ley 2466 de 2025)")
    lines.append("")
    lines.append(DISCLAIMER)
    lines.append("")

    if not alertas:
        lines.append(
            "No identifiqué puntos que señalar con la información proporcionada. "
            "Esto NO significa que el contrato cumpla la ley: solo que, con los "
            "datos dados, no se activó ninguna alerta. Te recomendamos una "
            "revisión profesional de todas formas."
        )
        lines.append("")
        lines.append(f"Referencia: {FUENTE}")
        return "\n".join(lines)

    lines.append(f"Encontré {len(alertas)} punto(s) que PODRÍAN requerir revisión:")
    lines.append("")
    for i, a in enumerate(alertas, 1):
        lines.append(f"{i}. {a['punto']}")
        lines.append(f"   • Observación: {a['observacion']}")
        lines.append(f"   • Por qué: {a['fuente']}")
        lines.append(f"   • Recomendación: {a['sugerencia']}")
        lines.append("")

    if modo_ia:
        lines.append(
            "(Algunas observaciones provienen del análisis del texto que pegaste; "
            "trátalas como señales para revisar, no como conclusiones.)"
        )
        lines.append("")
    lines.append(
        "Siguiente paso sugerido: lleva estos puntos a un abogado laboralista. "
        "Si quieres, puedo generarte una VERSIÓN DE REFERENCIA actualizada del "
        "contrato (marcada como borrador) para facilitar esa conversación."
    )
    return "\n".join(lines)
