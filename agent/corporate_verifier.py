"""
WEP AI — CorporateVerifier (v15.6.0)
═══════════════════════════════════════════════════════════════════════════

Módulo de verificación y actualización autónoma de datos corporativos
en tiempo real, específicamente para acuerdos de socios, estatutos
sociales y documentos societarios.

PROBLEMA QUE RESUELVE
    Los datos de derecho de sociedades (ley vigente, reservas legales,
    mayorías estatutarias, centros de arbitraje) en `corporate_config.py`
    son estáticos. Una reforma a la Ley General de Sociedades Mercantiles
    en México, o una modernización de la Ley 1258 en Colombia, dejaría
    los documentos generados con datos obsoletos hasta que un developer
    actualice el archivo manualmente.

SOLUCIÓN
    Antes de generar el acuerdo de socios, este módulo consulta a Claude
    con `web_search` habilitado para verificar la vigencia y exactitud
    de los datos críticos del país detectado. Si encuentra cambios, los
    aplica al `cfg` antes de pasarlo al builder.

    Si falla (red caída, timeout, respuesta ambigua), el flujo continúa
    con el config estático sin interrupción — el verifier es NUNCA
    bloqueante.

ARQUITECTURA
    1. `generate_word` detecta el país → obtiene cfg base
    2. `verify_fiscal_data(cfg)`  → datos fiscales actualizados
    3. `verify_corporate_data(cfg)` → datos societarios actualizados ← este módulo
    4. `_build_shareholders_agreement(doc, ..., cfg, datos)` → documento correcto

CACHÉ POR SESIÓN
    Si el usuario genera 3 acuerdos de socios para México en una sesión,
    la búsqueda web ocurre solo la primera vez. Las siguientes usan el
    cache en memoria.

CAMPOS VERIFICADOS
    - ley_sociedades              → Ley de sociedades vigente
    - autoridad_societaria        → Organismo registrador
    - reserva_legal_pct           → % anual a reservar
    - reserva_legal_tope          → % tope de reserva legal
    - mayoria_simple              → Quórum para decisiones ordinarias
    - mayoria_calificada          → Quórum para decisiones extraordinarias
    - centro_arbitraje            → Centro de arbitraje recomendado
    - ley_arbitraje               → Ley de arbitraje aplicable

LOG DE CAMBIOS
    Los cambios detectados se registran en
    ~/.wepai/corporate_changes.log
    con campo, valor anterior, valor nuevo y fuente web.

MODELO USADO
    claude-haiku-4-5-20251001 — rápido y económico para verificación.

NOTA LEGAL
    El verifier nunca produce asesoría legal: solo verifica si los datos
    base de configuración siguen siendo correctos. El documento generado
    debe ser revisado por un abogado licenciado en la jurisdicción.
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.corporate_verifier")

# ── Session cache ─────────────────────────────────────────────────────────
_SESSION_CACHE: dict = {}

# ── Rutas ─────────────────────────────────────────────────────────────────
USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "corporate_changes.log")

# ── Campos críticos a verificar ───────────────────────────────────────────
VERIFY_FIELDS = [
    "ley_sociedades",
    "autoridad_societaria",
    "reserva_legal_pct",
    "reserva_legal_tope",
    "mayoria_simple",
    "mayoria_calificada",
    "mayoria_disolucion",
    "centro_arbitraje",
    "ley_arbitraje",
    "art_distribucion_utilidades",
]

# ── Prompt de verificación ────────────────────────────────────────────────
VERIFY_PROMPT = """You are a corporate law verification specialist for {country}.

Today's date is {today}. Search the web and verify whether the following stored
corporate law data for {country} is still accurate and up-to-date as of {year}.

Our stored values are:
- Corporations / companies law: {ley_sociedades}
- Corporate registry authority: {autoridad_societaria}
- Annual legal reserve %: {reserva_pct}
- Legal reserve cap (% of capital): {reserva_tope}
- Simple majority threshold: {mayoria_simple}
- Qualified majority threshold: {mayoria_calificada}
- Dissolution majority threshold: {mayoria_disolucion}
- Recommended arbitration center: {centro_arbitraje}
- Arbitration law: {ley_arbitraje}
- Profit distribution article: {art_distribucion}

Search for: "ley sociedades {country} {year}" AND "reserva legal sociedad anónima {country}" AND "centro arbitraje comercial {country}".

After searching, respond with ONLY a valid JSON object. Include ONLY the fields where you found CONFIRMED changes from official sources (government websites, official gazettes, bar association publications, supreme court rulings, registered chamber-of-commerce publications). If nothing has changed, return {{}}.

Be CONSERVATIVE: only report a change if you are HIGHLY CONFIDENT based on official primary sources. Do NOT report changes based on draft bills, proposals, or media speculation.

Response format (JSON only, no markdown, no preamble):
{{
  "ley_sociedades": "updated law name if changed",
  "autoridad_societaria": "updated authority if changed",
  "reserva_legal_pct": "updated % if changed",
  "reserva_legal_tope": "updated % if changed",
  "mayoria_simple": "updated threshold if changed",
  "mayoria_calificada": "updated threshold if changed",
  "mayoria_disolucion": "updated threshold if changed",
  "centro_arbitraje": "updated center if changed",
  "ley_arbitraje": "updated arbitration law if changed",
  "art_distribucion_utilidades": "updated article if changed",
  "_source": "URL or official source",
  "_verified_date": "YYYY-MM-DD"
}}"""


# ─── Public API ───────────────────────────────────────────────────────────

def verify_corporate_data(cfg: dict, force: bool = False) -> dict:
    """Verifica datos corporativos del cfg usando web_search en tiempo real.

    Args:
        cfg:   Config combinado (legal_config + corporate_config) con campo 'pais'.
        force: Si True, ignora cache y fuerza nueva búsqueda.

    Returns:
        cfg dict actualizado con los datos verificados.
        Si falla por cualquier razón, devuelve el cfg original sin modificar.

    Garantía: NUNCA lanza excepciones — el flujo de generación nunca se interrumpe.
    """
    pais = cfg.get("pais", "")

    # No verificar países placeholder o genéricos
    if not pais or pais == "[País]":
        return cfg

    country_code = _country_code_from_cfg(cfg)

    # Usar cache de sesión si existe
    if not force and country_code in _SESSION_CACHE:
        cached = _SESSION_CACHE[country_code]
        log.debug("CorporateVerifier: cache hit para %s", pais)
        return _apply_overrides(cfg, cached) if cached else cfg

    try:
        import anthropic
        client = anthropic.Anthropic(
            api_key=os.environ.get("ANTHROPIC_API_KEY", ""),
            max_retries=1,
            timeout=30.0,
        )

        today = datetime.now().strftime("%Y-%m-%d")
        year  = datetime.now().strftime("%Y")
        prompt = VERIFY_PROMPT.format(
            country              = pais,
            today                = today,
            year                 = year,
            ley_sociedades       = cfg.get("ley_sociedades", "—"),
            autoridad_societaria = cfg.get("autoridad_societaria", "—"),
            reserva_pct          = cfg.get("reserva_legal_pct", "—"),
            reserva_tope         = cfg.get("reserva_legal_tope", "—"),
            mayoria_simple       = cfg.get("mayoria_simple", "—"),
            mayoria_calificada   = cfg.get("mayoria_calificada", "—"),
            mayoria_disolucion   = cfg.get("mayoria_disolucion", "—"),
            centro_arbitraje     = cfg.get("centro_arbitraje", "—"),
            ley_arbitraje        = cfg.get("ley_arbitraje", "—"),
            art_distribucion     = cfg.get("art_distribucion_utilidades", "—"),
        )

        response = client.messages.create(
            model      = "claude-haiku-4-5-20251001",
            max_tokens = 800,
            tools      = [{"type": "web_search_20250305", "name": "web_search"}],
            messages   = [{"role": "user", "content": prompt}],
        )

        # Extraer texto de la respuesta
        result_text = ""
        for block in response.content:
            if hasattr(block, "type") and block.type == "text":
                result_text += block.text

        if not result_text.strip():
            log.warning("CorporateVerifier: respuesta vacía para %s", pais)
            _SESSION_CACHE[country_code] = {}
            return cfg

        # Limpiar markdown si Claude envolvió el JSON
        clean = result_text.strip()
        if "```" in clean:
            parts = clean.split("```")
            for part in parts:
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                if part.startswith("{"):
                    clean = part
                    break

        # Extraer JSON
        start = clean.find("{")
        end   = clean.rfind("}") + 1
        if start >= 0 and end > start:
            clean = clean[start:end]
        else:
            log.warning("CorporateVerifier: no se encontró JSON para %s", pais)
            _SESSION_CACHE[country_code] = {}
            return cfg

        raw_result = json.loads(clean)

        # Separar metadatos
        source        = raw_result.pop("_source", "web search")
        verified_date = raw_result.pop("_verified_date", today)

        # Filtrar campos válidos y no vacíos
        overrides = {
            k: v for k, v in raw_result.items()
            if k in VERIFY_FIELDS and v and str(v).strip()
        }

        _SESSION_CACHE[country_code] = overrides

        if overrides:
            _log_changes(country_code, pais, cfg, overrides, source, verified_date)
            try:
                from agent.legal_overrides import save_overrides
                save_overrides(country_code, overrides, source=source, category="corporate",
                               current_cfg=cfg, require_official=True)
            except Exception:
                pass
            log.info("CorporateVerifier: %d campo(s) actualizado(s) para %s — %s",
                     len(overrides), pais, list(overrides.keys()))
        else:
            log.debug("CorporateVerifier: sin cambios para %s", pais)

        return _apply_overrides(cfg, overrides) if overrides else cfg

    except json.JSONDecodeError as e:
        log.warning("CorporateVerifier: JSON inválido para %s: %s", pais, e)
        _SESSION_CACHE[country_code] = {}
        return cfg

    except Exception as e:
        log.warning("CorporateVerifier: error para %s: %s", pais, e)
        _SESSION_CACHE[country_code] = {}
        return cfg


def clear_session_cache():
    """Limpia el cache de sesión."""
    global _SESSION_CACHE
    _SESSION_CACHE = {}
    log.debug("CorporateVerifier: cache limpiado")


def get_cache_status() -> dict:
    """Retorna el estado actual del cache (para debugging y UI)."""
    return {
        code: bool(overrides)
        for code, overrides in _SESSION_CACHE.items()
    }


def get_changes_log() -> str:
    """Retorna el contenido del log de cambios detectados."""
    if not os.path.exists(CHANGES_LOG):
        return "Sin cambios corporativos registrados aún."
    with open(CHANGES_LOG, "r", encoding="utf-8") as f:
        return f.read()


# ── Helpers internos ──────────────────────────────────────────────────────

def _apply_overrides(cfg: dict, overrides: dict) -> dict:
    """Devuelve un nuevo cfg con los overrides aplicados. No muta el original."""
    if not overrides:
        return cfg
    updated = dict(cfg)
    updated.update(overrides)
    updated["_corporate_verified"]  = True
    updated["_corporate_overrides"] = list(overrides.keys())
    return updated


def _country_code_from_cfg(cfg: dict) -> str:
    """Obtiene el código ISO del país a partir del cfg."""
    try:
        from office.legal_config import LEGAL_CONFIG
        pais = cfg.get("pais", "")
        for code, entry in LEGAL_CONFIG.items():
            if entry.get("pais") == pais:
                return code
    except Exception:
        pass
    return "GENERIC"


def _log_changes(code: str, pais: str, original: dict,
                 changes: dict, source: str, date: str):
    """Registra los cambios detectados en el log de cambios corporativos."""
    try:
        os.makedirs(USER_DATA, exist_ok=True)
        with open(CHANGES_LOG, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\n")
            f.write(f"[{date}] CAMBIO CORPORATIVO — {pais} ({code})\n")
            f.write(f"{'='*60}\n")
            for field, new_val in changes.items():
                old_val = original.get(field, "—")
                f.write(f"  Campo   : {field}\n")
                f.write(f"  Anterior: {old_val}\n")
                f.write(f"  Actual  : {new_val}\n")
                f.write(f"  ---\n")
            f.write(f"  Fuente  : {source}\n")
            f.write(f"  Acción  : Actualizar corporate_config.py['{code}']\n")
    except Exception as e:
        log.warning("CorporateVerifier: no se pudo escribir log: %s", e)
