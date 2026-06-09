"""
WEP AI — FiscalVerifier (v15.5.1)

Módulo de verificación y actualización autónoma de datos fiscales en tiempo real.

PROBLEMA QUE RESUELVE:
    Los datos fiscales (IVA, retenciones, leyes) en legal_config.py son estáticos.
    Si un gobierno cambia la tasa de IVA o la retención de honorarios, la app
    seguiría generando documentos con datos viejos hasta que un developer
    actualice el archivo manualmente.

SOLUCIÓN:
    Antes de generar cualquier documento, este módulo consulta a Claude con
    web_search habilitado para verificar los datos críticos del país detectado.
    Si encuentra cambios, los aplica al cfg antes de pasarlo al builder.
    Si falla (red caída, timeout, resultado ambiguo), el flujo continúa con
    el config estático sin interrupción — el verifier es NUNCA bloqueante.

ARQUITECTURA:
    1. generate_word / generate_excel detectan el país → obtienen cfg estático
    2. Llaman a verify_fiscal_data(cfg) → devuelve cfg verificado (o el mismo)
    3. El builder recibe el cfg ya actualizado → documento correcto

CACHÉ POR SESIÓN:
    Si el usuario genera 5 contratos para México en una sesión, la búsqueda
    web ocurre solo la primera vez. Las siguientes usan el cache en memoria.
    El cache se limpia al reiniciar la app (no persiste entre sesiones).

CAMPOS VERIFICADOS (los de mayor riesgo de cambio):
    - iva / iva_label       → Tasa e impuesto al valor agregado
    - retencion_honorarios  → Retención sobre honorarios profesionales
    - ley_servicios         → Ley de contratos de servicios vigente
    - autoridad_fiscal      → Nombre oficial de la autoridad fiscal

CAMPOS NO VERIFICADOS (estables, no cambian prácticamente):
    - simbolo / moneda      → Símbolo y código de moneda
    - id_empresa / id_persona → Tipo de identificador fiscal
    - idioma / pais         → Metadatos del país

LOG DE CAMBIOS:
    Si se detecta un cambio, se registra en ~/.wepai/fiscal_changes.log
    con el campo, valor anterior, valor nuevo y fuente web. Esto permite
    al developer actualizar legal_config.py en el siguiente release.

MODELO USADO:
    claude-haiku-4-5-20251001 — rápido y económico para verificación.
    No se usa Sonnet para no incrementar el costo por documento.
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.fiscal_verifier")

# ── Session cache ──────────────────────────────────────────────────────────────
# Estructura: { "MX": {"iva": "16%", ...} } — dict vacío = sin cambios
_SESSION_CACHE: dict = {}

# ── Rutas ──────────────────────────────────────────────────────────────────────
USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "fiscal_changes.log")

# ── Campos críticos a verificar ────────────────────────────────────────────────
VERIFY_FIELDS = [
    "iva",                  # tasa como string: "16%", "19%", "18%"
    "iva_label",            # "IVA", "IGV", "ITBMS", "ITBIS", "ISV"
    "retencion_honorarios", # texto con tasa y ley
    "ley_servicios",        # nombre completo de la ley de servicios
    "autoridad_fiscal",     # nombre oficial de la autoridad fiscal
    "moneda",               # código ISO de moneda (cambios monetarios: VE, AR)
    "simbolo",              # símbolo de la moneda ($, S/, Bs, ₡, etc.)
]

# ── Prompt de verificación ─────────────────────────────────────────────────────
VERIFY_PROMPT = """You are a fiscal and legal data verification specialist for {country}.

Search the web and verify the following current (as of today, {today}) fiscal and legal data for {country}.

Our stored values are:
- VAT/IVA/Tax rate: {iva} (label: {iva_label})
- Withholding on professional service fees: {retencion}
- Service contract governing law: {ley_servicios}
- Tax authority official name: {autoridad_fiscal}
- Official currency: {moneda} (symbol: {simbolo})

Search for: "tasa IVA {country} {year}" OR "VAT rate {country} {year}" AND "retención honorarios {country} {year}". Also verify the official currency has not changed (relevant for high-inflation countries with monetary reforms).

After searching, respond with ONLY a valid JSON object. Include ONLY fields where you found confirmed changes from our stored values. If nothing has changed, return {{}}.

Rules:
- Only report a change if you are HIGHLY CONFIDENT based on official sources (government websites, tax authority announcements, central bank announcements)
- CRITICAL: The "_source" MUST be a URL from an official government domain (e.g. dian.gov.co, mintrabajo.gov.co, a .gov.co / .gob.* domain, or the official tax authority site). If you cannot find an official government source confirming the change, return {{}} — do NOT report the change. A news article or blog is NOT an acceptable source for a legal/fiscal change.
- Do NOT report changes based on proposals, drafts, bills under discussion, or uncertain sources. Only report changes that are ALREADY IN FORCE.
- For iva: return the rate as a string like "19%" or "18%"
- For iva_label: return only the short label like "IVA", "IGV", "ITBMS"
- For retencion_honorarios: return the full descriptive string
- For ley_servicios: return the full official law name
- For autoridad_fiscal: return the official name only
- For moneda: return the ISO currency code (e.g. "VES", "ARS") ONLY if there was an official monetary reform
- For simbolo: return the currency symbol ONLY if it changed

Response format (JSON only, no markdown):
{{
  "iva": "new_rate_if_changed",
  "iva_label": "new_label_if_changed",
  "retencion_honorarios": "new_text_if_changed",
  "ley_servicios": "new_law_if_changed",
  "autoridad_fiscal": "new_name_if_changed",
  "moneda": "new_currency_code_if_changed",
  "simbolo": "new_symbol_if_changed",
  "_source": "URL or source confirming the change",
  "_verified_date": "YYYY-MM-DD"
}}"""


# ── Public API ─────────────────────────────────────────────────────────────────

def verify_fiscal_data(cfg: dict, force: bool = False) -> dict:
    """
    Verifica y actualiza los datos fiscales del cfg usando web_search en tiempo real.

    Args:
        cfg:   Config de país de legal_config.py (resultado de _detect_country)
        force: Si True, ignora el cache y fuerza una nueva búsqueda

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
        log.debug("FiscalVerifier: cache hit para %s", pais)
        return _apply_overrides(cfg, cached) if cached else cfg

    # v16.1 — POLÍTICA DE FRESCURA: si el overlay persistente ya tiene todos los
    # campos críticos verificados dentro del TTL (30 días), NO se vuelve a buscar
    # en la web. El cfg ya trae esos valores aplicados (apply_overlay se ejecuta
    # al resolver el país). Esto hace que la "actualización constante" sea eficiente:
    # se re-verifica solo cuando un dato vence, no en cada documento.
    if not force:
        try:
            from agent.legal_overrides import all_fresh
            if all_fresh(country_code, VERIFY_FIELDS):
                log.info("FiscalVerifier: overlay fresco para %s (dentro del TTL); "
                         "se omite búsqueda web.", pais)
                _SESSION_CACHE[country_code] = {}  # ya aplicado vía overlay
                return cfg
        except Exception as e:
            log.debug("FiscalVerifier: no se pudo evaluar frescura: %s", e)

    try:
        import anthropic
        client = anthropic.Anthropic(
            api_key=os.environ.get("ANTHROPIC_API_KEY", ""),
            max_retries=1,
            timeout=30.0,
        )

        today   = datetime.now().strftime("%Y-%m-%d")
        year    = datetime.now().strftime("%Y")
        prompt  = VERIFY_PROMPT.format(
            country          = pais,
            today            = today,
            year             = year,
            iva              = cfg.get("iva", "—"),
            iva_label        = cfg.get("iva_label", "IVA"),
            retencion        = cfg.get("retencion_honorarios", "—"),
            ley_servicios    = cfg.get("ley_servicios", "—"),
            autoridad_fiscal = cfg.get("autoridad_fiscal", "—"),
            moneda           = cfg.get("moneda", "—"),
            simbolo          = cfg.get("simbolo", "—"),
        )

        response = client.messages.create(
            model      = "claude-haiku-4-5-20251001",
            max_tokens = 600,
            tools      = [{"type": "web_search_20250305", "name": "web_search"}],
            messages   = [{"role": "user", "content": prompt}],
        )

        # Extraer texto de la respuesta (puede haber bloques tool_use intercalados)
        result_text = ""
        for block in response.content:
            if hasattr(block, "type") and block.type == "text":
                result_text += block.text

        if not result_text.strip():
            log.warning("FiscalVerifier: respuesta vacía para %s", pais)
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

        # Extraer solo el JSON si hay texto adicional
        start = clean.find("{")
        end   = clean.rfind("}") + 1
        if start >= 0 and end > start:
            clean = clean[start:end]
        else:
            log.warning("FiscalVerifier: no se encontró JSON en respuesta para %s", pais)
            _SESSION_CACHE[country_code] = {}
            return cfg

        raw_result = json.loads(clean)

        # Separar metadatos de campos reales
        source        = raw_result.pop("_source", "web search")
        verified_date = raw_result.pop("_verified_date", today)

        # Filtrar solo campos válidos y no vacíos
        overrides = {
            k: v for k, v in raw_result.items()
            if k in VERIFY_FIELDS and v and str(v).strip()
        }

        # Guardar en cache de sesión
        _SESSION_CACHE[country_code] = overrides

        # Registrar cambios si los hay
        if overrides:
            _log_changes(country_code, pais, cfg, overrides, source, verified_date)
            # v15.16.0 — PERSISTENCIA: guardar al overlay para que sobreviva reinicios
            try:
                from agent.legal_overrides import save_overrides
                save_overrides(country_code, overrides, source=source, category="fiscal",
                               current_cfg=cfg, require_official=True)
            except Exception as e:
                log.warning("FiscalVerifier: no se pudo persistir overlay: %s", e)
            # v16.8 — Detección de cambios estructurales también en lo fiscal
            # (una reforma tributaria puede cambiar reglas, no solo tasas).
            try:
                from agent.structural_alerts import detect_and_alert
                detect_and_alert(country_code, overrides, source=source)
            except Exception as e:
                log.warning("FiscalVerifier: detección estructural falló: %s", e)
            log.info("FiscalVerifier: %d campo(s) actualizado(s) para %s — %s",
                     len(overrides), pais, list(overrides.keys()))
        else:
            log.debug("FiscalVerifier: sin cambios para %s", pais)

        return _apply_overrides(cfg, overrides) if overrides else cfg

    except json.JSONDecodeError as e:
        log.warning("FiscalVerifier: JSON inválido para %s: %s", pais, e)
        _SESSION_CACHE[country_code] = {}
        return cfg

    except Exception as e:
        log.warning("FiscalVerifier: error para %s: %s", pais, e)
        _SESSION_CACHE[country_code] = {}
        return cfg


def clear_session_cache():
    """Limpia el cache de sesión. Llamar entre sesiones si es necesario."""
    global _SESSION_CACHE
    _SESSION_CACHE = {}
    log.debug("FiscalVerifier: cache limpiado")


def get_cache_status() -> dict:
    """Retorna el estado actual del cache (para debugging y UI)."""
    return {
        code: bool(overrides)
        for code, overrides in _SESSION_CACHE.items()
    }


def get_changes_log() -> str:
    """Retorna el contenido del log de cambios detectados."""
    if not os.path.exists(CHANGES_LOG):
        return "Sin cambios registrados aún."
    with open(CHANGES_LOG, "r", encoding="utf-8") as f:
        return f.read()


# ── Helpers internos ──────────────────────────────────────────────────────────

def _apply_overrides(cfg: dict, overrides: dict) -> dict:
    """Devuelve un nuevo cfg con los overrides aplicados. No muta el original."""
    if not overrides:
        return cfg
    updated = dict(cfg)
    updated.update(overrides)
    updated["_fiscal_verified"]    = True
    updated["_fiscal_overrides"]   = list(overrides.keys())
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
    """Registra los cambios detectados en el log de cambios fiscales."""
    try:
        os.makedirs(USER_DATA, exist_ok=True)
        with open(CHANGES_LOG, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\n")
            f.write(f"[{date}] CAMBIO DETECTADO — {pais} ({code})\n")
            f.write(f"{'='*60}\n")
            for field, new_val in changes.items():
                old_val = original.get(field, "—")
                f.write(f"  Campo   : {field}\n")
                f.write(f"  Anterior: {old_val}\n")
                f.write(f"  Actual  : {new_val}\n")
                f.write(f"  ---\n")
            f.write(f"  Fuente  : {source}\n")
            f.write(f"  Acción  : Actualizar legal_config.py['{code}']['{field}']\n")
    except Exception as e:
        log.warning("FiscalVerifier: no se pudo escribir log: %s", e)


# ═══════════════════════════════════════════════════════════════════════════════
# LABOR DATA VERIFIER — para contratos de trabajo
# Verifica artículos laborales, jornada, prestaciones y SS de cada país.
# ═══════════════════════════════════════════════════════════════════════════════

LABOR_VERIFY_FIELDS = [
    "ley_laboral",            # Nombre completo de la ley laboral vigente
    "jornada",                # Horas/semana y artículo aplicable
    "vacaciones",             # Días y artículo de vacaciones
    "aguinaldo",              # Aguinaldo / gratificación anual y artículo
    "gratificaciones",        # Beneficios adicionales legales
    "cts",                    # CTS / cesantías / indemnización
    "rescision_patron",       # Artículo de rescisión por el empleador
    "rescision_trabajador",   # Artículo de renuncia del trabajador
    "seguridad_social",       # Institución de seguridad social
    "tasa_ss_patron",         # Tasa de aportación patronal
    "pension",                # Sistema de pensiones vigente
    # v16.7 — Campos de la Reforma Laboral colombiana (Ley 2466/2025).
    # Permiten que la auto-verificación mantenga al día las REGLAS de la
    # reforma (no solo valores numéricos) cuando un nuevo decreto las cambie.
    "fijo_duracion_max",          # Duración máxima del contrato a término fijo
    "jornada_nocturna_inicio",    # Hora de inicio de la jornada nocturna
    "recargo_dominical",          # % de recargo dominical/festivo vigente
]

LABOR_VERIFY_PROMPT = """You are a labor law verification specialist for {country}.

Today's date is {today}. Search the web and verify whether the following current labor law data for {country} is still accurate and up-to-date.

Our stored values are:
- Labor law: {ley_laboral}
- Working hours / jornada: {jornada}
- Vacation entitlement: {vacaciones}
- Annual bonus (aguinaldo/gratificación): {aguinaldo}
- Additional legal benefits (gratificaciones): {gratificaciones}
- CTS / severance / cesantías: {cts}
- Employer dismissal article: {rescision_patron}
- Employee resignation article: {rescision_trabajador}
- Social security institution: {seguridad_social}
- Employer SS contribution rate: {tasa_ss_patron}
- Pension system: {pension}

Search for: "ley laboral {country} {year}" AND "jornada laboral {country} {year}" AND "prestaciones laborales {country} {year}". For Colombia, also search "reforma laboral Colombia {year}", "Ley 2466 de 2025 jornada nocturna", and "recargo dominical {year}" to detect changes in the labor reform rules.

After searching, respond with ONLY a valid JSON object. Include ONLY the fields where you found CONFIRMED changes from official government or legal sources. If nothing has changed, return {{}}.

Be conservative: only report a change if you are HIGHLY CONFIDENT based on official sources (government websites, official gazette, labor ministry announcements). Do NOT report changes based on proposals or drafts.

Response format (JSON only, no markdown, no preamble):
{{
  "ley_laboral": "updated law name if changed",
  "jornada": "updated hours/article if changed",
  "vacaciones": "updated vacation entitlement if changed",
  "aguinaldo": "updated bonus info if changed",
  "gratificaciones": "updated if changed",
  "cts": "updated if changed",
  "rescision_patron": "updated article if changed",
  "rescision_trabajador": "updated article if changed",
  "seguridad_social": "updated institution if changed",
  "tasa_ss_patron": "updated rate if changed",
  "pension": "updated system if changed",
  "fijo_duracion_max": "updated max fixed-term duration if changed",
  "jornada_nocturna_inicio": "updated night-shift start time if changed",
  "recargo_dominical": "updated Sunday/holiday surcharge if changed",
  "_source": "URL or official source",
  "_verified_date": "YYYY-MM-DD"
}}"""

# Session cache separado para datos laborales
_LABOR_CACHE: dict = {}


def verify_labor_data(cfg: dict, force: bool = False) -> dict:
    """
    Verifica y actualiza los datos laborales del cfg usando web_search en tiempo real.

    Específico para contratos de trabajo: verifica jornada, vacaciones, aguinaldo,
    CTS, artículos de rescisión, seguridad social y sistema de pensiones.

    Args:
        cfg:   Config de país de legal_config.py
        force: Si True, ignora el cache y fuerza nueva búsqueda

    Returns:
        cfg dict actualizado. Si falla, devuelve el original sin modificar.
    Garantía: NUNCA lanza excepciones.
    """
    pais = cfg.get("pais", "")
    if not pais or pais == "[País]":
        return cfg

    country_code = _country_code_from_cfg(cfg)

    if not force and country_code in _LABOR_CACHE:
        cached = _LABOR_CACHE[country_code]
        log.debug("LaborVerifier: cache hit para %s", pais)
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

        prompt = LABOR_VERIFY_PROMPT.format(
            country          = pais,
            today            = today,
            year             = year,
            ley_laboral      = cfg.get("ley_laboral", "—"),
            jornada          = cfg.get("jornada", "—"),
            vacaciones       = cfg.get("vacaciones", "—"),
            aguinaldo        = cfg.get("aguinaldo", "—"),
            gratificaciones  = cfg.get("gratificaciones", "—"),
            cts              = cfg.get("cts", "—"),
            rescision_patron = cfg.get("rescision_patron", "—"),
            rescision_trabajador = cfg.get("rescision_trabajador", "—"),
            seguridad_social = cfg.get("seguridad_social", "—"),
            tasa_ss_patron   = cfg.get("tasa_ss_patron", "—"),
            pension          = cfg.get("pension", "—"),
        )

        response = client.messages.create(
            model      = "claude-haiku-4-5-20251001",
            max_tokens = 800,
            tools      = [{"type": "web_search_20250305", "name": "web_search"}],
            messages   = [{"role": "user", "content": prompt}],
        )

        result_text = ""
        for block in response.content:
            if hasattr(block, "type") and block.type == "text":
                result_text += block.text

        if not result_text.strip():
            _LABOR_CACHE[country_code] = {}
            return cfg

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

        start = clean.find("{")
        end   = clean.rfind("}") + 1
        if start >= 0 and end > start:
            clean = clean[start:end]
        else:
            _LABOR_CACHE[country_code] = {}
            return cfg

        raw_result = json.loads(clean)

        source        = raw_result.pop("_source", "web search")
        verified_date = raw_result.pop("_verified_date", today)

        overrides = {
            k: v for k, v in raw_result.items()
            if k in LABOR_VERIFY_FIELDS and v and str(v).strip()
        }

        _LABOR_CACHE[country_code] = overrides

        if overrides:
            _log_changes(country_code, pais, cfg, overrides, source, verified_date)
            # v15.16.0 — PERSISTENCIA permanente del overlay laboral
            try:
                from agent.legal_overrides import save_overrides
                save_overrides(country_code, overrides, source=source, category="labor",
                               current_cfg=cfg, require_official=True)
            except Exception as e:
                log.warning("LaborVerifier: no se pudo persistir overlay: %s", e)
            # v16.8 — DETECCIÓN DE CAMBIOS ESTRUCTURALES. Si alguno de los cambios
            # confirmados parece alterar reglas/estructura (no solo valores), se
            # levanta una alerta para el desarrollador en vez de aplicarla a ciegas.
            try:
                from agent.structural_alerts import detect_and_alert
                detect_and_alert(country_code, overrides, source=source)
            except Exception as e:
                log.warning("LaborVerifier: detección estructural falló: %s", e)
            log.info("LaborVerifier: %d campo(s) laboral(es) actualizado(s) para %s — %s",
                     len(overrides), pais, list(overrides.keys()))

        return _apply_overrides(cfg, overrides) if overrides else cfg

    except Exception as e:
        log.warning("LaborVerifier: error para %s: %s", pais, e)
        _LABOR_CACHE[country_code] = {}
        return cfg
