"""
WEP AI — PrivacyVerifier (v15.8.0)
═══════════════════════════════════════════════════════════════════════════

Verificador autónomo de datos de privacidad y protección de datos personales.
Mantiene actualizada la ley aplicable, autoridad supervisora, multas máximas
y plazos legales (ej. plazo de respuesta ARCO) para cada país.

USO
    cfg = verify_privacy_data(cfg)

CAMPOS VERIFICADOS
    - ley_privacidad        — Ley vigente (GDPR, LFPDPPP, LGPD, CCPA, etc.)
    - autoridad_privacidad  — Autoridad supervisora competente
    - plazo_arco            — Plazo legal para responder solicitudes ARCO
    - multa_maxima          — Multa máxima por incumplimiento
    - notificacion_brechas  — Plazo para notificar brechas de seguridad
    - dpo_obligatorio       — Si el DPO es obligatorio por ley

PATRÓN
    Igual que fiscal_verifier.py y corporate_verifier.py:
    - claude-haiku-4-5-20251001 con web_search
    - Cache de sesión
    - Log a ~/.wepai/privacy_changes.log
    - NUNCA bloqueante
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.privacy_verifier")

_SESSION_CACHE: dict = {}

USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "privacy_changes.log")

VERIFY_FIELDS = [
    "ley_privacidad",
    "autoridad_privacidad",
    "plazo_arco",
    "multa_maxima",
    "notificacion_brechas",
    "dpo_obligatorio",
]

VERIFY_PROMPT = """You are a data protection law verification specialist for {country}.

Today's date is {today}. Search the web and verify the following stored privacy
and data protection data for {country} is still accurate and up-to-date as of {year}.

Our stored values are:
- Privacy law: {ley}
- Supervisory authority: {autoridad}
- ARCO/Subject rights response period: {plazo}
- Maximum fine for violations: {multa}
- Breach notification deadline: {brechas}
- DPO mandatory: {dpo}

Search for: "ley protección datos {country} {year}", "reforma {country} GDPR LFPDPPP LGPD",
"autoridad protección datos {country}", "plazo derechos ARCO {country}".

Be CONSERVATIVE: report changes only when CONFIRMED by official primary sources
(government gazettes, official authority publications, supreme court rulings).
Do NOT report changes based on draft bills or media speculation.

Response format (JSON only, no markdown, no preamble):
{{
  "ley_privacidad": "updated law if changed",
  "autoridad_privacidad": "updated authority if changed",
  "plazo_arco": "updated period if changed",
  "multa_maxima": "updated max fine if changed",
  "notificacion_brechas": "updated breach notification if changed",
  "dpo_obligatorio": "yes/no/conditional if changed",
  "_source": "URL or official source",
  "_verified_date": "YYYY-MM-DD"
}}"""


def verify_privacy_data(cfg: dict, force: bool = False) -> dict:
    """Verifica datos de privacidad. NUNCA bloquea."""
    pais = cfg.get("pais", "")
    if not pais or pais == "[País]":
        return cfg

    country_code = _country_code_from_cfg(cfg)
    if not force and country_code in _SESSION_CACHE:
        cached = _SESSION_CACHE[country_code]
        return _apply_overrides(cfg, cached) if cached else cfg

    try:
        import anthropic
        client = anthropic.Anthropic(
            api_key=os.environ.get("ANTHROPIC_API_KEY", ""),
            max_retries=1, timeout=30.0,
        )
        today = datetime.now().strftime("%Y-%m-%d")
        year  = datetime.now().strftime("%Y")
        prompt = VERIFY_PROMPT.format(
            country=pais, today=today, year=year,
            ley=cfg.get("ley_privacidad", "—"),
            autoridad=cfg.get("autoridad_privacidad", "—"),
            plazo=cfg.get("plazo_arco", "—"),
            multa=cfg.get("multa_maxima", "—"),
            brechas=cfg.get("notificacion_brechas", "—"),
            dpo=cfg.get("dpo_obligatorio", "—"),
        )
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=600,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": prompt}],
        )
        result_text = ""
        for block in response.content:
            if hasattr(block, "type") and block.type == "text":
                result_text += block.text
        if not result_text.strip():
            _SESSION_CACHE[country_code] = {}
            return cfg
        clean = result_text.strip()
        if "```" in clean:
            for part in clean.split("```"):
                part = part.strip()
                if part.startswith("json"):
                    part = part[4:].strip()
                if part.startswith("{"):
                    clean = part
                    break
        start = clean.find("{"); end = clean.rfind("}") + 1
        if start < 0 or end <= start:
            _SESSION_CACHE[country_code] = {}
            return cfg
        raw = json.loads(clean[start:end])
        source = raw.pop("_source", "web search")
        verified = raw.pop("_verified_date", today)
        overrides = {k: v for k, v in raw.items()
                     if k in VERIFY_FIELDS and v and str(v).strip()}
        _SESSION_CACHE[country_code] = overrides
        if overrides:
            _log_changes(country_code, pais, cfg, overrides, source, verified)
            try:
                from agent.legal_overrides import save_overrides
                save_overrides(country_code, overrides, source=source, category="privacy",
                               current_cfg=cfg, require_official=True)
            except Exception:
                pass
        return _apply_overrides(cfg, overrides) if overrides else cfg
    except json.JSONDecodeError as e:
        log.warning("PrivacyVerifier JSON: %s", e)
        _SESSION_CACHE[country_code] = {}
        return cfg
    except Exception as e:
        log.warning("PrivacyVerifier error: %s", e)
        _SESSION_CACHE[country_code] = {}
        return cfg


def clear_session_cache():
    global _SESSION_CACHE
    _SESSION_CACHE = {}


def get_cache_status() -> dict:
    return {code: bool(o) for code, o in _SESSION_CACHE.items()}


def get_changes_log() -> str:
    if not os.path.exists(CHANGES_LOG):
        return "Sin cambios de privacidad registrados."
    with open(CHANGES_LOG, "r", encoding="utf-8") as f:
        return f.read()


def _apply_overrides(cfg: dict, overrides: dict) -> dict:
    if not overrides:
        return cfg
    updated = dict(cfg)
    updated.update(overrides)
    updated["_privacy_verified"] = True
    updated["_privacy_overrides"] = list(overrides.keys())
    return updated


def _country_code_from_cfg(cfg: dict) -> str:
    try:
        from office.legal_config import LEGAL_CONFIG
        pais = cfg.get("pais", "")
        for code, entry in LEGAL_CONFIG.items():
            if entry.get("pais") == pais:
                return code
    except Exception:
        pass
    return "GENERIC"


def _log_changes(code, pais, original, changes, source, date):
    try:
        os.makedirs(USER_DATA, exist_ok=True)
        with open(CHANGES_LOG, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\n[{date}] PRIVACY CHANGE — {pais} ({code})\n{'='*60}\n")
            for field, new_val in changes.items():
                f.write(f"  Campo: {field}\n  Anterior: {original.get(field,'—')}\n  Actual: {new_val}\n  ---\n")
            f.write(f"  Fuente: {source}\n")
    except Exception:
        pass
