"""
WEP AI — CommercialVerifier (v15.8.0)
═══════════════════════════════════════════════════════════════════════════

Verificador autónomo de datos comerciales y de agencia/distribución.
Mantiene actualizada la ley aplicable, indemnización por terminación,
exclusividad permitida y plazos de aviso por país.

USO: cfg = verify_commercial_data(cfg)
PATRÓN: idéntico a privacy_verifier.py — NUNCA bloqueante.
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.commercial_verifier")
_SESSION_CACHE: dict = {}
USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "commercial_changes.log")

VERIFY_FIELDS = [
    "ley_distribucion",            # Ley aplicable a distribución/agencia
    "indemnizacion_terminacion",   # Régimen de indemnización por terminación
    "exclusividad_permitida",      # Si la exclusividad es permitida y bajo qué condiciones
    "plazo_aviso_terminacion_distrib",  # Plazo legal de aviso
    "competencia_post_contractual",     # Si se permite restricción post-terminación
    "autoridad_competencia",            # Autoridad de competencia (COFECE, SIC, etc.)
]

VERIFY_PROMPT = """You are a commercial distribution and agency law verification specialist for {country}.

Today: {today}. Verify the stored commercial distribution data for {country} as of {year}.

Stored values:
- Distribution law: {ley}
- Termination indemnification regime: {indem}
- Exclusivity allowed: {exclu}
- Termination notice period: {aviso}
- Post-contractual non-compete: {post}
- Competition authority: {comp}

Search for: "ley agencia distribución {country} {year}",
"indemnización clientela {country}", "exclusividad comercial {country}".

Be CONSERVATIVE. Report only CONFIRMED changes from official primary sources.

Response (JSON only):
{{
  "ley_distribucion": "updated if changed",
  "indemnizacion_terminacion": "updated regime if changed",
  "exclusividad_permitida": "yes/no/conditional if changed",
  "plazo_aviso_terminacion_distrib": "updated period if changed",
  "competencia_post_contractual": "yes/no/limited if changed",
  "autoridad_competencia": "updated authority if changed",
  "_source": "URL", "_verified_date": "YYYY-MM-DD"
}}"""


def verify_commercial_data(cfg: dict, force: bool = False) -> dict:
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
            max_retries=1, timeout=30.0)
        today = datetime.now().strftime("%Y-%m-%d")
        year  = datetime.now().strftime("%Y")
        prompt = VERIFY_PROMPT.format(
            country=pais, today=today, year=year,
            ley=cfg.get("ley_distribucion", "—"),
            indem=cfg.get("indemnizacion_terminacion", "—"),
            exclu=cfg.get("exclusividad_permitida", "—"),
            aviso=cfg.get("plazo_aviso_terminacion_distrib", "—"),
            post=cfg.get("competencia_post_contractual", "—"),
            comp=cfg.get("autoridad_competencia", "—"),
        )
        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=600,
            tools=[{"type": "web_search_20250305", "name": "web_search"}],
            messages=[{"role": "user", "content": prompt}],
        )
        result_text = "".join(b.text for b in response.content
                              if hasattr(b, "type") and b.type == "text")
        if not result_text.strip():
            _SESSION_CACHE[country_code] = {}
            return cfg
        clean = result_text.strip()
        if "```" in clean:
            for part in clean.split("```"):
                part = part.strip()
                if part.startswith("json"): part = part[4:].strip()
                if part.startswith("{"): clean = part; break
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
                save_overrides(country_code, overrides, source=source, category="commercial",
                               current_cfg=cfg, require_official=True)
            except Exception:
                pass
        return _apply_overrides(cfg, overrides) if overrides else cfg
    except Exception as e:
        log.warning("CommercialVerifier error: %s", e)
        _SESSION_CACHE[country_code] = {}
        return cfg


def clear_session_cache():
    global _SESSION_CACHE; _SESSION_CACHE = {}


def get_cache_status() -> dict:
    return {c: bool(o) for c, o in _SESSION_CACHE.items()}


def _apply_overrides(cfg, overrides):
    if not overrides: return cfg
    updated = dict(cfg)
    updated.update(overrides)
    updated["_commercial_verified"] = True
    updated["_commercial_overrides"] = list(overrides.keys())
    return updated


def _country_code_from_cfg(cfg):
    try:
        from office.legal_config import LEGAL_CONFIG
        for code, entry in LEGAL_CONFIG.items():
            if entry.get("pais") == cfg.get("pais", ""):
                return code
    except Exception: pass
    return "GENERIC"


def _log_changes(code, pais, original, changes, source, date):
    try:
        os.makedirs(USER_DATA, exist_ok=True)
        with open(CHANGES_LOG, "a", encoding="utf-8") as f:
            f.write(f"\n{'='*60}\n[{date}] COMMERCIAL CHANGE — {pais} ({code})\n{'='*60}\n")
            for field, new_val in changes.items():
                f.write(f"  {field}: {original.get(field,'—')} → {new_val}\n")
            f.write(f"  Fuente: {source}\n")
    except Exception: pass
