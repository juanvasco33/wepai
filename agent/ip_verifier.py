"""
WEP AI — IPVerifier (v15.8.0)
═══════════════════════════════════════════════════════════════════════════

Verificador autónomo de datos de propiedad intelectual y confidencialidad.
Mantiene actualizada la ley de PI, derechos morales, plazos de prescripción
de confidencialidad y autoridad registradora para cada país.

USO: cfg = verify_ip_data(cfg)
PATRÓN: idéntico a privacy_verifier.py — NUNCA bloqueante.
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.ip_verifier")
_SESSION_CACHE: dict = {}
USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "ip_changes.log")

VERIFY_FIELDS = [
    "ley_pi",                # Ley de propiedad intelectual vigente
    "autoridad_pi",          # Oficina registradora (IMPI, SIC, INDECOPI, INPI, etc.)
    "duracion_derechos_autor",  # Duración derechos patrimoniales
    "derechos_morales_irrenunciables",  # Si los derechos morales son irrenunciables
    "plazo_confidencialidad_default",   # Plazo legal/usual de confidencialidad
    "registro_obligatorio_obras",       # Si el registro es obligatorio
]

VERIFY_PROMPT = """You are an intellectual property law verification specialist for {country}.

Today: {today}. Verify the stored IP data for {country} as of {year}.

Stored values:
- IP law: {ley}
- IP authority: {autoridad}
- Copyright duration: {duracion}
- Moral rights inalienable: {morales}
- Default confidentiality period: {conf}
- Mandatory registration: {registro}

Search for: "ley propiedad intelectual {country} {year}", "oficina derechos autor {country}",
"NDA confidencialidad ley {country}".

Be CONSERVATIVE. Report only CONFIRMED changes from official primary sources.

Response (JSON only):
{{
  "ley_pi": "updated if changed",
  "autoridad_pi": "updated if changed",
  "duracion_derechos_autor": "updated if changed",
  "derechos_morales_irrenunciables": "yes/no if changed",
  "plazo_confidencialidad_default": "updated if changed",
  "registro_obligatorio_obras": "yes/no/declarative if changed",
  "_source": "URL", "_verified_date": "YYYY-MM-DD"
}}"""


def verify_ip_data(cfg: dict, force: bool = False) -> dict:
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
            ley=cfg.get("ley_pi", "—"),
            autoridad=cfg.get("autoridad_pi", "—"),
            duracion=cfg.get("duracion_derechos_autor", "—"),
            morales=cfg.get("derechos_morales_irrenunciables", "—"),
            conf=cfg.get("plazo_confidencialidad_default", "—"),
            registro=cfg.get("registro_obligatorio_obras", "—"),
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
                save_overrides(country_code, overrides, source=source, category="ip",
                               current_cfg=cfg, require_official=True)
            except Exception:
                pass
        return _apply_overrides(cfg, overrides) if overrides else cfg
    except Exception as e:
        log.warning("IPVerifier error: %s", e)
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
    updated["_ip_verified"] = True
    updated["_ip_overrides"] = list(overrides.keys())
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
            f.write(f"\n{'='*60}\n[{date}] IP CHANGE — {pais} ({code})\n{'='*60}\n")
            for field, new_val in changes.items():
                f.write(f"  {field}: {original.get(field,'—')} → {new_val}\n")
            f.write(f"  Fuente: {source}\n")
    except Exception: pass
