"""
WEP AI — NotarialVerifier (v15.8.0)
═══════════════════════════════════════════════════════════════════════════

Verificador autónomo de formalidades notariales por país.
Mantiene actualizadas las formalidades de poderes, declaraciones juradas,
apostilla de La Haya y vigencia de poderes.

USO: cfg = verify_notarial_data(cfg)
PATRÓN: idéntico a privacy_verifier.py — NUNCA bloqueante.
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.notarial_verifier")
_SESSION_CACHE: dict = {}
USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "notarial_changes.log")

VERIFY_FIELDS = [
    "ley_notariado",             # Ley del notariado vigente
    "vigencia_poder_default",    # Vigencia por defecto si no se especifica
    "requisitos_apostilla",      # Si el país es parte del Convenio de La Haya
    "formato_declaracion_jurada",# Formato requerido (notarial / privada / mixta)
    "registro_poder",            # Si el poder debe registrarse en alguna oficina
    "testigos_requeridos",       # Cuántos testigos requiere
]

VERIFY_PROMPT = """You are a notarial law verification specialist for {country}.

Today: {today}. Verify the stored notarial data for {country} as of {year}.

Stored values:
- Notarial law: {ley}
- Default power of attorney term: {vigencia}
- Apostille requirements: {apostilla}
- Affidavit format required: {formato}
- Power registration: {registro}
- Witnesses required: {testigos}

Search for: "ley del notariado {country} {year}", "apostilla La Haya {country}",
"declaración jurada formato {country}", "vigencia poder notarial {country}".

Be CONSERVATIVE. Report only CONFIRMED changes from official primary sources.

Response (JSON only):
{{
  "ley_notariado": "updated if changed",
  "vigencia_poder_default": "updated period if changed",
  "requisitos_apostilla": "yes/no if changed",
  "formato_declaracion_jurada": "notarial/private/mixed if changed",
  "registro_poder": "yes/no/conditional if changed",
  "testigos_requeridos": "number if changed",
  "_source": "URL", "_verified_date": "YYYY-MM-DD"
}}"""


def verify_notarial_data(cfg: dict, force: bool = False) -> dict:
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
            ley=cfg.get("ley_notariado", "—"),
            vigencia=cfg.get("vigencia_poder_default", "—"),
            apostilla=cfg.get("requisitos_apostilla", "—"),
            formato=cfg.get("formato_declaracion_jurada", "—"),
            registro=cfg.get("registro_poder", "—"),
            testigos=cfg.get("testigos_requeridos", "—"),
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
                save_overrides(country_code, overrides, source=source, category="notarial",
                               current_cfg=cfg, require_official=True)
            except Exception:
                pass
        return _apply_overrides(cfg, overrides) if overrides else cfg
    except Exception as e:
        log.warning("NotarialVerifier error: %s", e)
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
    updated["_notarial_verified"] = True
    updated["_notarial_overrides"] = list(overrides.keys())
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
            f.write(f"\n{'='*60}\n[{date}] NOTARIAL CHANGE — {pais} ({code})\n{'='*60}\n")
            for field, new_val in changes.items():
                f.write(f"  {field}: {original.get(field,'—')} → {new_val}\n")
            f.write(f"  Fuente: {source}\n")
    except Exception: pass
