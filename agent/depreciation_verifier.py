"""
WEP AI — DepreciationVerifier (v15.11.0)
═══════════════════════════════════════════════════════════════════════════

Verificador autónomo de tablas de depreciación y amortización por país.
Mantiene actualizadas las tasas máximas autorizadas por categoría de
activo (MACRS US, ISR MX, NIIF, etc.).

USO: cfg = verify_depreciation_data(cfg)
PATRÓN: idéntico a fiscal_verifier.py — NUNCA bloqueante.
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.depreciation_verifier")
_SESSION_CACHE: dict = {}
USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "depreciation_changes.log")

VERIFY_FIELDS = [
    "depreciacion_inmuebles_pct",    # % anual edificios
    "depreciacion_maquinaria_pct",   # % anual maquinaria
    "depreciacion_equipo_pct",       # % anual equipo cómputo
    "depreciacion_vehiculos_pct",    # % anual vehículos
    "metodo_depreciacion_default",   # Línea recta / Acelerada / MACRS
    "vida_util_inmuebles_anos",      # Vida útil en años para edificios
    "vida_util_equipo_anos",         # Vida útil para equipo de cómputo
    "ley_depreciacion",              # Norma legal (ISR Art. 33-35 MX, MACRS US)
]

VERIFY_PROMPT = """You are a tax depreciation specialist for {country}.

Today: {today}. Verify stored depreciation rate data for {country} as of {year}.

Stored values:
- Buildings annual rate: {inm}%
- Machinery annual rate: {maq}%
- Computer equipment annual rate: {eq}%
- Vehicles annual rate: {veh}%
- Default method: {met}
- Useful life buildings (years): {vu_inm}
- Useful life equipment (years): {vu_eq}
- Depreciation law: {ley}

Search for: "tasa depreciación fiscal {country} {year}", "vida útil activos fijos {country}",
"MACRS modified accelerated cost recovery", "ISR depreciación México".

Be CONSERVATIVE. Report only CONFIRMED changes from official primary sources.

Response (JSON only):
{{
  "depreciacion_inmuebles_pct": "% if changed",
  "depreciacion_maquinaria_pct": "% if changed",
  "depreciacion_equipo_pct": "% if changed",
  "depreciacion_vehiculos_pct": "% if changed",
  "metodo_depreciacion_default": "method if changed",
  "vida_util_inmuebles_anos": "years if changed",
  "vida_util_equipo_anos": "years if changed",
  "ley_depreciacion": "law if changed",
  "_source": "URL", "_verified_date": "YYYY-MM-DD"
}}"""


def verify_depreciation_data(cfg: dict, force: bool = False) -> dict:
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
            inm=cfg.get("depreciacion_inmuebles_pct", "—"),
            maq=cfg.get("depreciacion_maquinaria_pct", "—"),
            eq=cfg.get("depreciacion_equipo_pct", "—"),
            veh=cfg.get("depreciacion_vehiculos_pct", "—"),
            met=cfg.get("metodo_depreciacion_default", "—"),
            vu_inm=cfg.get("vida_util_inmuebles_anos", "—"),
            vu_eq=cfg.get("vida_util_equipo_anos", "—"),
            ley=cfg.get("ley_depreciacion", "—"),
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
                save_overrides(country_code, overrides, source=source, category="depreciation",
                               current_cfg=cfg, require_official=True)
            except Exception:
                pass
        return _apply_overrides(cfg, overrides) if overrides else cfg
    except Exception as e:
        log.warning("DepreciationVerifier error: %s", e)
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
    updated["_depreciation_verified"] = True
    updated["_depreciation_overrides"] = list(overrides.keys())
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
            f.write(f"\n{'='*60}\n[{date}] DEPRECIATION CHANGE — {pais} ({code})\n{'='*60}\n")
            for field, new_val in changes.items():
                f.write(f"  {field}: {original.get(field,'—')} → {new_val}\n")
            f.write(f"  Fuente: {source}\n")
    except Exception: pass
