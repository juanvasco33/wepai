"""
WEP AI — BankingVerifier (v15.11.0)
═══════════════════════════════════════════════════════════════════════════

Verificador autónomo de datos bancarios y financieros por país.
Mantiene actualizada la regulación bancaria, formato de cuenta interbancaria,
y tasas oficiales de tipo de cambio.

USO: cfg = verify_banking_data(cfg)
PATRÓN: idéntico a fiscal_verifier.py — NUNCA bloqueante.
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.banking_verifier")
_SESSION_CACHE: dict = {}
USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "banking_changes.log")

VERIFY_FIELDS = [
    "autoridad_bancaria",            # SUDEBAN/CNBV/SBS/SFC/SBIF/etc.
    "formato_cuenta_interbancaria",  # CLABE 18 dig, IBAN, Routing+Account
    "tipo_cambio_referencia",        # BANXICO, BCRP, BCRA, etc.
    "ley_blanqueo",                  # Ley antilavado
    "limite_efectivo_reportable",    # Monto que debe reportarse
    "comision_transferencia_max",    # Tope legal si existe
]

VERIFY_PROMPT = """You are a banking regulation specialist for {country}.

Today: {today}. Verify stored banking data for {country} as of {year}.

Stored values:
- Banking authority: {aut}
- Interbank account format: {formato}
- Reference exchange rate source: {tc}
- Anti-money laundering law: {aml}
- Cash transaction reporting threshold: {limit}
- Max transfer commission: {com}

Search for: "regulación bancaria {country} {year}", "ley antilavado {country}",
"BANXICO BCRP BCRA tipo cambio oficial".

Be CONSERVATIVE. Report only CONFIRMED changes from official primary sources.

Response (JSON only):
{{
  "autoridad_bancaria": "updated if changed",
  "formato_cuenta_interbancaria": "format if changed",
  "tipo_cambio_referencia": "source if changed",
  "ley_blanqueo": "updated law if changed",
  "limite_efectivo_reportable": "amount if changed",
  "comision_transferencia_max": "updated if changed",
  "_source": "URL", "_verified_date": "YYYY-MM-DD"
}}"""


def verify_banking_data(cfg: dict, force: bool = False) -> dict:
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
            aut=cfg.get("autoridad_bancaria", "—"),
            formato=cfg.get("formato_cuenta_interbancaria", "—"),
            tc=cfg.get("tipo_cambio_referencia", "—"),
            aml=cfg.get("ley_blanqueo", "—"),
            limit=cfg.get("limite_efectivo_reportable", "—"),
            com=cfg.get("comision_transferencia_max", "—"),
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
                save_overrides(country_code, overrides, source=source, category="banking",
                               current_cfg=cfg, require_official=True)
            except Exception:
                pass
        return _apply_overrides(cfg, overrides) if overrides else cfg
    except Exception as e:
        log.warning("BankingVerifier error: %s", e)
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
    updated["_banking_verified"] = True
    updated["_banking_overrides"] = list(overrides.keys())
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
            f.write(f"\n{'='*60}\n[{date}] BANKING CHANGE — {pais} ({code})\n{'='*60}\n")
            for field, new_val in changes.items():
                f.write(f"  {field}: {original.get(field,'—')} → {new_val}\n")
            f.write(f"  Fuente: {source}\n")
    except Exception: pass
