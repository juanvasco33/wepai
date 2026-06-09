"""
WEP AI — InventoryVerifier (v15.11.0)
═══════════════════════════════════════════════════════════════════════════

Verificador autónomo de método de valuación y normativa de inventarios por país.
Mantiene actualizado el método permitido (PEPS/UEPS/Promedio), tope de
deducibilidad, y reglas fiscales aplicables a inventarios.

USO: cfg = verify_inventory_data(cfg)
PATRÓN: idéntico a fiscal_verifier.py — NUNCA bloqueante.
"""

import os
import json
import logging
from datetime import datetime

log = logging.getLogger("wepai.inventory_verifier")
_SESSION_CACHE: dict = {}
USER_DATA   = os.path.expanduser("~/.wepai")
CHANGES_LOG = os.path.join(USER_DATA, "inventory_changes.log")

VERIFY_FIELDS = [
    "metodo_valuacion_inventario",   # PEPS/UEPS/Promedio/Identificación específica
    "metodos_permitidos_inventario", # Los que la ley permite (UEPS prohibido en NIIF)
    "norma_inventarios",             # NIC 2 / ASC 330 / Art. ISR
    "obligacion_inventario_fisico",  # Frecuencia mínima requerida
    "merma_deducible_pct",           # % de merma deducible si existe
    "stock_minimo_legal",            # Algunos sectores tienen obligación
]

VERIFY_PROMPT = """You are an inventory accounting and tax specialist for {country}.

Today: {today}. Verify stored inventory data for {country} as of {year}.

Stored values:
- Valuation method default: {met}
- Allowed methods: {permit}
- Inventory standard: {norma}
- Physical count obligation: {fisico}
- Deductible shrinkage %: {merma}
- Legal minimum stock: {stock}

Search for: "método valuación inventarios {country} {year}", "NIC 2 inventarios {country}",
"merma deducible {country}", "inventario físico obligatorio {country}".

Be CONSERVATIVE. Report only CONFIRMED changes from official primary sources.

Response (JSON only):
{{
  "metodo_valuacion_inventario": "method if changed",
  "metodos_permitidos_inventario": "list if changed",
  "norma_inventarios": "standard if changed",
  "obligacion_inventario_fisico": "frequency if changed",
  "merma_deducible_pct": "% if changed",
  "stock_minimo_legal": "yes/no/sector-specific if changed",
  "_source": "URL", "_verified_date": "YYYY-MM-DD"
}}"""


def verify_inventory_data(cfg: dict, force: bool = False) -> dict:
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
            met=cfg.get("metodo_valuacion_inventario", "—"),
            permit=cfg.get("metodos_permitidos_inventario", "—"),
            norma=cfg.get("norma_inventarios", "—"),
            fisico=cfg.get("obligacion_inventario_fisico", "—"),
            merma=cfg.get("merma_deducible_pct", "—"),
            stock=cfg.get("stock_minimo_legal", "—"),
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
                save_overrides(country_code, overrides, source=source, category="inventory",
                               current_cfg=cfg, require_official=True)
            except Exception:
                pass
        return _apply_overrides(cfg, overrides) if overrides else cfg
    except Exception as e:
        log.warning("InventoryVerifier error: %s", e)
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
    updated["_inventory_verified"] = True
    updated["_inventory_overrides"] = list(overrides.keys())
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
            f.write(f"\n{'='*60}\n[{date}] INVENTORY CHANGE — {pais} ({code})\n{'='*60}\n")
            for field, new_val in changes.items():
                f.write(f"  {field}: {original.get(field,'—')} → {new_val}\n")
            f.write(f"  Fuente: {source}\n")
    except Exception: pass
