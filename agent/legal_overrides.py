"""
WEP AI — LegalOverrides (v15.16.0)

Capa de PERSISTENCIA permanente para los datos legales verificados por IA.

PROBLEMA QUE RESUELVE:
    Hasta v15.15.x, cuando un verifier detectaba un cambio (ej. el IVA subió),
    lo aplicaba SOLO al documento en curso y lo anotaba en un log de texto con
    la nota "actualizar legal_config.py manualmente". El cambio NO persistía:
    el siguiente usuario que pedía un documento del mismo país volvía a pagar
    el costo de la búsqueda web, y si nadie editaba el código a mano, el dato
    seguía desactualizado para siempre.

SOLUCIÓN:
    Un overlay JSON en ~/.wepai/legal_overrides.json que:
      • Los verifiers ESCRIBEN cada vez que confirman un cambio oficial.
      • legal_config.py LEE al cargar, aplicando el overlay sobre los valores
        base del código. Así los cambios sobreviven reinicios y benefician a
        todos los usuarios siguientes — la app "aprende" de forma permanente.

POR QUÉ UN OVERLAY JSON Y NO REESCRIBIR legal_config.py:
    Reescribir un .py con dicts literales de 20 países es frágil: un fallo a
    mitad de escritura corrompe el código y tumba toda la app. El overlay JSON
    es seguro (escritura atómica), auditable (cada cambio lleva fecha y fuente),
    reversible (borrar el archivo restaura los valores base del código) y no
    toca el código fuente.

ESTRUCTURA DEL OVERLAY:
    {
      "MX": {
        "iva": {"value": "16%", "source": "https://sat.gob.mx/...",
                "verified_date": "2026-05-28", "category": "fiscal"},
        ...
      },
      ...
    }

POLÍTICA DE FRESCURA (TTL):
    Cada override guarda su fecha de verificación. Un campo se considera
    "fresco" durante OVERRIDE_TTL_DAYS días. Pasado ese plazo, el verifier
    vuelve a buscarlo en la web para confirmar que sigue vigente. Esto evita
    que un dato verificado hace un año se asuma correcto indefinidamente.
"""

import os
import json
import tempfile
import logging
from datetime import datetime, timedelta

log = logging.getLogger("wepai.legal_overrides")

USER_DATA      = os.path.expanduser("~/.wepai")
OVERRIDES_FILE = os.path.join(USER_DATA, "legal_overrides.json")

# Días que un override se considera vigente antes de re-verificar en web.
OVERRIDE_TTL_DAYS = 30

# Cache en memoria del overlay para no leer el disco en cada llamada.
_OVERLAY_CACHE = None


# ── Lectura ──────────────────────────────────────────────────────────────────

def _load_raw() -> dict:
    """Carga el overlay completo desde disco. Devuelve {} si no existe/corrupto."""
    global _OVERLAY_CACHE
    if _OVERLAY_CACHE is not None:
        return _OVERLAY_CACHE
    try:
        if os.path.exists(OVERRIDES_FILE):
            with open(OVERRIDES_FILE, "r", encoding="utf-8") as f:
                _OVERLAY_CACHE = json.load(f)
        else:
            _OVERLAY_CACHE = {}
    except (json.JSONDecodeError, OSError) as e:
        log.warning("LegalOverrides: overlay ilegible (%s), usando vacío", e)
        _OVERLAY_CACHE = {}
    return _OVERLAY_CACHE


def get_overrides(country_code: str) -> dict:
    """
    Devuelve {campo: valor} de los overrides guardados para un país.
    Solo los valores (no los metadatos), listos para aplicar sobre el cfg.
    """
    raw = _load_raw().get(country_code, {})
    return {field: meta["value"] for field, meta in raw.items()
            if isinstance(meta, dict) and "value" in meta}


def apply_overlay(country_code: str, cfg: dict) -> dict:
    """
    Devuelve un nuevo cfg con el overlay persistente aplicado encima.
    No muta el cfg original. Se llama desde legal_config al resolver un país.
    """
    overrides = get_overrides(country_code)
    if not overrides:
        return cfg
    updated = dict(cfg)
    updated.update(overrides)
    updated["_overlay_applied"] = list(overrides.keys())
    return updated


def is_fresh(country_code: str, field: str) -> bool:
    """
    True si el campo tiene un override verificado dentro del TTL.
    El verifier lo usa para decidir si re-buscar en web o confiar en lo guardado.
    """
    raw  = _load_raw().get(country_code, {})
    meta = raw.get(field)
    if not isinstance(meta, dict) or "verified_date" not in meta:
        return False
    try:
        verified = datetime.strptime(meta["verified_date"], "%Y-%m-%d")
        return datetime.now() - verified < timedelta(days=OVERRIDE_TTL_DAYS)
    except ValueError:
        return False


def all_fresh(country_code: str, fields: list) -> bool:
    """True si TODOS los campos dados están frescos (evita búsqueda web)."""
    return bool(fields) and all(is_fresh(country_code, f) for f in fields)


# ── Escritura (atómica) ──────────────────────────────────────────────────────

def save_overrides(country_code: str, changes: dict,
                   source: str = "web search", category: str = "general",
                   current_cfg: dict = None, require_official: bool = True) -> bool:
    """
    Persiste cambios confirmados al overlay. Escritura atómica (tmp + rename)
    para que un fallo nunca deje el archivo a medias.

    v16.1 — VALIDACIÓN DE SEGURIDAD: antes de persistir, cada cambio pasa por
    legal_validators.validate_changes(). Solo los campos que superan la
    validación de rango/formato/fuente se guardan. Los rechazados se registran
    para auditoría y NO contaminan el overlay. "Fallar seguro": ante la duda,
    se conserva el valor base conocido.

    Args:
        country_code:     código ISO ("MX", "CO", ...).
        changes:          {campo: nuevo_valor} confirmados por el verifier.
        source:           URL o fuente oficial que confirma el cambio.
        category:         "fiscal" | "labor" | "privacy" | ... (para auditoría).
        current_cfg:      config actual del país (para comparar valor viejo).
        require_official: si True, rechaza cambios sin fuente gubernamental.

    Returns:
        True si se guardó al menos un campo válido, False si nada pasó o falló.
    """
    global _OVERLAY_CACHE
    if not changes or not country_code:
        return False

    # ── v16.1 — Compuerta de validación ──────────────────────────────────────
    try:
        from agent.legal_validators import validate_changes
        aprobados, rechazados = validate_changes(
            changes, current_cfg=current_cfg, source=source,
            require_official=require_official,
        )
        if rechazados:
            for field, reason in rechazados.items():
                log.warning("LegalOverrides: RECHAZADO %s/%s — %s",
                            country_code, field, reason)
        if not aprobados:
            log.warning("LegalOverrides: ningún cambio superó la validación "
                        "para %s; no se persiste nada (fallar seguro).", country_code)
            return False
        changes = aprobados  # solo persistimos lo aprobado
    except Exception as e:
        # Si el validador mismo falla, NO persistimos (fallar seguro).
        log.error("LegalOverrides: validador no disponible (%s); "
                  "se rechaza el cambio por seguridad.", e)
        return False

    try:
        os.makedirs(USER_DATA, exist_ok=True)
        overlay = _load_raw()
        today   = datetime.now().strftime("%Y-%m-%d")

        country = overlay.setdefault(country_code, {})
        for field, value in changes.items():
            if value is None or not str(value).strip():
                continue
            country[field] = {
                "value":         value,
                "source":        source,
                "verified_date": today,
                "category":      category,
            }

        # Escritura atómica: escribir a tmp y renombrar (rename es atómico en POSIX).
        fd, tmp = tempfile.mkstemp(dir=USER_DATA, suffix=".tmp")
        try:
            with os.fdopen(fd, "w", encoding="utf-8") as f:
                json.dump(overlay, f, ensure_ascii=False, indent=2)
            os.replace(tmp, OVERRIDES_FILE)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

        _OVERLAY_CACHE = overlay  # refrescar cache
        log.info("LegalOverrides: %d campo(s) persistido(s) para %s (%s)",
                 len(changes), country_code, category)
        return True
    except Exception as e:
        log.warning("LegalOverrides: no se pudo persistir para %s: %s",
                    country_code, e)
        return False


# ── Utilidades de mantenimiento / UI ─────────────────────────────────────────

def get_overlay_status() -> dict:
    """
    Resumen del overlay para mostrar en UI/debugging:
    {país: {"campos": N, "ultima_verificacion": "YYYY-MM-DD"}}.
    """
    raw = _load_raw()
    status = {}
    for code, fields in raw.items():
        dates = [m.get("verified_date", "") for m in fields.values()
                 if isinstance(m, dict)]
        status[code] = {
            "campos": len(fields),
            "ultima_verificacion": max(dates) if dates else "—",
        }
    return status


def clear_overlay(country_code: str = None) -> bool:
    """
    Borra el overlay de un país (o todo si country_code es None).
    Restaura los valores base del código. Útil para revertir un dato erróneo.
    """
    global _OVERLAY_CACHE
    try:
        overlay = _load_raw()
        if country_code:
            overlay.pop(country_code, None)
        else:
            overlay = {}
        os.makedirs(USER_DATA, exist_ok=True)
        with open(OVERRIDES_FILE, "w", encoding="utf-8") as f:
            json.dump(overlay, f, ensure_ascii=False, indent=2)
        _OVERLAY_CACHE = overlay
        return True
    except Exception as e:
        log.warning("LegalOverrides: no se pudo limpiar: %s", e)
        return False


def reload_cache():
    """Fuerza recarga del overlay desde disco (tras edición externa)."""
    global _OVERLAY_CACHE
    _OVERLAY_CACHE = None
    _load_raw()


# ── Diagnóstico del estado de la verificación autónoma ───────────────────────

def verification_status() -> dict:
    """
    v15.16.0 — Reporta si la auto-actualización por IA está OPERATIVA.

    La verificación autónoma requiere DOS condiciones: el SDK `anthropic`
    instalado y una `ANTHROPIC_API_KEY` configurada. Si falta cualquiera, la
    app sigue funcionando pero genera con datos estáticos + overlay guardado,
    SIN buscar cambios nuevos en la web. La UI debe mostrar esto al usuario en
    vez de fallar en silencio.

    Returns dict:
        operativa (bool):  True si puede buscar cambios en web ahora mismo.
        sdk (bool):        SDK anthropic disponible.
        api_key (bool):    API key configurada.
        mensaje (str):     texto listo para mostrar al usuario.
    """
    sdk = False
    try:
        import anthropic  # noqa: F401
        sdk = True
    except ImportError:
        pass

    api_key = bool(os.environ.get("ANTHROPIC_API_KEY", "").strip())
    operativa = sdk and api_key

    if operativa:
        mensaje = ("✅ Verificación legal autónoma ACTIVA: los datos se "
                   "confirman en tiempo real y los cambios se guardan "
                   "permanentemente por país.")
    elif not sdk and not api_key:
        mensaje = ("⚠️ Verificación legal autónoma INACTIVA (falta el paquete "
                   "'anthropic' y la API key). Los documentos usan los datos "
                   "legales guardados, que pueden no reflejar cambios recientes.")
    elif not api_key:
        mensaje = ("⚠️ Verificación legal autónoma INACTIVA (falta "
                   "ANTHROPIC_API_KEY). Los documentos usan los datos legales "
                   "guardados, que pueden no reflejar cambios recientes.")
    else:
        mensaje = ("⚠️ Verificación legal autónoma INACTIVA (falta el paquete "
                   "'anthropic'). Instálalo con: pip install anthropic")

    return {
        "operativa": operativa,
        "sdk":       sdk,
        "api_key":   api_key,
        "mensaje":   mensaje,
    }
