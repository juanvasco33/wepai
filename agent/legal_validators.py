"""
WEP AI — Validador de datos legales (v16.1)
============================================
Capa de SEGURIDAD que valida cada dato legal/fiscal ANTES de persistirlo en el
overlay. Su trabajo es impedir que un valor absurdo, malformado o sin respaldo
oficial contamine los documentos de todos los usuarios.

FILOSOFÍA — "fallar seguro":
    Es preferible NO actualizar (y seguir usando el valor base conocido) que
    persistir un dato dudoso. Un IVA viejo conocido es menos peligroso que un
    IVA nuevo inventado. Por eso, ante la duda, este módulo RECHAZA el cambio.

QUÉ VALIDA:
    1. RANGO / FORMATO   — el valor tiene una forma plausible (IVA 0-30%,
       moneda ISO válida, texto de ley con largo razonable, etc.).
    2. FUENTE OFICIAL    — el cambio viene de un dominio gubernamental confiable
       (.gov.co, dian, mintrabajo...), no de un blog o noticia suelta.
    3. COHERENCIA        — el valor no es idéntico al que ya teníamos (no tiene
       sentido "actualizar" a lo mismo) ni un salto absurdo.

Cada validación devuelve (es_valido, razon). La razón se registra para auditoría.
"""

import re
from urllib.parse import urlparse

# ── Rangos y reglas por campo ────────────────────────────────────────────────
# IVA: ningún país hispanohablante tiene IVA fuera de ~0-30%. Colombia=19%.
IVA_MIN_PCT = 0.0
IVA_MAX_PCT = 30.0

# Códigos ISO 4217 de las monedas de los países soportados (+ comunes).
VALID_CURRENCY_CODES = {
    "COP", "MXN", "PEN", "ARS", "CLP", "EUR", "USD", "BOB", "VES", "VED",
    "PYG", "UYU", "PAB", "GTQ", "HNL", "NIO", "CRC", "CUP", "DOP", "BRL",
}

# Símbolos de moneda plausibles (lista permisiva pero acotada).
VALID_CURRENCY_SYMBOLS = {"$", "S/", "S/.", "Bs", "Bs.", "₲", "B/.", "Q", "L",
                          "C$", "₡", "RD$", "€", "US$", "COP$"}

# Etiquetas de impuesto al valor agregado válidas en la región.
VALID_VAT_LABELS = {"IVA", "IGV", "ITBMS", "ITBIS", "ISV", "ITBM", "VAT"}

# Dominios oficiales confiables (gobiernos y autoridades). El cambio de un dato
# legal solo se acepta si la fuente pertenece a uno de estos o termina en .gov.*
OFFICIAL_DOMAIN_HINTS = (
    ".gov.co", ".gob.mx", ".gob.pe", ".gob.cl", ".gob.es", ".gob.ar",
    ".gov.ar", ".gob.bo", ".gob.ec", ".gub.uy", ".gov.py", ".gob.ve",
    ".gob.pa", ".gob.gt", ".gob.hn", ".gob.ni", ".gob.sv", ".go.cr",
    ".gob.cu", ".gov.do", ".gov", ".gob", ".edu.co",
    # Autoridades específicas (por si el dominio no lleva .gov pero es oficial)
    "dian.gov.co", "mintrabajo.gov.co", "funcionpublica.gov.co",
    "sat.gob.mx", "sunat.gob.pe", "afip.gob.ar", "sii.cl", "aeat.es",
)


def _pct_to_float(value) -> float | None:
    """Convierte '19%' o '19' o '0.19' a 19.0. None si no parsea."""
    s = str(value).strip().replace("%", "").replace(",", ".")
    try:
        f = float(s)
    except ValueError:
        return None
    # Si viene como decimal (0.19) lo normalizamos a porcentaje.
    if 0 < f < 1:
        f *= 100
    return f


def is_official_source(source: str) -> bool:
    """True si la fuente parece un dominio gubernamental/oficial confiable."""
    if not source:
        return False
    s = source.lower().strip()
    # Aceptar si menciona explícitamente un dominio oficial conocido.
    if any(hint in s for hint in OFFICIAL_DOMAIN_HINTS):
        return True
    # Intentar parsear como URL y revisar el host.
    try:
        host = urlparse(s if "//" in s else "https://" + s).netloc.lower()
    except Exception:
        return False
    return any(hint in host for hint in OFFICIAL_DOMAIN_HINTS)


# ── Validadores por campo ────────────────────────────────────────────────────
def _validate_iva(new_value, old_value) -> tuple[bool, str]:
    f = _pct_to_float(new_value)
    if f is None:
        return False, f"IVA no numérico: {new_value!r}"
    if not (IVA_MIN_PCT <= f <= IVA_MAX_PCT):
        return False, f"IVA fuera de rango plausible ({f}% no está en {IVA_MIN_PCT}-{IVA_MAX_PCT}%)"
    old_f = _pct_to_float(old_value)
    if old_f is not None:
        if abs(f - old_f) < 0.001:
            return False, "IVA idéntico al actual (no es un cambio)"
        # Un salto de más de 10 puntos porcentuales es altamente sospechoso.
        if abs(f - old_f) > 10:
            return False, f"Salto de IVA sospechoso ({old_f}% → {f}%); se rechaza por seguridad"
    return True, "ok"


def _validate_currency(new_value, old_value) -> tuple[bool, str]:
    code = str(new_value).strip().upper()
    if code not in VALID_CURRENCY_CODES:
        return False, f"Código de moneda no reconocido (ISO 4217): {new_value!r}"
    return True, "ok"


def _validate_symbol(new_value, old_value) -> tuple[bool, str]:
    sym = str(new_value).strip()
    if not sym or len(sym) > 5:
        return False, f"Símbolo de moneda implausible: {new_value!r}"
    return True, "ok"


def _validate_vat_label(new_value, old_value) -> tuple[bool, str]:
    label = str(new_value).strip().upper()
    if label not in VALID_VAT_LABELS:
        return False, f"Etiqueta de impuesto no reconocida en la región: {new_value!r}"
    return True, "ok"


def _validate_text_field(new_value, old_value) -> tuple[bool, str]:
    """Para ley_servicios, autoridad_fiscal, retencion_honorarios, etc."""
    s = str(new_value).strip()
    if len(s) < 3:
        return False, f"Texto demasiado corto para ser un dato legal válido: {new_value!r}"
    if len(s) > 300:
        return False, "Texto excesivamente largo; posible respuesta malformada"
    # Rechazar valores que claramente son ruido/errores del modelo.
    if s.lower() in ("n/a", "none", "null", "desconocido", "unknown", "—", "-"):
        return False, f"Valor no informativo: {new_value!r}"
    return True, "ok"


def _validate_percentage(new_value, old_value) -> tuple[bool, str]:
    """Para cualquier campo _pct: porcentaje plausible entre 0 y 100."""
    f = _pct_to_float(new_value)
    if f is None:
        return False, f"Porcentaje no numérico: {new_value!r}"
    if not (0.0 <= f <= 100.0):
        return False, f"Porcentaje fuera de rango (0-100%): {f}%"
    return True, "ok"


# Mapa campo → validador específico. Los no listados usan texto genérico.
_FIELD_VALIDATORS = {
    "iva": _validate_iva,
    "moneda": _validate_currency,
    "simbolo": _validate_symbol,
    "iva_label": _validate_vat_label,
    "retencion_honorarios": _validate_text_field,
    "ley_servicios": _validate_text_field,
    "autoridad_fiscal": _validate_text_field,
}


def validate_field(field: str, new_value, old_value=None) -> tuple[bool, str]:
    """
    Valida un campo individual. Devuelve (es_valido, razon).
    Campos _pct usan validación de porcentaje; los demás sin validador
    específico pasan por validación de texto genérica.
    """
    if field in _FIELD_VALIDATORS:
        validator = _FIELD_VALIDATORS[field]
    elif field.endswith("_pct") or field.endswith("_porcentaje"):
        validator = _validate_percentage
    else:
        validator = _validate_text_field
    try:
        return validator(new_value, old_value)
    except Exception as e:
        return False, f"Error validando {field}: {e}"


def validate_changes(changes: dict, current_cfg: dict | None = None,
                     source: str = "", require_official: bool = True) -> tuple[dict, dict]:
    """
    Valida un conjunto de cambios propuestos por el verificador.

    Args:
        changes:          {campo: nuevo_valor} propuestos.
        current_cfg:      config actual del país (para comparar contra el valor viejo).
        source:           fuente del cambio (para exigir que sea oficial).
        require_official: si True, rechaza TODOS los cambios cuya fuente no sea oficial.

    Returns:
        (aprobados, rechazados)
        - aprobados: {campo: valor} que pasaron todas las validaciones.
        - rechazados: {campo: razon} que NO pasaron (para auditoría/log).
    """
    current_cfg = current_cfg or {}
    aprobados, rechazados = {}, {}

    # Compuerta de fuente: si se exige fuente oficial y no lo es, se rechaza todo.
    if require_official and not is_official_source(source):
        for field in changes:
            rechazados[field] = f"Fuente no oficial o ausente ({source!r}); rechazado por seguridad"
        return aprobados, rechazados

    for field, new_value in changes.items():
        if new_value is None or not str(new_value).strip():
            rechazados[field] = "Valor vacío"
            continue
        old_value = current_cfg.get(field)
        ok, reason = validate_field(field, new_value, old_value)
        if ok:
            aprobados[field] = new_value
        else:
            rechazados[field] = reason

    return aprobados, rechazados
