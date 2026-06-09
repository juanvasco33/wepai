"""
Tests del validador de datos legales (v16.1).
Verifica que los datos peligrosos NUNCA se persistan.
"""
import os, sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from agent.legal_validators import (
    validate_field, validate_changes, is_official_source, _pct_to_float,
)

OFICIAL = "https://www.dian.gov.co/normatividad/iva-2026"
NO_OFICIAL = "https://algunblog.com/impuestos-colombia"


# ── Conversión de porcentajes ────────────────────────────────────────────────
def test_pct_parsing():
    assert _pct_to_float("19%") == 19.0
    assert _pct_to_float("19") == 19.0
    assert _pct_to_float("0.19") == 19.0   # decimal normalizado
    assert _pct_to_float("basura") is None


# ── IVA: rangos y saltos ─────────────────────────────────────────────────────
def test_iva_valido():
    ok, _ = validate_field("iva", "19%", "16%")
    assert ok

def test_iva_fuera_de_rango():
    ok, reason = validate_field("iva", "50%", "19%")
    assert not ok and "rango" in reason.lower()

def test_iva_decimal_mal_interpretado_se_rechaza_si_es_igual():
    # "0.19" → 19% == valor actual 19% → no es cambio
    ok, reason = validate_field("iva", "0.19", "19%")
    assert not ok and "idéntico" in reason.lower()

def test_iva_salto_sospechoso():
    ok, reason = validate_field("iva", "2%", "19%")
    assert not ok and "sospechoso" in reason.lower()

def test_iva_no_numerico():
    ok, reason = validate_field("iva", "diecinueve", "19%")
    assert not ok


# ── Moneda ───────────────────────────────────────────────────────────────────
def test_moneda_valida():
    ok, _ = validate_field("moneda", "COP", "COP")
    assert ok

def test_moneda_invalida():
    ok, reason = validate_field("moneda", "XYZ", "COP")
    assert not ok and "iso" in reason.lower()


# ── Etiqueta IVA ─────────────────────────────────────────────────────────────
def test_vat_label_valida():
    ok, _ = validate_field("iva_label", "IVA", "IVA")
    assert ok

def test_vat_label_invalida():
    ok, _ = validate_field("iva_label", "IMPUESTOX", "IVA")
    assert not ok


# ── Campos de texto ──────────────────────────────────────────────────────────
def test_texto_demasiado_corto():
    ok, _ = validate_field("autoridad_fiscal", "X", "DIAN")
    assert not ok

def test_texto_no_informativo():
    ok, _ = validate_field("ley_servicios", "N/A", "Código Civil")
    assert not ok

def test_texto_valido():
    ok, _ = validate_field("autoridad_fiscal", "DIAN — Dirección de Impuestos", "DIAN")
    assert ok


# ── Fuente oficial ───────────────────────────────────────────────────────────
def test_fuente_oficial_gov_co():
    assert is_official_source(OFICIAL)
    assert is_official_source("dian.gov.co")

def test_fuente_no_oficial():
    assert not is_official_source(NO_OFICIAL)
    assert not is_official_source("")
    assert not is_official_source("web search")


# ── Integración: validate_changes (la compuerta real) ────────────────────────
def test_compuerta_rechaza_todo_si_fuente_no_oficial():
    cambios = {"iva": "19%"}
    aprob, rech = validate_changes(cambios, {"iva": "16%"}, source=NO_OFICIAL)
    assert aprob == {} and "iva" in rech

def test_compuerta_acepta_cambio_valido_con_fuente_oficial():
    cambios = {"iva": "19%"}
    aprob, rech = validate_changes(cambios, {"iva": "16%"}, source=OFICIAL)
    assert aprob == {"iva": "19%"} and rech == {}

def test_compuerta_mezcla_aprobados_y_rechazados():
    # iva válido, moneda basura → uno pasa, otro no
    cambios = {"iva": "19%", "moneda": "XYZ"}
    aprob, rech = validate_changes(cambios, {"iva": "16%", "moneda": "COP"}, source=OFICIAL)
    assert "iva" in aprob and "moneda" in rech

def test_compuerta_iva_peligroso_nunca_pasa():
    # El caso que más preocupa: un IVA absurdo con fuente oficial igual se rechaza
    cambios = {"iva": "99%"}
    aprob, rech = validate_changes(cambios, {"iva": "19%"}, source=OFICIAL)
    assert aprob == {} and "iva" in rech


# ── Integración con save_overrides (fallar seguro) ───────────────────────────
def test_save_overrides_rechaza_dato_peligroso(tmp_path, monkeypatch):
    import agent.legal_overrides as lo
    monkeypatch.setattr(lo, "USER_DATA", str(tmp_path))
    monkeypatch.setattr(lo, "OVERRIDES_FILE", str(tmp_path / "legal_overrides.json"))
    monkeypatch.setattr(lo, "_OVERLAY_CACHE", None)
    # IVA absurdo → no debe persistir
    saved = lo.save_overrides("CO", {"iva": "80%"}, source=OFICIAL,
                              current_cfg={"iva": "19%"}, require_official=True)
    assert saved is False

def test_save_overrides_acepta_dato_valido(tmp_path, monkeypatch):
    import agent.legal_overrides as lo
    monkeypatch.setattr(lo, "USER_DATA", str(tmp_path))
    monkeypatch.setattr(lo, "OVERRIDES_FILE", str(tmp_path / "legal_overrides.json"))
    monkeypatch.setattr(lo, "_OVERLAY_CACHE", None)
    saved = lo.save_overrides("CO", {"iva": "21%"}, source=OFICIAL,
                              current_cfg={"iva": "19%"}, require_official=True)
    assert saved is True

def test_save_overrides_rechaza_fuente_no_oficial(tmp_path, monkeypatch):
    import agent.legal_overrides as lo
    monkeypatch.setattr(lo, "USER_DATA", str(tmp_path))
    monkeypatch.setattr(lo, "OVERRIDES_FILE", str(tmp_path / "legal_overrides.json"))
    monkeypatch.setattr(lo, "_OVERLAY_CACHE", None)
    saved = lo.save_overrides("CO", {"iva": "21%"}, source=NO_OFICIAL,
                              current_cfg={"iva": "19%"}, require_official=True)
    assert saved is False


# ── Validación de porcentajes genéricos (_pct) v16.2 ─────────────────────────
def test_pct_field_valido():
    ok, _ = validate_field("depreciacion_inmuebles_pct", "5%", "5%")
    assert ok

def test_pct_field_fuera_de_rango():
    ok, reason = validate_field("reserva_legal_pct", "250%", "50%")
    assert not ok and "rango" in reason.lower()

def test_pct_field_negativo():
    ok, _ = validate_field("merma_deducible_pct", "-10%", "3%")
    assert not ok

def test_pct_field_no_numerico():
    ok, _ = validate_field("depreciacion_equipo_pct", "alto", "30%")
    assert not ok
