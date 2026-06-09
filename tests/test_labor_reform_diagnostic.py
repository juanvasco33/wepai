"""
Tests del módulo de diagnóstico Ley 2466 (bache 1, v16.7).

Estos tests cubren el MODO REGLAS (determinista, sin API key). El MODO IA no se
testea aquí porque requiere red/API; el diseño garantiza que devuelve [] sin key,
lo cual se verifica indirectamente (run_diagnostic no falla sin key).
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from agent.labor_reform_diagnostic import (
    diagnose_rules,
    run_diagnostic,
    format_diagnostic,
    diagnose_text_ai,
    DISCLAIMER,
    FUENTE,
)


# ── Reglas individuales ───────────────────────────────────────────────────────

def test_fijo_sin_escrito_genera_alerta():
    alertas = diagnose_rules({"tipo_contrato": "fijo", "por_escrito": False})
    assert any("escrito" in a["punto"].lower() for a in alertas)


def test_fijo_excede_4_anios():
    alertas = diagnose_rules({"tipo_contrato": "fijo", "duracion_meses": 60})
    assert any("4 años" in a["punto"] or "superior a 4" in a["punto"] for a in alertas)


def test_fijo_dentro_de_4_anios_no_alerta_duracion():
    alertas = diagnose_rules({"tipo_contrato": "fijo", "duracion_meses": 24})
    assert not any("superior a 4" in a["punto"] for a in alertas)


def test_jornada_nocturna_9pm_mediana_empresa():
    alertas = diagnose_rules({
        "jornada_nocturna": True,
        "es_microempresa": False,
        "hora_inicio_nocturna": "21:00",
    })
    assert any("nocturna" in a["punto"].lower() for a in alertas)


def test_jornada_nocturna_microempresa_TAMBIEN_alerta():
    # CORREGIDO v16.9: la jornada nocturna a las 7pm aplica a TODAS las empresas.
    # Una microempresa que calcula desde las 9pm también debe recibir la alerta.
    alertas = diagnose_rules({
        "jornada_nocturna": True,
        "es_microempresa": True,
        "hora_inicio_nocturna": "21:00",
    })
    assert any("nocturna" in a["punto"].lower() for a in alertas)


def test_recargo_dominical_bajo():
    alertas = diagnose_rules({"recargo_dominical_pct": 75})
    assert any("dominical" in a["punto"].lower() or "festivo" in a["punto"].lower() for a in alertas)


def test_recargo_dominical_100_no_alerta():
    alertas = diagnose_rules({"recargo_dominical_pct": 100})
    assert not any("dominical" in a["punto"].lower() for a in alertas)


def test_sin_procedimiento_disciplinario():
    alertas = diagnose_rules({"tiene_procedimiento_disciplinario": False})
    assert any("disciplinario" in a["punto"].lower() for a in alertas)


def test_cuota_aprendices_incumplida():
    alertas = diagnose_rules({"num_empleados": 40, "num_aprendices": 0})
    assert any("aprendices" in a["punto"].lower() for a in alertas)


def test_cuota_aprendices_cumplida():
    alertas = diagnose_rules({"num_empleados": 40, "num_aprendices": 2})
    assert not any("aprendices" in a["punto"].lower() for a in alertas)


def test_jornada_semanal_excesiva():
    alertas = diagnose_rules({"horas_semanales": 48})
    assert any("jornada semanal" in a["punto"].lower() for a in alertas)


def test_jornada_42_no_alerta():
    alertas = diagnose_rules({"horas_semanales": 42})
    assert not any("jornada semanal" in a["punto"].lower() for a in alertas)


# ── Lógica escalonada por fecha (CORRECCIÓN v16.9) ────────────────────────────

def test_recargo_dominical_vigente_por_fecha():
    from datetime import date
    from agent.labor_reform_diagnostic import recargo_dominical_vigente
    assert recargo_dominical_vigente(date(2025, 3, 1)) == 75   # antes de la reforma
    assert recargo_dominical_vigente(date(2025, 8, 1)) == 80   # etapa 1
    assert recargo_dominical_vigente(date(2026, 8, 1)) == 90   # etapa 2
    assert recargo_dominical_vigente(date(2027, 8, 1)) == 100  # etapa 3


def test_jornada_maxima_vigente_por_fecha():
    from datetime import date
    from agent.labor_reform_diagnostic import jornada_maxima_vigente
    assert jornada_maxima_vigente(date(2026, 1, 1)) == 44   # antes del 15-jul-2026
    assert jornada_maxima_vigente(date(2026, 8, 1)) == 42   # después


def test_recargo_dominical_90_no_alerta_en_etapa_90():
    """Un 90% NO debe alertar si la etapa vigente es 90% (evita falso positivo)."""
    from datetime import date
    from agent.labor_reform_diagnostic import recargo_dominical_vigente
    vigente = recargo_dominical_vigente(date.today())
    # Pactar exactamente el vigente no debe disparar alerta.
    alertas = diagnose_rules({"recargo_dominical_pct": vigente})
    assert not any("dominical" in a["punto"].lower() for a in alertas)


def test_recargo_75_alerta_hoy():
    """Un 75% (régimen viejo) sí debe alertar en cualquier etapa de la reforma."""
    alertas = diagnose_rules({"recargo_dominical_pct": 75})
    assert any("dominical" in a["punto"].lower() for a in alertas)


def test_procedimiento_disciplinario_empresa_pequena():
    """Empresa <10 empleados: alerta de procedimiento simplificado, no el completo."""
    alertas = diagnose_rules({"tiene_procedimiento_disciplinario": False, "num_empleados": 5})
    disc = [a for a in alertas if "despido" in a["punto"].lower() or "sanción" in a["punto"].lower()]
    assert disc
    assert "escuchar" in disc[0]["observacion"].lower() or "pequeña" in disc[0]["punto"].lower()


# ── Garantías de seguridad (lenguaje, fuente, disclaimer) ─────────────────────

def test_toda_alerta_cita_fuente():
    alertas = diagnose_rules({
        "tipo_contrato": "fijo", "por_escrito": False, "duracion_meses": 60,
        "horas_semanales": 50,
    })
    assert alertas
    for a in alertas:
        assert a.get("fuente"), "Toda alerta debe citar una fuente normativa"
        assert a.get("sugerencia"), "Toda alerta debe recomendar revisión profesional"


def test_lenguaje_condicional_no_afirmativo():
    """Ninguna observación debe afirmar que algo es 'ilegal' o 'inválido'."""
    alertas = diagnose_rules({
        "tipo_contrato": "fijo", "por_escrito": False, "duracion_meses": 60,
        "jornada_nocturna": True, "es_microempresa": False, "hora_inicio_nocturna": "21:00",
        "recargo_dominical_pct": 75, "tiene_procedimiento_disciplinario": False,
        "horas_semanales": 50, "num_empleados": 40, "num_aprendices": 0,
    })
    prohibidas = ["es ilegal", "es inválido", "incumple la ley", "viola la ley"]
    for a in alertas:
        texto = (a["observacion"] + " " + a["punto"]).lower()
        for p in prohibidas:
            assert p not in texto, f"Lenguaje afirmativo prohibido detectado: '{p}'"
        # Debe contener lenguaje condicional
        assert "podría" in texto or "podria" in texto


# ── run_diagnostic (API pública) ──────────────────────────────────────────────

def test_run_diagnostic_estructura():
    res = run_diagnostic(datos={"tipo_contrato": "fijo", "por_escrito": False})
    assert "alertas" in res
    assert "disclaimer" in res and res["disclaimer"] == DISCLAIMER
    assert "fuente" in res and "2466" in res["fuente"]
    assert "resumen" in res
    assert res["total"] == len(res["alertas"])


def test_run_diagnostic_sin_api_key_no_falla():
    """Sin API key, el modo IA degrada a [] y el diagnóstico sigue funcionando."""
    key_backup = os.environ.pop("ANTHROPIC_API_KEY", None)
    try:
        res = run_diagnostic(
            datos={"tipo_contrato": "fijo", "duracion_meses": 60},
            texto_contrato="CONTRATO DE TRABAJO a término fijo por cinco años...",
        )
        assert res["modo_ia_activo"] is False
        assert res["total"] >= 1  # las reglas siguen funcionando
    finally:
        if key_backup is not None:
            os.environ["ANTHROPIC_API_KEY"] = key_backup


def test_diagnose_text_ai_sin_key_devuelve_vacio():
    key_backup = os.environ.pop("ANTHROPIC_API_KEY", None)
    try:
        assert diagnose_text_ai("cualquier texto") == []
    finally:
        if key_backup is not None:
            os.environ["ANTHROPIC_API_KEY"] = key_backup


def test_resumen_incluye_disclaimer_siempre():
    # Con alertas
    r1 = format_diagnostic(diagnose_rules({"horas_semanales": 50}))
    assert DISCLAIMER in r1
    # Sin alertas
    r2 = format_diagnostic([])
    assert DISCLAIMER in r2
    assert "2466" in r2


def test_sin_datos_no_falla():
    res = run_diagnostic(datos={})
    assert res["total"] == 0
    assert DISCLAIMER in res["resumen"]


# ── Config legal actualizada ──────────────────────────────────────────────────

def test_config_colombia_tiene_reforma():
    from office.legal_config import LEGAL_CONFIG
    co = LEGAL_CONFIG["CO"]
    assert "2466" in co.get("reforma_laboral", "")
    assert "42" in co.get("jornada", ""), "La jornada debe estar actualizada a 42h"
    assert co.get("contrato_regla_general"), "Debe declarar el indefinido como regla"
    assert co.get("procedimiento_disciplinario"), "Debe incluir el procedimiento disciplinario"
