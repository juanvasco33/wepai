"""
Tests del mecanismo de alertas estructurales (v16.8).

Verifican: detección heurística (sin API key), persistencia, no-duplicación,
resolución, avisos al usuario y al desarrollador, y robustez (no lanza excepciones).
"""

import os
import sys
import json
import tempfile
import importlib

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _fresh_module(tmp_dir):
    """Carga structural_alerts apuntando su almacén a un dir temporal aislado."""
    os.environ["WEPAI_TEST_USER_DATA"] = tmp_dir
    import agent.structural_alerts as sa
    importlib.reload(sa)
    # Redirigir el almacén al tmp para no tocar ~/.wepai del entorno real.
    sa._USER_DATA = tmp_dir
    sa._ALERTS_PATH = os.path.join(tmp_dir, "structural_alerts.json")
    return sa


def test_heuristica_detecta_reforma():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        assert sa.looks_structural("ley_laboral", "Nueva reforma laboral Ley 2466 que modifica el artículo 5") is True


def test_heuristica_ignora_valor_simple():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        # Un cambio de valor puro (UVT a un nuevo monto) no debe verse estructural.
        assert sa.looks_structural("uvt", "52374") is False
        assert sa.looks_structural("iva", "19%") is False


def test_raise_y_get_pending():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        aid = sa.raise_alert("CO", "ley_laboral", "Reforma que crea nuevo procedimiento", source="https://mintrabajo.gov.co")
        assert aid
        pend = sa.get_pending_alerts()
        assert len(pend) == 1
        assert pend[0]["country"] == "CO"
        assert sa.has_pending_alerts() is True


def test_no_duplica_misma_alerta():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        a1 = sa.raise_alert("CO", "ley_laboral", "Reforma X", source="gov")
        a2 = sa.raise_alert("CO", "ley_laboral", "Reforma X", source="gov")
        assert a1 == a2
        assert len(sa.get_pending_alerts()) == 1
        # El contador de avistamientos sube
        assert sa.get_pending_alerts()[0]["seen_count"] == 2


def test_persistencia_entre_cargas():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        sa.raise_alert("CO", "jornada", "Nueva ley reduce la jornada y deroga el artículo previo")
        # Releer desde disco
        with open(sa._ALERTS_PATH, "r", encoding="utf-8") as f:
            store = json.load(f)
        assert len(store["alerts"]) == 1


def test_resolver_alerta():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        aid = sa.raise_alert("CO", "ley_laboral", "Reforma que sustitúyase el procedimiento")
        assert sa.has_pending_alerts() is True
        ok = sa.resolve_alert(aid, note="Implementado en v16.9")
        assert ok is True
        assert sa.has_pending_alerts() is False
        # Queda en historial pero resuelta
        assert len(sa.get_pending_alerts()) == 0


def test_detect_and_alert_estructural():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        # Sin API key, cae a heurística. "reforma" + "nuevo procedimiento" → estructural.
        os.environ.pop("ANTHROPIC_API_KEY", None)
        changes = {
            "ley_laboral": "Reforma laboral 2027 que crea un nuevo procedimiento de despido",
            "uvt": "55000",  # esto es valor puro → no alerta
            "_source": "https://mintrabajo.gov.co",
        }
        created = sa.detect_and_alert("CO", changes, source="https://mintrabajo.gov.co")
        assert len(created) == 1  # solo la reforma, no la uvt
        assert sa.has_pending_alerts() is True


def test_detect_and_alert_solo_valores_no_alerta():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        os.environ.pop("ANTHROPIC_API_KEY", None)
        changes = {"uvt": "55000", "iva": "19%", "salario_minimo": "1750905"}
        created = sa.detect_and_alert("CO", changes)
        assert created == []
        assert sa.has_pending_alerts() is False


def test_user_notice_suave():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        assert sa.user_notice() == ""  # sin alertas, sin aviso
        sa.raise_alert("CO", "ley_laboral", "Nueva reforma estructural")
        notice = sa.user_notice()
        assert notice  # hay aviso
        # El aviso es suave: no menciona "desarrollador" ni detalles técnicos
        assert "desarrollador" not in notice.lower()
        assert "programador" not in notice.lower()


def test_developer_report_detallado():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        assert "No hay alertas" in sa.developer_report()
        sa.raise_alert("CO", "ley_laboral", "Reforma X", detail="Actualizar _build_employment_contract")
        rep = sa.developer_report()
        assert "PENDIENTE" in rep.upper()
        assert "_build_employment_contract" in rep


def test_robustez_no_lanza_excepciones():
    with tempfile.TemporaryDirectory() as d:
        sa = _fresh_module(d)
        # Entradas raras no deben tumbar nada
        assert sa.detect_and_alert("CO", None) == []
        assert sa.detect_and_alert("CO", {}) == []
        assert sa.resolve_alert("inexistente") is False
        assert isinstance(sa.looks_structural("x", None), bool)
