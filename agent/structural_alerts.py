"""
WEP AI — StructuralAlerts (v16.8)

Mecanismo de DETECCIÓN Y NOTIFICACIÓN de cambios estructurales legales que están
FUERA DEL ALCANCE de la auto-actualización y requieren intervención de un
desarrollador.

═══════════════════════════════════════════════════════════════════════════════
EL PROBLEMA QUE RESUELVE
═══════════════════════════════════════════════════════════════════════════════

La auto-verificación (fiscal_verifier.py) puede actualizar VALORES sola: si la
UVT pasa de $49.799 a $52.374, lo detecta y lo cambia. Pero NO puede reescribir
la ESTRUCTURA de las plantillas cuando una ley cambia las reglas de fondo:
  • una reforma que añade cláusulas obligatorias a un contrato,
  • un decreto que crea un procedimiento nuevo,
  • un cambio que modifica qué documentos son válidos.

Aplicar eso automáticamente sería peligroso (la IA inventaría estructura). Lo
correcto es DETECTAR la señal y AVISAR a un humano (el desarrollador / dueño).

Este módulo es ese "detector de humo": no apaga el incendio, levanta la alerta.

═══════════════════════════════════════════════════════════════════════════════
CÓMO FUNCIONA
═══════════════════════════════════════════════════════════════════════════════

1. DETECCIÓN: dos vías.
   a) Heurística (sin API key): analiza los resultados de la auto-verificación
      buscando señales lingüísticas de cambio estructural (palabras como
      "reforma", "deroga", "nuevo procedimiento", "obligatorio a partir de...").
   b) IA (con API key): una pasada explícita que pregunta al modelo si un cambio
      detectado es de tipo VALOR (auto-actualizable) o ESTRUCTURAL (requiere
      programador).

2. NOTIFICACIÓN: las alertas estructurales se PERSISTEN en
   ~/.wepai/structural_alerts.json (escritura atómica) para que no se pierdan
   entre sesiones, y quedan disponibles para:
      • el desarrollador (vía get_pending_alerts / un panel de admin),
      • el usuario final (un aviso suave de que el contenido podría requerir
        actualización, sin alarmar).

3. RESOLUCIÓN: cuando el desarrollador implementa el cambio, marca la alerta como
   resuelta (resolve_alert) y deja de notificarse.

NINGUNA función de este módulo lanza excepciones hacia afuera: la detección de
alertas jamás debe tumbar la generación de un documento.
"""

import os
import json
import hashlib
import logging
from datetime import datetime

log = logging.getLogger("wepai.structural_alerts")

# Reutilizamos el directorio de datos del usuario que ya usa el resto de la app.
try:
    from config import USER_DATA as _USER_DATA
except Exception:
    _USER_DATA = os.path.expanduser("~/.wepai")

_ALERTS_PATH = os.path.join(_USER_DATA, "structural_alerts.json")

try:
    from config import DEFAULT_CHAT_MODEL as _MODEL
except Exception:
    _MODEL = "claude-sonnet-4-6"


# ─────────────────────────────────────────────────────────────────────────────
# Señales heurísticas de cambio ESTRUCTURAL (no requieren API key)
# ─────────────────────────────────────────────────────────────────────────────
# Si el texto de un cambio detectado contiene estas señales, probablemente no es
# un simple cambio de valor sino algo que altera reglas/estructura.
_STRUCTURAL_SIGNALS = [
    "reforma", "deroga", "deroga", "derogado", "deroga ", "deroga.",
    "nuevo procedimiento", "procedimiento mínimo", "nueva modalidad",
    "se elimina", "se suprime", "queda prohibido", "obligatorio a partir",
    "entra en vigencia", "entrará en vigencia", "nueva ley", "ley nueva",
    "modifica el artículo", "modifíquese", "adiciónese", "sustitúyase",
    "nueva cláusula", "cláusula obligatoria", "requisito nuevo",
    "decreto", "sentencia c-",  # sentencias de constitucionalidad cambian reglas
]

# Campos que, si cambian, son casi siempre VALOR (auto-actualizable) y NO deben
# disparar alerta estructural por sí solos.
_VALUE_ONLY_FIELDS = {
    "iva", "uvt", "salario_minimo", "smmlv", "tasa_ss_patron", "tasa_pension",
    "recargo_dominical",  # el % es valor; pero ojo si viene con texto de reforma
}


def _signature(country_code: str, field: str, summary: str) -> str:
    """ID estable de una alerta, para no duplicarla en cada verificación."""
    raw = f"{country_code}|{field}|{summary}".lower().encode("utf-8")
    return hashlib.sha256(raw).hexdigest()[:16]


def _load() -> dict:
    """Carga el almacén de alertas. Devuelve estructura vacía si no existe."""
    try:
        if os.path.exists(_ALERTS_PATH):
            with open(_ALERTS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
                if isinstance(data, dict) and "alerts" in data:
                    return data
    except Exception as e:
        log.warning("StructuralAlerts: almacén ilegible (%s), usando vacío", e)
    return {"alerts": {}}


def _save(store: dict) -> bool:
    """Escritura atómica (tmp + rename) para no corromper el archivo."""
    try:
        os.makedirs(_USER_DATA, exist_ok=True)
        tmp = _ALERTS_PATH + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(store, f, ensure_ascii=False, indent=2)
        os.replace(tmp, _ALERTS_PATH)
        return True
    except Exception as e:
        log.warning("StructuralAlerts: no se pudo guardar (%s)", e)
        return False


# ─────────────────────────────────────────────────────────────────────────────
# DETECCIÓN
# ─────────────────────────────────────────────────────────────────────────────

def looks_structural(field: str, summary: str) -> bool:
    """
    Heurística sin API key: ¿este cambio parece estructural?

    Un cambio es estructural si su descripción contiene señales de reforma /
    nuevo procedimiento / derogación, AÚN si el campo afectado es uno de valor
    (porque una reforma puede cambiar el % Y la regla detrás de él).
    """
    text = (summary or "").lower()
    has_signal = any(sig in text for sig in _STRUCTURAL_SIGNALS)
    if not has_signal:
        return False
    # Si NO hay señal de reforma y el campo es puramente de valor, no es
    # estructural. (Este caso ya retornó False arriba, pero lo dejamos explícito.)
    return True


def classify_change_ai(field: str, old_value: str, new_value: str, source: str) -> dict:
    """
    Clasificación con IA (opcional). Pregunta al modelo si un cambio confirmado
    es de tipo VALUE (auto-actualizable) o STRUCTURAL (requiere desarrollador).

    Degrada con elegancia: si no hay API key/SDK o algo falla, devuelve
    {"tipo": "unknown"} y el llamador cae a la heurística.
    """
    if not os.environ.get("ANTHROPIC_API_KEY", "").strip():
        return {"tipo": "unknown"}
    try:
        import anthropic
    except ImportError:
        return {"tipo": "unknown"}

    prompt = f"""Un sistema detectó un cambio legal en Colombia. Clasifícalo.

Campo: {field}
Valor anterior: {old_value}
Valor nuevo: {new_value}
Fuente: {source}

Responde ÚNICAMENTE con un JSON válido, sin texto adicional:
{{"tipo": "VALUE" o "STRUCTURAL", "razon": "explicación breve en español", "que_actualizar": "qué tendría que cambiar un programador si es STRUCTURAL, o vacío si es VALUE"}}

Definiciones:
- VALUE: solo cambió un número o monto (ej. la UVT, el salario mínimo, una tasa). El sistema puede actualizarlo solo.
- STRUCTURAL: cambió una REGLA, un PROCEDIMIENTO, una CLÁUSULA obligatoria, una modalidad de contrato, o algo que requiere reescribir la estructura de una plantilla. Requiere un programador."""

    try:
        client = anthropic.Anthropic(
            api_key=os.environ.get("ANTHROPIC_API_KEY", ""),
            max_retries=1, timeout=30.0,
        )
        resp = client.messages.create(
            model=_MODEL, max_tokens=400,
            messages=[{"role": "user", "content": prompt}],
        )
        out = ""
        for b in (resp.content or []):
            if getattr(b, "type", None) == "text":
                out += b.text
        out = out.strip().replace("```json", "").replace("```", "").strip()
        data = json.loads(out)
        if data.get("tipo") in ("VALUE", "STRUCTURAL"):
            return data
        return {"tipo": "unknown"}
    except Exception:
        return {"tipo": "unknown"}


# ─────────────────────────────────────────────────────────────────────────────
# REGISTRO Y NOTIFICACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def raise_alert(country_code: str, field: str, summary: str,
                source: str = "", detail: str = "",
                severity: str = "media") -> str:
    """
    Registra (o refresca) una alerta estructural persistente.

    Si ya existe una alerta con la misma firma y sigue pendiente, no la duplica:
    solo actualiza last_seen. Devuelve el id de la alerta.
    """
    try:
        store = _load()
        aid = _signature(country_code, field, summary)
        now = datetime.now().isoformat(timespec="seconds")
        existing = store["alerts"].get(aid)
        if existing and existing.get("status") == "pending":
            existing["last_seen"] = now
            existing["seen_count"] = existing.get("seen_count", 1) + 1
        else:
            store["alerts"][aid] = {
                "id": aid,
                "country": country_code,
                "field": field,
                "summary": summary,
                "detail": detail,
                "source": source,
                "severity": severity,
                "status": "pending",
                "created": now,
                "last_seen": now,
                "seen_count": 1,
            }
        _save(store)
        log.info("StructuralAlert registrada [%s] %s/%s", aid, country_code, field)
        return aid
    except Exception as e:
        log.warning("StructuralAlerts: raise_alert falló (%s)", e)
        return ""


def detect_and_alert(country_code: str, changes: dict, source: str = "") -> list:
    """
    Punto de entrada principal para la auto-verificación.

    Recibe el dict de cambios confirmados (campo -> nuevo valor / descripción)
    que produjo el verificador, decide cuáles son ESTRUCTURALES y registra una
    alerta por cada uno. Devuelve la lista de ids de alertas creadas/refrescadas.

    Estrategia: primero intenta clasificar con IA (precisa); si no está
    disponible, cae a la heurística lingüística (siempre disponible).
    """
    created = []
    if not changes:
        return created
    try:
        for field, value in changes.items():
            if field.startswith("_"):  # _source, _verified_date, etc.
                continue
            summary = str(value)
            # 1) Intento con IA
            cls = classify_change_ai(field, "(anterior)", summary, source)
            is_structural = False
            detail = ""
            if cls.get("tipo") == "STRUCTURAL":
                is_structural = True
                detail = cls.get("que_actualizar", "")
            elif cls.get("tipo") == "VALUE":
                is_structural = False
            else:
                # 2) Sin IA o resultado desconocido → heurística
                is_structural = looks_structural(field, summary)

            if is_structural:
                aid = raise_alert(
                    country_code=country_code,
                    field=field,
                    summary=summary,
                    source=source,
                    detail=detail or "Cambio estructural detectado: requiere revisión de un desarrollador.",
                    severity="alta",
                )
                if aid:
                    created.append(aid)
    except Exception as e:
        log.warning("StructuralAlerts: detect_and_alert falló (%s)", e)
    return created


# ─────────────────────────────────────────────────────────────────────────────
# CONSULTA Y RESOLUCIÓN (para el desarrollador / panel admin / UI)
# ─────────────────────────────────────────────────────────────────────────────

def get_pending_alerts() -> list:
    """Lista de alertas pendientes, más recientes primero."""
    store = _load()
    pend = [a for a in store["alerts"].values() if a.get("status") == "pending"]
    pend.sort(key=lambda a: a.get("created", ""), reverse=True)
    return pend


def has_pending_alerts() -> bool:
    """True si hay al menos una alerta estructural pendiente."""
    return any(a.get("status") == "pending" for a in _load()["alerts"].values())


def resolve_alert(alert_id: str, note: str = "") -> bool:
    """
    Marca una alerta como resuelta (el desarrollador ya implementó el cambio).
    Deja de notificarse pero queda en el historial para auditoría.
    """
    try:
        store = _load()
        a = store["alerts"].get(alert_id)
        if not a:
            return False
        a["status"] = "resolved"
        a["resolved_at"] = datetime.now().isoformat(timespec="seconds")
        a["resolution_note"] = note
        return _save(store)
    except Exception as e:
        log.warning("StructuralAlerts: resolve_alert falló (%s)", e)
        return False


def user_notice() -> str:
    """
    Aviso SUAVE para el usuario final (no alarmista). Vacío si no hay nada
    pendiente. No revela detalles técnicos; solo sugiere prudencia.
    """
    if not has_pending_alerts():
        return ""
    return (
        "ℹ️ Nota: se detectaron posibles cambios normativos recientes que aún "
        "están en revisión por nuestro equipo. El contenido legal sigue "
        "disponible, pero te recomendamos verificar con un profesional los "
        "documentos sensibles hasta que se confirme la actualización."
    )


def developer_report() -> str:
    """
    Reporte DETALLADO para el desarrollador / dueño. Texto listo para mostrar en
    consola, log o panel de administración.
    """
    pend = get_pending_alerts()
    if not pend:
        return "✅ No hay alertas estructurales pendientes."
    lines = [f"🔧 {len(pend)} ALERTA(S) ESTRUCTURAL(ES) PENDIENTE(S) — requieren acción de un desarrollador:", ""]
    for i, a in enumerate(pend, 1):
        lines.append(f"{i}. [{a['id']}] {a['country']} · campo: {a['field']} · severidad: {a.get('severity','?')}")
        lines.append(f"   Cambio: {a['summary']}")
        if a.get("detail"):
            lines.append(f"   Qué actualizar: {a['detail']}")
        if a.get("source"):
            lines.append(f"   Fuente: {a['source']}")
        lines.append(f"   Detectada: {a['created']} · visto {a.get('seen_count',1)} vez/veces")
        lines.append(f"   Para resolver: structural_alerts.resolve_alert('{a['id']}', 'nota')")
        lines.append("")
    return "\n".join(lines)
