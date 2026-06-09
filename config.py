"""
WEP AI — Configuración central del proyecto.

Centraliza los model IDs de Anthropic en un solo lugar para evitar
desincronización entre módulos (problema histórico: brain.py quedó con
un modelo retirado mientras word_translator.py y llm_legal_generator.py
ya usaban el modelo actual).

Variables de entorno reconocidas:
    WEPAI_MODEL_CHAT        — default para chat principal y vision
    WEPAI_MODEL_TITLE       — default para generación de títulos cortos
    WEPAI_MODEL_TRANSLATOR  — override específico del traductor
    WEPAI_MODEL_LEGAL       — override específico del generador legal
    WEPAI_MODEL_VISION      — override específico de vision (si != chat)

Si no se setean, cae a los defaults definidos abajo.
"""

import os

# ── DEFAULTS — Modelos activos en mayo 2026 ────────────────────────────────────
# Sonnet 4.6 es el balanced workhorse de Anthropic; reemplazó al Sonnet 4 original
# (claude-sonnet-4-20250514) que fue retirado el 20-abr-2026.
# Si en algún momento se necesita máxima calidad: claude-opus-4-7 (2-3x más caro).
# Haiku 4.5 sigue vigente y es ~10x más barato — apropiado para tareas simples
# como generar títulos de 5 palabras.
DEFAULT_CHAT_MODEL       = os.environ.get("WEPAI_MODEL_CHAT",       "claude-sonnet-4-6")
DEFAULT_TITLE_MODEL      = os.environ.get("WEPAI_MODEL_TITLE",      "claude-haiku-4-5-20251001")
DEFAULT_VISION_MODEL     = os.environ.get("WEPAI_MODEL_VISION",     DEFAULT_CHAT_MODEL)
DEFAULT_TRANSLATOR_MODEL = os.environ.get("WEPAI_MODEL_TRANSLATOR", DEFAULT_CHAT_MODEL)
DEFAULT_LEGAL_MODEL      = os.environ.get("WEPAI_MODEL_LEGAL",      DEFAULT_CHAT_MODEL)

# ── RUTAS ──────────────────────────────────────────────────────────────────────
OUTPUT_DIR  = os.path.expanduser("~/Documents/WEP_AI")
USER_DATA   = os.path.expanduser("~/.wepai")
DB_PATH     = os.path.join(USER_DATA, "wepai.db")
CACHE_DIR   = USER_DATA  # word_translations_cache.json vive aquí
