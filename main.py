#!/usr/bin/env python3
"""
WEP AI — Asistente inteligente de Microsoft Office para macOS
Word · Excel · PowerPoint

Fase 1 MVP — Agente conversacional con control de Office

Para ejecutar:
    python main.py

Requisitos:
    - macOS con Microsoft Office instalado (Word, Excel, PowerPoint)
    - Python 3.11+
    - pip install -r requirements.txt
    - Variable de entorno: ANTHROPIC_API_KEY=tu-clave
"""

import os, sys

# ── VALIDATE ENVIRONMENT ──────────────────────────────────────────────────────
# v15.16.0 — Diagnóstico claro del estado de la verificación legal autónoma.
# En vez de revisar solo la API key, comprobamos SDK + key y avisamos al usuario
# si la auto-actualización por IA va a quedar inactiva (en vez de fallar callado).
try:
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from agent.legal_overrides import verification_status
    _vs = verification_status()
    print(_vs["mensaje"])
    if not _vs["operativa"]:
        if not _vs["api_key"]:
            print("   Ejecuta: export ANTHROPIC_API_KEY='tu-clave-de-anthropic'")
            print("   Obtén tu clave en: https://console.anthropic.com")
        if not _vs["sdk"]:
            print("   Instala el SDK: pip install anthropic")
    print()
    # v16.8 — Aviso al desarrollador si hay cambios legales ESTRUCTURALES
    # pendientes de implementar (fuera del alcance de la auto-actualización).
    try:
        from agent.structural_alerts import has_pending_alerts, developer_report
        if has_pending_alerts():
            print("=" * 70)
            print(developer_report())
            print("=" * 70)
            print()
    except Exception:
        pass
except Exception:
    # Fallback al chequeo simple si el módulo no carga
    if not os.environ.get("ANTHROPIC_API_KEY"):
        print("⚠️  ANTHROPIC_API_KEY no configurada.")
        print("   Ejecuta: export ANTHROPIC_API_KEY='tu-clave-de-anthropic'")
        print()

# ── PYTHON VERSION CHECK ──────────────────────────────────────────────────────
if sys.version_info < (3, 11):
    print("❌ Se requiere Python 3.11 o superior.")
    print(f"   Versión actual: {sys.version}")
    sys.exit(1)

# ── CREATE OUTPUT DIRECTORY ───────────────────────────────────────────────────
os.makedirs(os.path.expanduser("~/Documents/WEP_AI"), exist_ok=True)
os.makedirs(os.path.expanduser("~/.wepai"), exist_ok=True)

# ── LAUNCH APP ────────────────────────────────────────────────────────────────
try:
    from ui.chat_window import run
    print("🚀 Iniciando WEP AI...")
    run()
except ImportError as e:
    print(f"❌ Error importando módulos: {e}")
    print("   Asegúrate de instalar las dependencias:")
    print("   pip install -r requirements.txt")
    sys.exit(1)
except Exception as e:
    print(f"❌ Error al iniciar la app: {e}")
    raise
