#!/usr/bin/env python3
"""
WEP AI — E2E test suite para macOS con Microsoft Office (v14.6)

PARA QUE ESTE SCRIPT FUNCIONE:
─────────────────────────────
1. Estar en macOS (cualquier versión 10.15+)
2. Tener Microsoft Office instalado (Word, Excel y PowerPoint)
3. Tener Python 3.11+ con las deps de WEP AI:
       pip install -r requirements.txt
4. Tener configurada la variable ANTHROPIC_API_KEY (no es estrictamente
   necesaria si testeás solo los controllers — los pasos del LLM se skippean)
5. Correr desde la raíz del paquete wepai/:
       python3 tests/test_mac_e2e.py

QUÉ HACE:
─────────
Valida los pasos del round-trip completo que NO se pueden testear en CI:

  1. Genera un .docx, .xlsx y .pptx con los controllers
  2. Los abre con Word/Excel/PowerPoint vía AppleScript
  3. Toma una screenshot del documento abierto
  4. (Opcional) Llama a Claude Vision para verificar calidad
  5. Cierra las apps

CADA STEP REPORTA:
  ✓ pasado     → la funcionalidad funciona
  ⚠ skip       → no aplica en este entorno (ej. sin API key, sin Office)
  ✗ falló      → BUG real, requiere fix

REGISTRO:
─────────
Los logs detallados se escriben en /tmp/wepai_e2e_$(date).log para que
puedas mandarlos a soporte si algo falla.
"""

import os
import sys
import time
import subprocess
import platform
import tempfile
import logging
from datetime import datetime

# ── Setup ──────────────────────────────────────────────────────────────────
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

log_path = f"/tmp/wepai_e2e_{datetime.now():%Y%m%d_%H%M%S}.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    handlers=[
        logging.FileHandler(log_path),
        logging.StreamHandler(),
    ],
)
log = logging.getLogger("wepai.e2e")

# ── Estado del test ────────────────────────────────────────────────────────
RESULTS = []  # (name, status, detail)
def step(name, status, detail=""):
    RESULTS.append((name, status, detail))
    icon = {"PASS": "✓", "FAIL": "✗", "SKIP": "⚠"}.get(status, "?")
    msg = f"{icon} {name}"
    if detail:
        msg += f"  ({detail})"
    log.info(msg)


# ── Step 0: ambiente ───────────────────────────────────────────────────────
def step_environment():
    if platform.system() != "Darwin":
        step("S0. macOS detected", "FAIL", f"corriendo en {platform.system()}")
        return False
    step("S0. macOS detected", "PASS", platform.mac_ver()[0])

    # ¿Office instalado?
    for app, path in [
        ("Microsoft Word", "/Applications/Microsoft Word.app"),
        ("Microsoft Excel", "/Applications/Microsoft Excel.app"),
        ("Microsoft PowerPoint", "/Applications/Microsoft PowerPoint.app"),
    ]:
        if os.path.exists(path):
            step(f"S0. {app} installed", "PASS")
        else:
            step(f"S0. {app} installed", "FAIL", f"no encontrado en {path}")

    # ¿Python deps?
    for pkg in ["docx", "openpyxl", "pptx", "anthropic", "bcrypt"]:
        try:
            __import__(pkg)
            step(f"S0. python pkg '{pkg}'", "PASS")
        except ImportError as e:
            step(f"S0. python pkg '{pkg}'", "FAIL", str(e))

    # ¿API key?
    if os.getenv("ANTHROPIC_API_KEY"):
        step("S0. ANTHROPIC_API_KEY set", "PASS")
    else:
        step("S0. ANTHROPIC_API_KEY set", "SKIP",
             "tests de LLM/vision se skippean")
    return True


# ── Step 1: Generación local de docs ───────────────────────────────────────
def step_generate_docs():
    """Generar 3 docs (uno por tipo) y verificar que existen."""
    try:
        from office.word_controller import generate_word
        from office.excel_controller import generate_excel
        from office.ppt_controller import generate_powerpoint
    except ImportError as e:
        step("S1. Import controllers", "FAIL", str(e))
        return None
    step("S1. Import controllers", "PASS")

    files = {}
    # Word
    try:
        files["word"] = generate_word({
            "template_id": "nda",
            "titulo": "E2E_NDA_Test",
            "instrucciones": "Test NDA",
            "datos": {"idioma": "en"},
        })
        if os.path.exists(files["word"]):
            step("S1. Word generated", "PASS", files["word"])
        else:
            step("S1. Word generated", "FAIL", "archivo no existe")
    except Exception as e:
        step("S1. Word generated", "FAIL", str(e))
        files["word"] = None
    # Excel
    try:
        fp, _ = generate_excel({
            "template_id": "inventory",
            "titulo": "E2E_Inv_Test",
            "instrucciones": "Test inventory",
            "datos": {"idioma": "en"},
        })
        files["excel"] = fp
        if os.path.exists(fp):
            step("S1. Excel generated", "PASS", fp)
        else:
            step("S1. Excel generated", "FAIL", "archivo no existe")
    except Exception as e:
        step("S1. Excel generated", "FAIL", str(e))
        files["excel"] = None
    # PowerPoint
    try:
        files["pptx"] = generate_powerpoint({
            "template_id": "pitch",
            "titulo": "E2E_Pitch_Test",
            "instrucciones": "Test pitch",
            "datos": {"idioma": "en"},
        })
        if os.path.exists(files["pptx"]):
            step("S1. PowerPoint generated", "PASS", files["pptx"])
        else:
            step("S1. PowerPoint generated", "FAIL", "archivo no existe")
    except Exception as e:
        step("S1. PowerPoint generated", "FAIL", str(e))
        files["pptx"] = None
    return files


# ── Step 2: Abrir cada uno con AppleScript ─────────────────────────────────
def step_open_in_office(files):
    """Abrir Word, Excel, PowerPoint con AppleScript y verificar."""
    if not files:
        step("S2. Open in Office", "SKIP", "sin archivos del paso anterior")
        return

    def _run_osa(script, label, timeout=15):
        try:
            r = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=timeout)
            if r.returncode == 0:
                step(f"S2. {label}", "PASS")
                return True
            step(f"S2. {label}", "FAIL", r.stderr.strip()[:120])
            return False
        except subprocess.TimeoutExpired:
            step(f"S2. {label}", "FAIL", f"timeout {timeout}s")
            return False
        except Exception as e:
            step(f"S2. {label}", "FAIL", str(e))
            return False

    # Word
    if files.get("word"):
        script = f'''
            tell application "Microsoft Word"
                activate
                open POSIX file "{files['word']}"
                delay 2
            end tell
        '''
        _run_osa(script, "Open Word file")
    # Excel
    if files.get("excel"):
        script = f'''
            tell application "Microsoft Excel"
                activate
                open POSIX file "{files['excel']}"
                delay 2
            end tell
        '''
        _run_osa(script, "Open Excel file")
    # PowerPoint
    if files.get("pptx"):
        script = f'''
            tell application "Microsoft PowerPoint"
                activate
                open POSIX file "{files['pptx']}"
                delay 2
            end tell
        '''
        _run_osa(script, "Open PowerPoint file")


# ── Step 3: Screenshot ─────────────────────────────────────────────────────
def step_screenshot():
    """Tomar screenshot de la pantalla y verificar tamaño razonable."""
    try:
        from agent.vision import take_screenshot
    except ImportError as e:
        step("S3. Import vision", "FAIL", str(e))
        return None
    step("S3. Import vision", "PASS")
    try:
        ss_b64 = take_screenshot()
        if ss_b64 and len(ss_b64) > 10000:
            step("S3. Screenshot captured", "PASS",
                 f"{len(ss_b64)//1024} KB base64")
            return ss_b64
        step("S3. Screenshot captured", "FAIL",
             f"resultado vacío o demasiado pequeño ({len(ss_b64) if ss_b64 else 0} bytes)")
    except Exception as e:
        step("S3. Screenshot captured", "FAIL", str(e))
    return None


# ── Step 4: Vision verification (opcional) ─────────────────────────────────
def step_vision_verify(ss_b64):
    """Llamar a Claude Vision para verificar el doc abierto."""
    if not ss_b64:
        step("S4. Vision verify", "SKIP", "sin screenshot")
        return
    if not os.getenv("ANTHROPIC_API_KEY"):
        step("S4. Vision verify", "SKIP", "sin ANTHROPIC_API_KEY")
        return
    try:
        from agent.brain import analyze_screenshot
        feedback = analyze_screenshot(ss_b64, "word", "NDA test")
        if feedback and len(feedback) > 10:
            step("S4. Vision verify", "PASS",
                 f"{len(feedback)} chars de feedback")
            log.info(f"  Vision feedback: {feedback[:200]}...")
        else:
            step("S4. Vision verify", "FAIL",
                 f"feedback vacío o muy corto: {feedback!r}")
    except Exception as e:
        step("S4. Vision verify", "FAIL", str(e))


# ── Step 5: Cerrar las apps ────────────────────────────────────────────────
def step_close_apps():
    """Cerrar Word/Excel/PowerPoint sin guardar."""
    for app in ["Microsoft Word", "Microsoft Excel", "Microsoft PowerPoint"]:
        script = f'tell application "{app}" to quit saving no'
        try:
            r = subprocess.run(
                ["osascript", "-e", script],
                capture_output=True, text=True, timeout=10)
            if r.returncode == 0:
                step(f"S5. Quit {app}", "PASS")
            else:
                step(f"S5. Quit {app}", "FAIL", r.stderr.strip()[:120])
        except Exception as e:
            step(f"S5. Quit {app}", "FAIL", str(e))


# ── MAIN ───────────────────────────────────────────────────────────────────
def main():
    print("═" * 65)
    print("WEP AI — E2E test suite para macOS + Microsoft Office")
    print(f"Log: {log_path}")
    print("═" * 65)
    print()

    if not step_environment():
        print("\nSe abortó: no es macOS. Este script solo corre en Mac.")
        sys.exit(2)

    files = step_generate_docs()

    print("\n--- Steps que requieren Office abierto ---")
    confirm = input("¿Continuar abriendo Word/Excel/PowerPoint? (s/n): ")
    if confirm.lower() not in ("s", "y", "yes", "si", "sí"):
        print("Saltando los steps de Office. Tests parciales completados.")
    else:
        step_open_in_office(files)
        time.sleep(3)  # dar tiempo a que las apps muestren contenido
        ss = step_screenshot()
        step_vision_verify(ss)
        confirm_close = input("\n¿Cerrar Word/Excel/PowerPoint ahora? (s/n): ")
        if confirm_close.lower() in ("s", "y", "yes", "si", "sí"):
            step_close_apps()

    # ── Resumen ────────────────────────────────────────────────────────────
    print("\n" + "═" * 65)
    print("RESUMEN")
    print("═" * 65)
    passed = sum(1 for _, s, _ in RESULTS if s == "PASS")
    failed = sum(1 for _, s, _ in RESULTS if s == "FAIL")
    skipped = sum(1 for _, s, _ in RESULTS if s == "SKIP")
    total = len(RESULTS)
    print(f"Total: {total}  |  ✓ PASS: {passed}  |  ⚠ SKIP: {skipped}  |  ✗ FAIL: {failed}")
    print(f"Log completo: {log_path}")
    if failed > 0:
        print("\nFAILS encontrados:")
        for name, status, detail in RESULTS:
            if status == "FAIL":
                print(f"  • {name} — {detail}")
        sys.exit(1)
    print("\n✓ Todos los tests pasaron o se skippearon razonablemente.")


if __name__ == "__main__":
    main()
