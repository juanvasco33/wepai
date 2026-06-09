#!/usr/bin/env bash
# =============================================================================
# WEP AI — Empaquetado para macOS (.app + .dmg)
# =============================================================================
# IMPORTANTE: este script DEBE ejecutarse EN UNA MAC. PyInstaller genera un
# binario para el sistema operativo en el que corre; no se puede construir un
# .app de Mac desde Linux/Windows.
#
# Uso:
#   chmod +x build_macos.sh
#   ./build_macos.sh
#
# Resultado:
#   dist/WEP AI.app   → la aplicación
#   dist/WEP-AI.dmg   → el instalador para distribuir en la web
# =============================================================================
set -euo pipefail

APP_NAME="WEP AI"
DMG_NAME="WEP-AI"
ENTRY="main.py"
ICON="web/static/logo.png"   # idealmente convertir a .icns (ver nota abajo)

cd "$(dirname "$0")"

echo "==> 1/5  Verificando herramientas…"
command -v python3 >/dev/null || { echo "Falta python3"; exit 1; }

echo "==> 2/5  Creando entorno virtual de build…"
python3 -m venv .build_venv
# shellcheck disable=SC1091
source .build_venv/bin/activate
pip install --upgrade pip wheel >/dev/null
pip install -r requirements.txt >/dev/null
pip install pyinstaller pillow >/dev/null

# --- Convertir el PNG del logo a .icns (formato de icono de Mac) -------------
echo "==> 3/5  Generando icono .icns…"
ICNS_PATH="build/wepai.icns"
mkdir -p build/wepai.iconset
if [ -f "$ICON" ]; then
  # genera los tamaños requeridos por macOS a partir del PNG
  for size in 16 32 64 128 256 512; do
    sips -z $size $size      "$ICON" --out "build/wepai.iconset/icon_${size}x${size}.png"      >/dev/null 2>&1 || true
    dbl=$((size*2))
    sips -z $dbl  $dbl       "$ICON" --out "build/wepai.iconset/icon_${size}x${size}@2x.png"   >/dev/null 2>&1 || true
  done
  iconutil -c icns build/wepai.iconset -o "$ICNS_PATH" 2>/dev/null || ICNS_PATH=""
else
  echo "   (No se encontró $ICON; la app saldrá con icono genérico.)"
  ICNS_PATH=""
fi

# --- Construir el .app con PyInstaller --------------------------------------
echo "==> 4/5  Construyendo el .app con PyInstaller…"
ICON_FLAG=""
[ -n "$ICNS_PATH" ] && ICON_FLAG="--icon $ICNS_PATH"

# Notas de flags:
#  --windowed         → app GUI sin terminal
#  --collect-all customtkinter  → CTk usa archivos de datos que hay que incluir
#  --noconfirm        → sobreescribe dist/ sin preguntar
pyinstaller \
  --name "$APP_NAME" \
  --windowed \
  --noconfirm \
  --clean \
  $ICON_FLAG \
  --collect-all customtkinter \
  --collect-all docx \
  --collect-all pptx \
  --collect-all openpyxl \
  --hidden-import PIL._tkinter_finder \
  "$ENTRY"

# --- Empaquetar en .dmg ------------------------------------------------------
echo "==> 5/5  Generando el .dmg…"
if command -v create-dmg >/dev/null; then
  rm -f "dist/${DMG_NAME}.dmg"
  create-dmg \
    --volname "$APP_NAME" \
    --window-size 540 380 \
    --icon-size 110 \
    --icon "$APP_NAME.app" 140 180 \
    --app-drop-link 400 180 \
    "dist/${DMG_NAME}.dmg" \
    "dist/$APP_NAME.app"
else
  echo "   'create-dmg' no está instalado. Instálalo con:  brew install create-dmg"
  echo "   Alternativa rápida sin create-dmg:"
  echo "   hdiutil create -volname \"$APP_NAME\" -srcfolder \"dist/$APP_NAME.app\" -ov -format UDZO \"dist/${DMG_NAME}.dmg\""
  hdiutil create -volname "$APP_NAME" -srcfolder "dist/$APP_NAME.app" -ov -format UDZO "dist/${DMG_NAME}.dmg"
fi

deactivate
echo ""
echo "✅ Listo:"
echo "   dist/$APP_NAME.app   (la aplicación)"
echo "   dist/${DMG_NAME}.dmg  (súbelo a la web para descargar)"
echo ""
echo "─────────────────────────────────────────────────────────────────────"
echo "FIRMA Y NOTARIZACIÓN (opcional pero recomendado para distribuir):"
echo ""
echo "Sin firmar, el usuario verá 'desarrollador no identificado' y deberá"
echo "ir a Ajustes → Privacidad y Seguridad → 'Abrir de todos modos'."
echo ""
echo "Para que abra sin advertencias necesitas el Apple Developer Program"
echo "(US\$99/año) y luego:"
echo ""
echo "  codesign --deep --force --options runtime \\"
echo "    --sign \"Developer ID Application: TU NOMBRE (TEAMID)\" \\"
echo "    \"dist/$APP_NAME.app\""
echo ""
echo "  xcrun notarytool submit \"dist/${DMG_NAME}.dmg\" \\"
echo "    --apple-id TU_APPLE_ID --team-id TEAMID --password APP_SPECIFIC_PWD \\"
echo "    --wait"
echo ""
echo "  xcrun stapler staple \"dist/${DMG_NAME}.dmg\""
echo "─────────────────────────────────────────────────────────────────────"
