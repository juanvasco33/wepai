#!/bin/bash
# WEP AI — Script de instalación para macOS
# Ejecutar con: bash install.sh

set -e

echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║           WEP AI — Instalación                  ║"
echo "║   Word · Excel · PowerPoint con IA para Mac     ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""

# ── 1. VERIFICAR PYTHON ───────────────────────────────────────────────────────
echo "▶ Verificando Python..."
if ! command -v python3 &>/dev/null; then
    echo "❌ Python 3 no encontrado."
    echo "   Instálalo desde: https://www.python.org/downloads/"
    exit 1
fi

PYTHON_VERSION=$(python3 -c 'import sys; print(f"{sys.version_info.major}.{sys.version_info.minor}")')
echo "   ✓ Python $PYTHON_VERSION encontrado"

# ── 2. VERIFICAR/CREAR ENTORNO VIRTUAL ───────────────────────────────────────
echo ""
echo "▶ Creando entorno virtual..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo "   ✓ Entorno virtual creado en ./venv"
else
    echo "   ✓ Entorno virtual ya existe"
fi

source venv/bin/activate

# ── 3. INSTALAR DEPENDENCIAS ──────────────────────────────────────────────────
echo ""
echo "▶ Instalando dependencias..."
pip install --upgrade pip -q
pip install -r requirements.txt -q
echo "   ✓ Dependencias instaladas"

# ── 4. CONFIGURAR API KEY ─────────────────────────────────────────────────────
echo ""
echo "▶ Configuración de API Key de Anthropic"
echo ""

if [ -z "$ANTHROPIC_API_KEY" ]; then
    echo "   No se detectó ANTHROPIC_API_KEY en el entorno."
    echo ""
    echo "   Para obtener tu clave:"
    echo "   1. Ve a https://console.anthropic.com"
    echo "   2. Crea una cuenta o inicia sesión"
    echo "   3. Agrega crédito (mínimo \$10 USD)"
    echo "   4. Ve a 'API Keys' y crea una nueva clave"
    echo ""
    read -p "   Pega tu ANTHROPIC_API_KEY aquí (o presiona Enter para saltarlo): " API_KEY

    if [ -n "$API_KEY" ]; then
        SHELL_RC="$HOME/.zshrc"
        [ ! -f "$SHELL_RC" ] && SHELL_RC="$HOME/.bash_profile"
        echo "" >> "$SHELL_RC"
        echo "# WEP AI — API Key de Anthropic" >> "$SHELL_RC"
        echo "export ANTHROPIC_API_KEY='$API_KEY'" >> "$SHELL_RC"
        export ANTHROPIC_API_KEY="$API_KEY"
        echo "   ✓ API Key guardada en $SHELL_RC"
    else
        echo "   ⚠️  API Key no configurada. Puedes agregarla después con:"
        echo "      export ANTHROPIC_API_KEY='tu-clave'"
    fi
else
    echo "   ✓ ANTHROPIC_API_KEY ya configurada"
fi

# ── 5. CREAR DIRECTORIOS ──────────────────────────────────────────────────────
echo ""
echo "▶ Creando directorios de trabajo..."
mkdir -p ~/Documents/WEP_AI
mkdir -p ~/.wepai
echo "   ✓ Directorio de documentos: ~/Documents/WEP_AI"
echo "   ✓ Directorio de configuración: ~/.wepai"

# ── 6. CREAR SCRIPT DE INICIO ─────────────────────────────────────────────────
echo ""
echo "▶ Creando script de inicio..."
cat > run_wepai.sh << 'EOF'
#!/bin/bash
DIR="$(cd "$(dirname "$0")" && pwd)"
source "$DIR/venv/bin/activate"
python "$DIR/main.py"
EOF
chmod +x run_wepai.sh
echo "   ✓ Script de inicio creado: run_wepai.sh"

# ── 7. VERIFICAR OFFICE ───────────────────────────────────────────────────────
echo ""
echo "▶ Verificando Microsoft Office..."
WORD_OK=false; EXCEL_OK=false; PPT_OK=false

[ -d "/Applications/Microsoft Word.app" ]        && WORD_OK=true  && echo "   ✓ Microsoft Word encontrado"
[ -d "/Applications/Microsoft Excel.app" ]       && EXCEL_OK=true && echo "   ✓ Microsoft Excel encontrado"
[ -d "/Applications/Microsoft PowerPoint.app" ]  && PPT_OK=true   && echo "   ✓ Microsoft PowerPoint encontrado"

if ! $WORD_OK && ! $EXCEL_OK && ! $PPT_OK; then
    echo "   ⚠️  Microsoft Office no encontrado."
    echo "   WEP AI generará los documentos pero no podrá abrirlos en Office."
fi

# ── 8. PERMISOS ────────────────────────────────────────────────────────────────
echo ""
echo "▶ Nota sobre permisos en macOS:"
echo "   La primera vez que WEP AI tome capturas de pantalla,"
echo "   macOS pedirá autorización. Por favor, concede el permiso."

# ── DONE ──────────────────────────────────────────────────────────────────────
echo ""
echo "╔══════════════════════════════════════════════════╗"
echo "║         ✓ Instalación de WEP AI completa        ║"
echo "╚══════════════════════════════════════════════════╝"
echo ""
echo "Para iniciar WEP AI, ejecuta:"
echo ""
echo "   bash run_wepai.sh"
echo ""
echo "O también:"
echo "   source venv/bin/activate && python main.py"
echo ""
