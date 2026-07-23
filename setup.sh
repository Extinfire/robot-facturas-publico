#!/bin/bash
# Setup del SAT Invoice Bot
# Instala dependencias, configura Ollama y prepara el bot

set -e

echo "================================"
echo "🤖 SAT Invoice Bot Setup"
echo "================================"
echo ""

# Colores
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Paso 1: Verificar Python
echo "1️⃣  Verificando Python..."
if ! command -v python3 &> /dev/null; then
    echo -e "${RED}❌ Python3 no encontrado${NC}"
    exit 1
fi
PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo -e "${GREEN}✓ Python $PYTHON_VERSION${NC}"

# Paso 2: Crear venv
echo ""
echo "2️⃣  Creando entorno virtual..."
if [ ! -d "venv" ]; then
    python3 -m venv venv
    echo -e "${GREEN}✓ Entorno creado${NC}"
else
    echo -e "${GREEN}✓ Entorno ya existe${NC}"
fi

source venv/bin/activate

# Paso 3: Instalar dependencias Python
echo ""
echo "3️⃣  Instalando dependencias Python (desde requirements.txt)..."
pip install --upgrade pip > /dev/null 2>&1
if [ -f requirements.txt ]; then
    pip install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo -e "${RED}✗ Fallo la instalacion de requirements.txt${NC}"
        exit 1
    fi
else
    echo -e "${YELLOW}⚠ No se encontro requirements.txt, instalando minimas...${NC}"
    pip install requests python-telegram-bot --upgrade > /dev/null 2>&1
fi
echo -e "${GREEN}✓ Dependencias instaladas${NC}"

# Paso 4: Verificar Ollama
echo ""
echo "4️⃣  Verificando Ollama..."
if ! command -v ollama &> /dev/null; then
    echo -e "${YELLOW}⚠️  Ollama no está instalado${NC}"
    echo "Instálalo desde: https://ollama.ai"
    echo ""
    read -p "¿Continuar sin Ollama? (s/n): " -n 1 -r
    echo
    if [[ ! $REPLY =~ ^[Ss]$ ]]; then
        exit 1
    fi
else
    echo -e "${GREEN}✓ Ollama encontrado${NC}"
    
    # Verificar si está corriendo
    if curl -s http://localhost:11434/api/tags > /dev/null 2>&1; then
        echo -e "${GREEN}✓ Ollama está corriendo${NC}"
    else
        echo -e "${YELLOW}⚠️  Ollama no está corriendo${NC}"
        echo "Inicia con: ollama serve"
    fi
fi

# Paso 5: Crear estructura de carpetas
echo ""
echo "5️⃣  Creando estructura de carpetas..."
mkdir -p data
mkdir -p downloads
mkdir -p logs
echo -e "${GREEN}✓ Carpetas creadas${NC}"

# Paso 6: Descargar modelo Ollama
echo ""
echo "6️⃣  Modelos Ollama..."
if curl -s http://localhost:11434/api/tags 2>/dev/null | grep -q "llava-phi"; then
    echo -e "${GREEN}✓ llava-phi ya está descargado${NC}"
else
    echo -e "${YELLOW}⚠️  llava-phi no está descargado${NC}"
    echo ""
    echo "Descárgalo en otra terminal con:"
    echo -e "${YELLOW}ollama pull llava-phi${NC}"
    echo ""
fi

# Paso 7: Crear archivo de configuración
echo ""
echo "7️⃣  Configurando Telegram..."

if [ -z "$TELEGRAM_BOT_TOKEN" ]; then
    echo -e "${YELLOW}⚠️  TELEGRAM_BOT_TOKEN no está configurado${NC}"
    echo ""
    echo "1. Crea un bot con @BotFather en Telegram"
    echo "2. Luego ejecuta:"
    echo -e "${YELLOW}export TELEGRAM_BOT_TOKEN='tu_token_aqui'${NC}"
    echo ""
else
    echo -e "${GREEN}✓ TELEGRAM_BOT_TOKEN configurado${NC}"
fi

# Paso 8: Crear .env
echo ""
echo "8️⃣  Creando archivo .env..."
cat > .env << 'EOF'
# Telegram
TELEGRAM_BOT_TOKEN=YOUR_TOKEN_HERE

# Ollama
OLLAMA_BASE_URL=http://localhost:11434
OLLAMA_MODEL=llava-phi

# Directorio de datos
DATA_DIR=./data
DOWNLOAD_DIR=./downloads

# Logging
LOG_LEVEL=INFO
EOF
echo -e "${GREEN}✓ Archivo .env creado${NC}"
echo "   ⚠️  Actualiza TELEGRAM_BOT_TOKEN en .env"

# Paso 9: Resumen
echo ""
echo "================================"
echo -e "${GREEN}✅ Setup completado!${NC}"
echo "================================"
echo ""
echo "Próximos pasos:"
echo ""
echo "1️⃣  Asegúrate que Ollama está corriendo:"
echo -e "   ${YELLOW}ollama serve${NC}"
echo ""
echo "2️⃣  Descarga el modelo (si no lo está):"
echo -e "   ${YELLOW}ollama pull llava-phi${NC}"
echo ""
echo "3️⃣  Configura tu token de Telegram:"
echo -e "   ${YELLOW}export TELEGRAM_BOT_TOKEN='tu_token'${NC}"
echo "   o edita .env"
echo ""
echo "4️⃣  Inicia el bot:"
echo -e "   ${YELLOW}source venv/bin/activate${NC}"
echo -e "   ${YELLOW}python3 telegram_bot.py${NC}"
echo ""
echo "5️⃣  O prueba el módulo directamente:"
echo -e "   ${YELLOW}python3 sat_invoice_bot.py${NC}"
echo ""
echo "📝 Estructura del proyecto:"
echo "   ├── sat_invoice_bot.py (núcleo OCR + catálogo)"
echo "   ├── telegram_bot.py (interfaz Telegram)"
echo "   ├── data/ (catálogo JSON, caché)"
echo "   ├── downloads/ (imágenes descargadas)"
echo "   ├── logs/ (logs del bot)"
echo "   └── .env (configuración)"
echo ""
echo "💡 Tips:"
echo "   - Copia facturas a 'downloads/' para testing local"
echo "   - El catálogo crece automáticamente"
echo "   - Ver logs en 'logs/'"
echo ""
