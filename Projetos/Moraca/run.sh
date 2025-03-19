#!/bin/bash

# Mudar para o diretório do script
cd "$(dirname "$0")"

# Verificar se o tkinter está instalado
if ! python3 -c "import tkinter" 2>/dev/null; then
    echo "Instalando tkinter..."
    sudo apt-get update
    sudo apt-get install -y python3-tk
fi

# Criar ambiente virtual se não existir
if [ ! -d "venv" ]; then
    echo "Criando ambiente virtual..."
    python3 -m venv venv
fi

# Ativar ambiente virtual
source venv/bin/activate

# Instalar dependências se necessário
if [ ! -f "venv/installed" ]; then
    echo "Instalando dependências..."
    pip install -r requirements.txt
    touch venv/installed
fi

# Executar o programa
python main.py 