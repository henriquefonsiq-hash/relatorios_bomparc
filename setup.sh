#!/bin/bash
# Script para instalar as fontes locais no Streamlit Cloud

# Cria a pasta de fontes do sistema se não existir
mkdir -p ~/.local/share/fonts

# Copia as fontes da pasta .fonts para a pasta do sistema
if [ -d ".fonts" ]; then
    cp .fonts/*.ttf ~/.local/share/fonts/
fi

# Atualiza o cache de fontes do Linux
fc-cache -f -v
