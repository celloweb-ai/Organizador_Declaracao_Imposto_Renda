#!/bin/bash
# Script para download dos arquivos Excel do projeto
# Controle de Imposto de Renda 2025

set -e

REPO="celloweb-ai/controle-ir-2025-excel-dio-challenge"
VERSION="v1.0.0"
BASE_URL="https://github.com/${REPO}/releases/download/${VERSION}"

echo "📥 Baixando arquivos Excel do Controle IR 2025..."
echo ""

# Criar diretório se não existir
mkdir -p excel
cd excel

# Download Controle_IR_2025.xlsx
echo "📊 Baixando Controle_IR_2025.xlsx..."
if command -v wget &> /dev/null; then
    wget -q --show-progress "${BASE_URL}/Controle_IR_2025.xlsx" -O Controle_IR_2025.xlsx
elif command -v curl &> /dev/null; then
    curl -L --progress-bar "${BASE_URL}/Controle_IR_2025.xlsx" -o Controle_IR_2025.xlsx
else
    echo "❌ Erro: wget ou curl não encontrado"
    exit 1
fi

# Download bancos_apoio.xlsx
echo "🏦 Baixando bancos_apoio.xlsx..."
if command -v wget &> /dev/null; then
    wget -q --show-progress "${BASE_URL}/bancos_apoio.xlsx" -O bancos_apoio.xlsx
elif command -v curl &> /dev/null; then
    curl -L --progress-bar "${BASE_URL}/bancos_apoio.xlsx" -o bancos_apoio.xlsx
fi

# Download projeto_completo.xlsx
echo "📋 Baixando projeto_completo.xlsx..."
if command -v wget &> /dev/null; then
    wget -q --show-progress "${BASE_URL}/projeto_completo.xlsx" -O projeto_completo.xlsx
elif command -v curl &> /dev/null; then
    curl -L --progress-bar "${BASE_URL}/projeto_completo.xlsx" -o projeto_completo.xlsx
fi

echo ""
echo "✅ Download concluído com sucesso!"
echo ""
echo "📁 Arquivos disponíveis em: $(pwd)"
ls -lh *.xlsx 2>/dev/null || echo "Verifique a pasta excel/"

echo ""
echo "🚀 Próximos passos:"
echo "   1. Abra o arquivo Controle_IR_2025.xlsx"
echo "   2. Habilite a edição se solicitado"
echo "   3. Comece a preencher seus dados"
echo ""
echo "📚 Documentação: https://github.com/${REPO}/tree/main/docs"
