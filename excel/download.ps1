# Script PowerShell para download dos arquivos Excel
# Controle de Imposto de Renda 2025

$ErrorActionPreference = "Stop"

$repo = "celloweb-ai/controle-ir-2025-excel-dio-challenge"
$version = "v1.0.0"
$baseUrl = "https://github.com/$repo/releases/download/$version"

Write-Host "📥 Baixando arquivos Excel do Controle IR 2025..." -ForegroundColor Cyan
Write-Host ""

# Criar diretório se não existir
if (!(Test-Path "excel")) {
    New-Item -ItemType Directory -Path "excel" | Out-Null
}
Set-Location excel

try {
    # Download Controle_IR_2025.xlsx
    Write-Host "📊 Baixando Controle_IR_2025.xlsx..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri "$baseUrl/Controle_IR_2025.xlsx" -OutFile "Controle_IR_2025.xlsx" -UseBasicParsing
    
    # Download bancos_apoio.xlsx
    Write-Host "🏦 Baixando bancos_apoio.xlsx..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri "$baseUrl/bancos_apoio.xlsx" -OutFile "bancos_apoio.xlsx" -UseBasicParsing
    
    # Download projeto_completo.xlsx
    Write-Host "📋 Baixando projeto_completo.xlsx..." -ForegroundColor Yellow
    Invoke-WebRequest -Uri "$baseUrl/projeto_completo.xlsx" -OutFile "projeto_completo.xlsx" -UseBasicParsing
    
    Write-Host ""
    Write-Host "✅ Download concluído com sucesso!" -ForegroundColor Green
    Write-Host ""
    Write-Host "📁 Arquivos disponíveis em: $(Get-Location)"
    Get-ChildItem -Filter *.xlsx | Format-Table Name, @{Label="Tamanho";Expression={"$([math]::Round($_.Length/1KB, 2)) KB"}}
    
    Write-Host ""
    Write-Host "🚀 Próximos passos:" -ForegroundColor Cyan
    Write-Host "   1. Abra o arquivo Controle_IR_2025.xlsx"
    Write-Host "   2. Habilite a edição se solicitado"
    Write-Host "   3. Comece a preencher seus dados"
    Write-Host ""
    Write-Host "📚 Documentação: https://github.com/$repo/tree/main/docs" -ForegroundColor Blue
    
} catch {
    Write-Host ""
    Write-Host "❌ Erro ao baixar arquivos: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "💡 Tente fazer o download manual em:" -ForegroundColor Yellow
    Write-Host "   https://github.com/$repo/releases/tag/$version"
    exit 1
}
