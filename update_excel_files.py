#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para atualizar os arquivos Excel no repositório
Com o conteúdo binário completo dos arquivos originais

Uso:
    python update_excel_files.py
    
Requisitos:
    - Arquivos projeto_completo.xlsx e bancos_apoio.xlsx no diretório atual
    - GitHub CLI (gh) instalado e configurado
    - Permissões de escrita no repositório
"""

import base64
import os
import sys
import subprocess

def check_file_exists(filename):
    """Verifica se o arquivo existe no diretório atual"""
    if not os.path.exists(filename):
        print(f"❌ Erro: Arquivo '{filename}' não encontrado!")
        print(f"   Por favor, coloque o arquivo no diretório atual.")
        return False
    return True

def read_and_encode(filename):
    """Lê arquivo binário e codifica em base64"""
    try:
        with open(filename, 'rb') as f:
            content = f.read()
            base64_content = base64.b64encode(content).decode('utf-8')
        print(f"✅ {filename}: {len(content)} bytes -> {len(base64_content)} chars base64")
        return base64_content
    except Exception as e:
        print(f"❌ Erro ao ler {filename}: {e}")
        return None

def save_base64_file(content, output_filename):
    """Salva conteúdo base64 em arquivo temporário"""
    try:
        with open(output_filename, 'w') as f:
            f.write(content)
        print(f"✅ Salvo: {output_filename}")
        return True
    except Exception as e:
        print(f"❌ Erro ao salvar {output_filename}: {e}")
        return False

def upload_to_github(filename, message):
    """Faz upload do arquivo para o GitHub usando gh CLI"""
    try:
        # Adiciona arquivo
        subprocess.run(['git', 'add', filename], check=True)
        
        # Commit
        subprocess.run(['git', 'commit', '-m', message], check=True)
        
        # Push
        subprocess.run(['git', 'push'], check=True)
        
        print(f"✅ {filename} enviado com sucesso para o GitHub!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao enviar {filename}: {e}")
        return False

def main():
    print("="*70)
    print("   Atualização de Arquivos Excel - Repositório GitHub")
    print("   Organizador de Declaração de Imposto de Renda")
    print("="*70)
    print()
    
    files = [
        ('projeto_completo.xlsx', 'fix: Atualiza projeto_completo.xlsx com conteúdo binário completo'),
        ('bancos_apoio.xlsx', 'fix: Atualiza bancos_apoio.xlsx com conteúdo binário completo')
    ]
    
    success_count = 0
    
    for filename, commit_message in files:
        print(f"\n📄 Processando: {filename}")
        print("-" * 70)
        
        # Verifica existência
        if not check_file_exists(filename):
            continue
        
        # Lê e codifica
        base64_content = read_and_encode(filename)
        if not base64_content:
            continue
        
        # Salva arquivo base64 temporário (opcional, para debug)
        temp_filename = f"{filename}.base64.txt"
        save_base64_file(base64_content, temp_filename)
        
        print(f"\n🚀 Enviando {filename} para o GitHub...")
        print(f"   Mensagem: {commit_message}")
        
        if upload_to_github(filename, commit_message):
            success_count += 1
        
        # Remove arquivo temporário
        if os.path.exists(temp_filename):
            os.remove(temp_filename)
    
    print("\n" + "="*70)
    print(f"✅ Resumo: {success_count}/{len(files)} arquivos enviados com sucesso!")
    print("="*70)
    
    if success_count == len(files):
        print("\n✨ Todos os arquivos foram atualizados no repositório GitHub!")
        return 0
    else:
        print("\n⚠️  Alguns arquivos não foram atualizados. Verifique os erros acima.")
        return 1

if __name__ == "__main__":
    sys.exit(main())
