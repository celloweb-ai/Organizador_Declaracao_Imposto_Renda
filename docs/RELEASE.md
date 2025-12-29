# 📦 Guia de Release - Arquivos Excel

Este guia explica como criar uma release e disponibilizar os arquivos Excel do projeto.

---

## 🎯 Por que usar Releases?

Arquivos Excel (especialmente `.xlsx`) são binários e podem ser grandes:
- **Controle_IR_2025.xlsx**: ~290 KB
- **projeto_completo.xlsx**: ~77 KB
- **bancos_apoio.xlsx**: ~17 KB

O GitHub recomenda usar **Releases** para distribuir arquivos binários grandes ao invés de cometá-los diretamente no repositório.

---

## 🚀 Criando uma Release

### Passo 1: Preparar os Arquivos

1. **Certifique-se de ter os 3 arquivos Excel:**
   ```
   excel/
   ├── Controle_IR_2025.xlsx
   ├── bancos_apoio.xlsx
   └── projeto_completo.xlsx
   ```

2. **Verifique a integridade:**
   - Abra cada arquivo no Excel
   - Confirme que todas as abas estão presentes
   - Teste as fórmulas principais

### Passo 2: Criar a Release via Interface Web

1. **Acesse a página de Releases:**
   ```
   https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/releases
   ```

2. **Clique em "Draft a new release"**

3. **Preencha os campos:**

   **Tag version:**
   ```
   v1.0.0
   ```

   **Release title:**
   ```
   🎉 Controle IR 2025 - Lançamento Inicial v1.0.0
   ```

   **Description:**
   ```markdown
   ## 📊 Controle de Imposto de Renda 2025
   
   Primeira versão oficial do sistema de controle de declaração de Imposto de Renda 2025.
   
   ### ✨ Novidades
   
   - Dashboard interativo com resumo fiscal
   - Controle completo de rendimentos
   - Gestão de deduções e despesas profissionais
   - Registro de bens e direitos
   - Cálculo automático da base de cálculo
   - Validação automática de dados
   - 6 abas totalmente integradas
   
   ### 📁 Arquivos Incluídos
   
   - **Controle_IR_2025.xlsx** - Planilha principal (290 KB)
   - **bancos_apoio.xlsx** - Lista de bancos brasileiros (17 KB)
   - **projeto_completo.xlsx** - Exemplo completo (77 KB)
   
   ### 📚 Documentação
   
   - [Guia de Instalação](docs/INSTALACAO.md)
   - [Exemplos Práticos](docs/EXEMPLOS.md)
   - [Estrutura das Planilhas](docs/ESTRUTURA.md)
   - [Fórmulas Utilizadas](docs/FORMULAS.md)
   
   ### 🚀 Como Usar
   
   1. Baixe o arquivo `Controle_IR_2025.xlsx`
   2. Abra no Microsoft Excel 2016+
   3. Habilite a edição se solicitado
   4. Comece pela aba "Dashboard"
   5. Preencha seus dados nas abas específicas
   
   ### 🔧 Requisitos
   
   - Microsoft Excel 2016+ ou Microsoft 365
   - Macros habilitadas (se houver)
   - 4GB RAM recomendado
   
   ### 📝 Notas da Versão
   
   - Lançamento inicial
   - Testado no Excel 2019 e Microsoft 365
   - Compatível com Windows e macOS
   - Suporte limitado no LibreOffice
   
   ### 🐛 Reportar Problemas
   
   Encontrou um bug? [Abra uma issue](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues)
   
   ---
   
   **Desenvolvido como parte do Desafio DIO Bootcamp** 💜
   ```

4. **Anexar os arquivos Excel:**
   - Arraste os 3 arquivos `.xlsx` para a área de upload
   - Aguarde o upload completar
   - Confirme que todos os arquivos aparecem listados

5. **Configurar opções:**
   - [ ] Set as a pre-release (deixe desmarcado)
   - [ ] Set as latest release (marque esta opção)
   - [ ] Create a discussion (opcional)

6. **Publicar:**
   - Clique em **"Publish release"**

### Passo 3: Criar Release via GitHub CLI

Alternativamente, use o GitHub CLI:

```bash
# Instalar gh CLI se necessário
# https://cli.github.com/

# Fazer login
gh auth login

# Criar release com arquivos
gh release create v1.0.0 \
  excel/Controle_IR_2025.xlsx \
  excel/bancos_apoio.xlsx \
  excel/projeto_completo.xlsx \
  --title "🎉 Controle IR 2025 - Lançamento Inicial v1.0.0" \
  --notes-file release-notes.md
```

### Passo 4: Criar Release via GitHub API

Para automação:

```bash
# Criar a release
curl -X POST \
  -H "Accept: application/vnd.github+json" \
  -H "Authorization: Bearer YOUR_TOKEN" \
  https://api.github.com/repos/celloweb-ai/controle-ir-2025-excel-dio-challenge/releases \
  -d '{
    "tag_name": "v1.0.0",
    "name": "🎉 Controle IR 2025 - Lançamento Inicial v1.0.0",
    "body": "Primeira versão oficial...",
    "draft": false,
    "prerelease": false
  }'

# Upload dos arquivos
# (Requer upload_url obtido na resposta anterior)
curl -X POST \
  -H "Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" \
  -H "Authorization: Bearer YOUR_TOKEN" \
  --data-binary @excel/Controle_IR_2025.xlsx \
  "https://uploads.github.com/repos/celloweb-ai/controle-ir-2025-excel-dio-challenge/releases/RELEASE_ID/assets?name=Controle_IR_2025.xlsx"
```

---

## ✅ Verificar a Release

### Após publicar, confirme:

1. **Arquivos anexados:**
   - ✅ Controle_IR_2025.xlsx
   - ✅ bancos_apoio.xlsx
   - ✅ projeto_completo.xlsx

2. **Links de download funcionando:**
   ```bash
   # Teste o download
   wget https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/releases/download/v1.0.0/Controle_IR_2025.xlsx
   
   # Verifique o tamanho
   ls -lh Controle_IR_2025.xlsx
   ```

3. **Integridade dos arquivos:**
   - Baixe cada arquivo
   - Abra no Excel
   - Teste funcionalidades básicas

---

## 🔄 Atualizar uma Release

### Quando lançar uma nova versão:

1. **Crie uma nova tag:**
   ```bash
   git tag v1.1.0
   git push origin v1.1.0
   ```

2. **Crie nova release:**
   - Siga os mesmos passos acima
   - Use a nova tag (ex: `v1.1.0`)
   - Liste as mudanças no changelog

3. **Changelog sugerido:**
   ```markdown
   ## Mudanças em v1.1.0
   
   ### ✨ Novidades
   - Nova funcionalidade X
   - Melhoria na performance Y
   
   ### 🐛 Correções
   - Corrigido bug Z
   - Ajustado cálculo W
   
   ### 📚 Documentação
   - Atualizado guia de instalação
   - Novos exemplos adicionados
   ```

---

## 📊 Versionamento Semântico

Siga o padrão [SemVer](https://semver.org/):

```
v MAJOR . MINOR . PATCH
  │       │       └──── Correções de bugs
  │       └────────── Novas funcionalidades (compatível)
  └─────────────── Mudanças incompatíveis
```

**Exemplos:**
- `v1.0.0` - Primeira versão estável
- `v1.0.1` - Correção de bug pequeno
- `v1.1.0` - Nova funcionalidade
- `v2.0.0` - Mudança significativa (quebra compatibilidade)

---

## 📝 Template de Release Notes

```markdown
## 📊 Controle IR 2025 - v1.0.0

### 🎉 Destaques

Breve descrição das principais mudanças...

### ✨ Novidades

- Nova funcionalidade 1
- Nova funcionalidade 2
- Nova funcionalidade 3

### 🐛 Correções

- Corrigido bug A (#123)
- Corrigido bug B (#456)
- Ajustado comportamento C

### 📚 Documentação

- Atualizado README
- Novos exemplos em EXEMPLOS.md
- Melhorado guia de instalação

### 🔧 Melhorias Técnicas

- Otimização de fórmulas
- Performance melhorada
- Código VBA refatorado

### ⚠️ Breaking Changes

- Mudança X que quebra compatibilidade
- Nova estrutura de dados Y

### 📥 Downloads

| Arquivo | Tamanho | Descrição |
|---------|---------|-------------|
| Controle_IR_2025.xlsx | 290 KB | Planilha principal |
| bancos_apoio.xlsx | 17 KB | Lista de bancos |
| projeto_completo.xlsx | 77 KB | Exemplo completo |

### 🚀 Migração

Para atualizar da versão anterior:

1. Faça backup dos seus dados
2. Baixe a nova versão
3. Copie seus dados da versão antiga
4. Verifique compatibilidade

---

**Full Changelog**: v0.9.0...v1.0.0
```

---

## 🔗 Links Úteis

- [Documentação oficial GitHub Releases](https://docs.github.com/en/repositories/releasing-projects-on-github)
- [GitHub CLI - Releases](https://cli.github.com/manual/gh_release)
- [Versionamento Semântico](https://semver.org/lang/pt-BR/)
- [Keep a Changelog](https://keepachangelog.com/pt-BR/)

---

## ❓ FAQ

### Posso editar uma release depois de publicada?
**Sim!** Você pode:
- Editar o título e descrição
- Adicionar/remover arquivos
- Marcar/desmarcar como pre-release

### Quanto tempo os arquivos ficam disponíveis?
**Indefinidamente**, enquanto o repositório existir.

### Há limite de tamanho para arquivos?
**Sim**, o GitHub recomenda até 2GB por arquivo em releases.

### Posso deletar uma release?
**Sim**, mas é melhor marcar como "deprecated" para não quebrar links existentes.

---

**💡 Dica:** Sempre teste os downloads após publicar uma release!
