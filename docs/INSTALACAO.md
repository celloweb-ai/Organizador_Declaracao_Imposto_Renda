# 💻 Guia de Instalação

## Requisitos do Sistema

### Software Necessário
- **Microsoft Excel 2016 ou superior** (Windows)
- **Microsoft Excel 2019 ou Microsoft 365** (macOS)
- **LibreOffice Calc 7.0+** (alternativa gratuita, com compatibilidade limitada)

### Configurações Recomendadas
- Sistema operacional atualizado
- Pelo menos 4GB de RAM
- 100MB de espaço livre em disco
- Macros habilitadas no Excel

---

## Instalação Passo a Passo

### 1. Download dos Arquivos

#### Opção A: Clone o repositório
```bash
git clone https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge.git
cd controle-ir-2025-excel-dio-challenge
```

#### Opção B: Download direto
1. Acesse [Releases](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/releases)
2. Baixe a última versão `Controle_IR_2025.xlsx`
3. Salve em uma pasta de fácil acesso

### 2. Habilitar Macros no Excel

#### Windows
1. Abra o Excel
2. Vá em **Arquivo > Opções > Central de Confiabilidade**
3. Clique em **Configurações da Central de Confiabilidade**
4. Selecione **Configurações de Macro**
5. Marque **Habilitar todas as macros** (atenção: use apenas para arquivos confiáveis)
6. Marque **Confiar no acesso ao modelo de objeto do projeto VBA**
7. Clique em **OK**

#### macOS
1. Abra o Excel
2. Vá em **Excel > Preferências > Segurança e Privacidade**
3. Em **Segurança de Macro**, selecione **Habilitar todas as macros**
4. Feche e reabra o Excel

### 3. Abrir a Planilha

1. Localize o arquivo `Controle_IR_2025.xlsx`
2. Clique duas vezes para abrir
3. Se aparecer o aviso de segurança, clique em **Habilitar Conteúdo**
4. A planilha estará pronta para uso

---

## Configuração Inicial

### Primeira Utilização

1. **Abra a aba Dashboard**
   - Verifique se todas as fórmulas estão funcionando
   - Confirme que a data está atualizada

2. **Configure seus dados**
   - Navegue até cada aba e preencha com suas informações
   - Comece pela aba "Rendimentos"

3. **Validação automática**
   - O sistema validará automaticamente os dados inseridos
   - Campos obrigatórios aparecerão destacados

---

## Scripts VBA (Opcional)

### Instalar Scripts de Alinhamento

1. **Abra o Editor VBA**
   - Pressione `Alt + F11` (Windows) ou `Opt + F11` (macOS)

2. **Insira um novo módulo**
   - Menu **Inserir > Módulo**

3. **Cole o script**
   - Abra o arquivo `src/scripts/MoverIconeParaPosicao.vba`
   - Copie todo o conteúdo
   - Cole no módulo criado

4. **Execute o script**
   - Pressione `F5` ou clique em **Executar**

---

## Resolução de Problemas

### Problema: Fórmulas não calculam
**Solução:**
- Verifique se o cálculo automático está habilitado
- Vá em **Fórmulas > Opções de Cálculo > Automático**

### Problema: Macros não funcionam
**Solução:**
- Confirme que as macros estão habilitadas
- Verifique se clicou em "Habilitar Conteúdo" ao abrir o arquivo

### Problema: Arquivo abre com erro
**Solução:**
- Certifique-se de usar Excel 2016 ou superior
- Tente reparar o arquivo: **Arquivo > Abrir > Procurar > Ferramentas > Abrir e Reparar**

### Problema: Dados não aparecem no Dashboard
**Solução:**
- Verifique se preencheu os dados nas abas corretas
- Pressione `Ctrl + Alt + F9` para recalcular todas as fórmulas

---

## Backup e Segurança

### Recomendações de Backup
1. **Salvamento automático**: Configure o Excel para salvar automaticamente a cada 10 minutos
2. **Cópias de segurança**: Mantenha cópias em cloud (OneDrive, Google Drive)
3. **Versões**: Salve versões mensais com data no nome do arquivo

### Segurança dos Dados
- **Senha**: Proteja o arquivo com senha (**Arquivo > Informações > Proteger Pasta de Trabalho**)
- **Criptografia**: Use criptografia de disco se o computador for compartilhado
- **Não compartilhe**: Dados fiscais são sensíveis - nunca envie por e-mail não criptografado

---

## Atualizações

### Como atualizar para versão mais recente

1. **Backup dos dados atuais**
   - Faça cópia do arquivo atual

2. **Baixe a nova versão**
   - Acesse [Releases](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/releases)
   - Baixe a versão mais recente

3. **Migre os dados**
   - Copie e cole seus dados do arquivo antigo para o novo
   - Verifique se todas as informações foram transferidas

---

## Suporte

Precisa de ajuda? 

- 🐛 [Reportar um bug](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues)
- 💡 [Sugerir melhorias](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues)
- 💬 [Discussões da comunidade](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/discussions)

---

## Próximos Passos

Após a instalação:

1. 📚 Leia a [Documentação da Estrutura](ESTRUTURA.md)
2. 🧮 Explore os [Exemplos Práticos](EXEMPLOS.md)
3. 📊 Entenda as [Fórmulas Utilizadas](FORMULAS.md)
4. ✅ Comece a inserir seus dados reais
