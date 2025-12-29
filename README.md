# 🧾 Organizador de Declaração de Imposto de Renda - Excel

<div align="center">

![Excel](https://img.shields.io/badge/Microsoft_Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![DIO](https://img.shields.io/badge/DIO-Bootcamp-purple?style=for-the-badge)
![License](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-Completo-success?style=for-the-badge)

**Ferramenta Excel para Organização e Controle de Declaração de Imposto de Renda 2025**

Desenvolvido como parte do desafio DIO Bootcamp

[Sobre o Projeto](#-sobre-o-projeto) •
[Funcionalidades](#-funcionalidades) •
[Estrutura](#-estrutura-do-projeto) •
[Como Usar](#-como-usar) •
[Documentação](#-documentação) •
[Autor](#-autor)

</div>

---

## 📊 Sobre o Projeto

Este projeto consiste em uma **ferramenta completa desenvolvida em Microsoft Excel** para auxiliar na organização e reunião de informações essenciais para a declaração de imposto de renda. A solução foi construída como parte do desafio DIO Bootcamp, aplicando conceitos de organização de dados, validações automáticas, interface amigável e documentação técnica.

### 🎯 Objetivos do Projeto

- ✅ Criar agregador de dados fiscais validado e eficiente
- ✅ Implementar menus de navegação intuitivos entre abas
- ✅ Desenvolver validações automáticas de entrada de dados
- ✅ Gerar resumo fiscal consolidado automaticamente
- ✅ Fornecer interface amigável e profissional
- ✅ Documentar todo o processo de forma clara e estruturada

### 💡 Conceitos Aplicados

- **Organização de Dados**: Estruturação lógica de informações fiscais
- **Validação de Dados**: Uso de validadores do Excel para garantir integridade
- **Fórmulas e Funções**: Cálculos automáticos e consolidação de dados
- **Interface de Usuário**: Design profissional e navegabilidade
- **Documentação Técnica**: README estruturado e arquivos de suporte
- **Controle de Versão**: Utilização do GitHub para compartilhamento

---

## ✨ Funcionalidades

### 📊 Dashboard Executivo
- **Resumo Fiscal 2025** com visualização instantânea
- Indicadores principais:
  - Rendimentos Totais
  - Deduções Totais
  - Despesas Profissionais
  - Base de Cálculo
  - Imposto Estimado (15%)
  - Patrimônio Total
- Data de última atualização
- Menu de navegação rápida entre abas

### 💵 Controle de Rendimentos
- Registro detalhado de lançamentos por data
- Categorias pré-definidas:
  - Salário
  - Aluguel
  - Honorários
  - Outros rendimentos
- Soma automática de valores
- Formatação monetária brasileira (R$)

### 📋 Gestão de Deduções
- Categorias de deduções:
  - Previdenciária (INSS)
  - Educação
  - Saúde
  - Dependentes
- Rastreamento por data e descrição
- Totalização automática

### 💼 Despesas Profissionais
- Controle de despesas dedutíveis:
  - Material de escritório
  - Software/Licenças
  - Aluguel comercial
  - Energia/Água
  - Transporte profissional
- Organização cronológica
- Cálculo automático de totais

### 🏠 Bens e Direitos
- Registro de patrimônio:
  - Imóveis (residenciais e comerciais)
  - Veículos
  - Investimentos (ações, títulos, fundos)
  - Outros bens
- Campos detalhados:
  - Descrição completa
  - Localização
  - Data de aquisição
  - Valor atualizado
  - Observações
- Soma automática do patrimônio total

### 📈 Resumo Fiscal Automático
- Consolidação de todos os lançamentos
- Cálculos automáticos:
  - Base de cálculo (Rendimentos - Deduções - Despesas)
  - Imposto estimado (alíquota de 15%)
  - Patrimônio total declarado
- Visão consolidada para declaração

---

## 📁 Estrutura do Projeto

```
controle-ir-2025-excel-dio-challenge/
│
├── README.md                      # Documentação principal do projeto
├── LICENSE                        # Licença MIT
├── .gitattributes                 # Configuração Git LFS para arquivos Excel
│
├── excel/
│   ├── Controle_IR_2025.xlsx      # Planilha principal do projeto
│   ├── bancos_apoio.xlsx          # Lista de bancos brasileiros (validador)
│   └── projeto_completo.xlsx      # Exemplo completo de referência
│
├── scripts/
│   └── script_de_alinhamentos.txt # VBA para alinhamento de ícones
│
├── docs/
│   ├── ESTRUTURA.md               # Documentação da estrutura das planilhas
│   ├── FORMULAS.md                # Documentação de fórmulas utilizadas
│   └── GUIA_USO.md                # Guia passo a passo de utilização
│
└── images/
    ├── dashboard_screenshot.png   # Captura de tela do Dashboard
    ├── rendimentos_screenshot.png # Captura de tela de Rendimentos
    └── resumo_screenshot.png      # Captura de tela do Resumo Fiscal
```

### 🗂️ Descrição dos Arquivos

#### Planilhas Excel

**`Controle_IR_2025.xlsx`** - Planilha principal do projeto
- **Dashboard**: Painel executivo com resumo e navegação
- **Rendimentos**: Controle de entradas financeiras
- **Deduções**: Registro de gastos dedutíveis
- **Despesas Profissionais**: Controle de despesas de trabalho
- **Bens e Direitos**: Inventário patrimonial
- **Resumo Fiscal**: Consolidação automática de dados

**`bancos_apoio.xlsx`** - Base de dados auxiliar
- Lista completa de 50+ instituições financeiras brasileiras
- Formato: Código - Nome do Banco
- Utilizado para validação de dados de entrada

**`projeto_completo.xlsx`** - Arquivo de referência
- Exemplo completo preenchido
- Demonstra todas as funcionalidades
- Inclui dados fictícios para aprendizado
- Sistema de navegação "LION APP"

#### Scripts

**`script_de_alinhamentos.txt`** - Macro VBA
- Código para posicionamento automático de ícones
- Função: `MoverIconeParaPosicao()`
- Parametrizável para ajustes de layout

---

## 🚀 Como Usar

### Pré-requisitos

- Microsoft Excel 2016 ou superior
- Sistema operacional: Windows, macOS ou Linux (com LibreOffice Calc)
- Conhecimentos básicos de Excel

### Instalação

1. **Clone este repositório:**
```bash
git clone https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge.git
cd controle-ir-2025-excel-dio-challenge
```

2. **Ou faça o download direto:**
- Clique no botão verde "Code" > "Download ZIP"
- Extraia os arquivos em uma pasta local

### Uso da Planilha

1. **Abra o arquivo principal:**
   - Navegue até a pasta `excel/`
   - Abra `Controle_IR_2025.xlsx`

2. **Navegue pelas abas:**
   - Use os botões de navegação no Dashboard
   - Ou clique diretamente nas abas na parte inferior

3. **Preencha seus dados:**
   - **Rendimentos**: Adicione linhas com data, categoria, descrição e valor
   - **Deduções**: Registre gastos dedutíveis por categoria
   - **Despesas Profissionais**: Lance despesas relacionadas ao trabalho
   - **Bens e Direitos**: Cadastre seu patrimônio detalhadamente

4. **Consulte o Resumo:**
   - Acesse a aba "Resumo Fiscal" ou "Dashboard"
   - Todos os cálculos são atualizados automaticamente

### Dicas de Uso

✅ **Faça backup regular** da planilha antes de grandes alterações

✅ **Use as categorias pré-definidas** para padronizar seus lançamentos

✅ **Verifique o resumo fiscal** periodicamente para acompanhar sua situação

✅ **Mantenha documentos de suporte** (recibos, notas) organizados

✅ **Consulte o arquivo `projeto_completo.xlsx`** para ver exemplos de preenchimento

---

## 📚 Documentação Adicional

### Arquivos de Documentação

- **[ESTRUTURA.md](docs/ESTRUTURA.md)**: Documentação detalhada da estrutura de cada aba
- **[FORMULAS.md](docs/FORMULAS.md)**: Explicação de todas as fórmulas utilizadas
- **[GUIA_USO.md](docs/GUIA_USO.md)**: Guia passo a passo para iniciantes

### Recursos Úteis

#### Documentações Oficiais Microsoft
- [Documentação do Excel](https://support.microsoft.com/pt-br/excel)
- [Fórmulas e Funções do Excel](https://support.microsoft.com/pt-br/office/formulas-e-funcoes-do-excel)
- [Validação de Dados no Excel](https://support.microsoft.com/pt-br/office/aplicar-validação-de-dados-a-celulas)

#### GitHub e Documentação
- [GitHub Quick Start](https://github.com/digitalinnovationone/github-quickstart)
- [GitBook: Formação GitHub Certification](https://aline-antunes.gitbook.io/formacao-fundamentos-github)
- [Documentação do GitHub](https://docs.github.com/pt)
- [GitHub Markdown](https://docs.github.com/pt/get-started/writing-on-github)

#### Receita Federal
- [Portal da Receita Federal - IRPF 2025](https://www.gov.br/receitafederal/pt-br/assuntos/meu-imposto-de-renda)
- [Perguntas e Respostas IRPF](https://www.gov.br/receitafederal/pt-br/centrais-de-conteudo/publicacoes/perguntas-e-respostas/dirpf)

---

## 👨‍💻 Autor

**Marcus Vasconcellos**

<div align="left">

[![GitHub](https://img.shields.io/badge/GitHub-celloweb--ai-181717?style=for-the-badge&logo=github)](https://github.com/celloweb-ai)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Marcus_Vasconcellos-0077B5?style=for-the-badge&logo=linkedin)](https://www.linkedin.com/in/marcusvasconcellos)
[![Email](https://img.shields.io/badge/Email-marcus@vasconcellos.net.br-D14836?style=for-the-badge&logo=gmail&logoColor=white)](mailto:marcus@vasconcellos.net.br)

</div>

### Sobre Mim

Engenheiro de Eletrônica e Computação com MBA, com mais de 20 anos de experiência em liderança de projetos complexos de automação industrial e cibersegurança. Atualmente trabalhando na [@Prio3](https://github.com/Prio3) e desenvolvendo projetos de IA generativa e automação.

---

## 🎓 Sobre o Desafio DIO

Este projeto foi desenvolvido como parte do **DIO Bootcamp** (Digital Innovation One), aplicando os conceitos aprendidos em:

- ✅ Organização e estruturação de dados
- ✅ Criação de interfaces funcionais e amigáveis
- ✅ Documentação técnica profissional
- ✅ Controle de versão com Git e GitHub
- ✅ Boas práticas de desenvolvimento

### Objetivos de Aprendizagem Alcançados

✅ Aplicar conceitos aprendidos em ambiente prático
✅ Documentar processos técnicos de forma clara e estruturada
✅ Utilizar o GitHub como ferramenta para compartilhamento de documentação técnica
✅ Criar solução robusta com interface amigável
✅ Implementar validações e automações no Excel

---

## 📝 Licença

Este projeto está licenciado sob a **Licença MIT** - veja o arquivo [LICENSE](LICENSE) para mais detalhes.

```
MIT License

Copyright (c) 2025 Marcus Vasconcellos

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

---

## ⭐ Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para:

1. Fazer um fork do projeto
2. Criar uma branch para sua feature (`git checkout -b feature/MinhaFeature`)
3. Commit suas mudanças (`git commit -m 'Adiciona MinhaFeature'`)
4. Push para a branch (`git push origin feature/MinhaFeature`)
5. Abrir um Pull Request

### Sugestões de Melhorias

- 📄 Adicionar mais categorias de rendimentos e deduções
- 📊 Criar gráficos visuais no Dashboard
- 🔒 Implementar proteção de células com senha
- 📅 Adicionar calendário de prazos da Receita Federal
- 💾 Criar função de exportação para PDF
- 🌐 Desenvolver versão web da ferramenta

---

## 📞 Contato e Suporte

Se tiver dúvidas, sugestões ou encontrar problemas:

- **Abra uma issue**: [GitHub Issues](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues)
- **Email**: marcus@vasconcellos.net.br
- **LinkedIn**: [Marcus Vasconcellos](https://www.linkedin.com/in/marcusvasconcellos)

---

## 🚀 Próximos Passos

- [ ] Adicionar vídeo tutorial no YouTube
- [ ] Criar versão em Google Sheets
- [ ] Implementar macros VBA avançadas
- [ ] Adicionar exportação para formato IRPF
- [ ] Criar templates por tipo de declaração (simplificada/completa)
- [ ] Desenvolver calculadora de imposto com todas as alíquotas

---

<div align="center">

**Desenvolvido com 💜 por [Marcus Vasconcellos](https://github.com/celloweb-ai)**

**Projeto DIO Bootcamp 2025** | [Digital Innovation One](https://www.dio.me/)

Se este projeto foi útil para você, considere dar uma ⭐!

[![Star on GitHub](https://img.shields.io/github/stars/celloweb-ai/controle-ir-2025-excel-dio-challenge?style=social)](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge)

</div>
