# 📊 Organizador de Declaração de Imposto de Renda

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![VBA](https://img.shields.io/badge/VBA-217346?style=for-the-badge&logo=microsoft&logoColor=white)
![DIO](https://img.shields.io/badge/DIO-Bootcamp-purple?style=for-the-badge)

## 📝 Descrição do Projeto

Este projeto foi desenvolvido como parte do desafio **"Criando Um Organizador de Declaração de Imposto de Renda"** do Bootcamp **Santander Excel com Inteligência Artificial** oferecido pela [Digital Innovation One (DIO)](https://web.dio.me).

O objetivo é criar uma ferramenta completa em Excel que auxilie na organização e reunião de informações essenciais para a declaração de imposto de renda, com uma interface amigável, validações automáticas e navegação intuitiva.

## ✨ Funcionalidades

### 1. 👤 Dados do Titular (Aba TITULAR)
- Cadastro completo de dados pessoais
- Campos incluem: nome, CPF, nascimento, título de eleitor, cônjuge, endereço completo, contatos
- Indicações de alterações da entrega anterior
- Campos para dependente cônjuge e residente do exterior

### 2. 🏦 Informes de Rendimentos Bancários (Aba INFORMES)
- Registro de até 3 bancos diferentes
- Lista completa de 51 instituições financeiras brasileiras com validação por lista suspensa
- Cálculo automático do total de valores
- Campo para anexar documentos (PDFs dos informes)
- Bancos disponíveis incluem: Banco do Brasil, Caixa Econômica, Bradesco, Santander, Itaú, Nubank, Inter, C6 Bank, PagBank, PicPay, entre outros

### 3. 💰 Notas Bancárias e Extratos (Aba NOTAS)
- Registro de entradas de receitas mês a mês
- Campos: data, categoria e valor
- Ideal para controle de holerites, rendimentos e outras receitas

### 4. 🧭 Navegação e Interface
- **Sistema LION APP** com identidade visual profissional
- Botões de navegação entre abas (PRÓXIMO / ANTERIOR)
- Links diretos entre seções usando hiperlinks do Excel
- Link de e-mail integrado para contato rápido
- Design clean com identificação "SYSTEM BY DIO 💜"

### 5. 🔧 Recursos Técnicos
- **Script VBA** para alinhamento automático de ícones
- Validações de dados com listas suspensas
- Aba de tabelas de apoio (TABELAS) com lista completa de bancos
- Formatoção condicional e cálculos automáticos

## 📂 Estrutura do Repositório

```
organizador-declaracao-imposto-renda/
│
├── projeto_completo.xlsx          # Planilha principal do organizador
├── bancos_apoio.xlsx              # Arquivo auxiliar com lista de bancos
├── script_de_alinhamentos.txt     # Código VBA para alinhamento de ícones
└── README.md                      # Documentação do projeto
```

## 🚀 Como Utilizar

### Pré-requisitos
- Microsoft Excel 2016 ou superior
- Habilitar macros caso utilize o script VBA

### Passo a Passo

1. **Download**
   ```bash
   git clone https://github.com/celloweb-ai/organizador-declaracao-imposto-renda.git
   ```

2. **Abrir o Arquivo Principal**
   - Abra `projeto_completo.xlsx` no Microsoft Excel
   - Se solicitado, habilite o conteúdo e macros

3. **Preencher os Dados**
   - Comece pela aba **TITULAR** com seus dados pessoais
   - Navegue para **INFORMES** e cadastre seus bancos
   - Use a aba **NOTAS** para registrar entradas mensais

4. **Navegação**
   - Use os botões **PRÓXIMO** e **ANTERIOR** para navegar
   - Clique nos links diretos das abas (exemplo: `<#TITULAR!C1>`)

5. **Script VBA (Opcional)**
   - Abra o Editor VBA (Alt + F11)
   - Insira um novo módulo
   - Cole o conteúdo de `script_de_alinhamentos.txt`
   - Execute a macro para ajustar posições de ícones

## 📚 Tecnologias Utilizadas

- **Microsoft Excel** - Plataforma principal
- **VBA (Visual Basic for Applications)** - Automações e scripts
- **Fórmulas Excel** - Cálculos e validações
- **Formatação Condicional** - Interface visual
- **Validação de Dados** - Listas suspensas e controles
- **Hiperlinks** - Navegação entre abas

## 🎯 Objetivos de Aprendizagem Alcançados

✅ Aplicar conceitos de Excel avançado em ambiente prático  
✅ Criar estruturas de dados validadas e organizadas  
✅ Implementar navegação intuitiva entre múltiplas abas  
✅ Documentar processos técnicos de forma clara e estruturada  
✅ Utilizar GitHub para compartilhamento de documentação técnica  
✅ Desenvolver interface amigável para usuário final  
✅ Aplicar automações com VBA  

## 💡 Destaques do Projeto

- **Interface Profissional**: Design limpo e organizado com identidade visual "LION APP"
- **Validações Robustas**: Lista completa de 51 bancos brasileiros para seleção validada
- **Navegação Intuitiva**: Sistema de links e botões para movimentação entre seções
- **Cálculos Automáticos**: Totalização automática de valores bancários
- **Organização Modular**: Separação clara entre dados pessoais, informes e notas
- **Integração de Anexos**: Campo para referência de documentos PDF

## 💬 Detalhes Técnicos

### Estrutura das Abas

#### TITULAR
- 15 campos de dados pessoais
- Validações de formato (CPF, telefone, e-mail)
- Opções sim/não para campos booleanos

#### INFORMES
- Até 3 registros de contas bancárias
- Lista suspensa com 51 bancos validados
- Fórmula de soma automática para total
- Campo de anexo para cada banco

#### NOTAS
- Tabela de entradas de receitas
- Campos: data, categoria e valor
- Formato de moeda brasileira (R$)

#### TABELAS
- Tabela de apoio com códigos e nomes de bancos
- Utilizada para validação de dados
- 51 instituições financeiras catalogadas

### Script VBA

O arquivo `script_de_alinhamentos.txt` contém uma macro VBA que:
- Localiza ícones por nome na planilha ativa
- Permite reposicionamento automático com coordenadas X e Y
- Exibe mensagens de confirmação ou erro
- Útil para ajustar layouts e elementos visuais

```vba
Sub MoverIconeParaPosicao()
    ' Move ícones para posições específicas
    ' Personalize nomeIconeProcurado, novaPosicaoX e novaPosicaoY
End Sub
```

## 👨‍💻 Autor

**Marcus Vasconcellos**

[![GitHub](https://img.shields.io/badge/GitHub-celloweb--ai-181717?style=for-the-badge&logo=github)](https://github.com/celloweb-ai)
[![LinkedIn](https://img.shields.io/badge/LinkedIn-Marcus%20Vasconcellos-0077B5?style=for-the-badge&logo=linkedin)](https://www.linkedin.com/in/marcusvasconcellos)
[![Email](https://img.shields.io/badge/Email-marcus%40vasconcellos.net.br-D14836?style=for-the-badge&logo=gmail&logoColor=white)](mailto:marcus@vasconcellos.net.br)

- 🏛️ **Empresa**: Prio3
- 📍 **Localização**: Rio de Janeiro, Brasil
- 🎓 **Formação**: Engenheiro Eletrônico e de Computação, MBA
- 🛠️ **Especializações**: Automação Industrial, Cibersegurança
- 🚀 **Experiência**: +20 anos em liderança técnica

## 🏫 Bootcamp DIO

Este projeto faz parte do **Bootcamp Santander Excel com Inteligência Artificial - 2º Semestre**, uma parceria entre:

- 💜 **[Digital Innovation One (DIO)](https://www.dio.me/)** - Plataforma de educação em tecnologia
- 🏦 **Banco Santander** - Patrocinador do bootcamp

### Competências Desenvolvidas

- Excel Avançado
- Validação de Dados
- Automação com VBA
- Design de Interface
- Documentação Técnica
- Controle de Versão (Git/GitHub)

## 📝 Licença

Este projeto foi desenvolvido para fins educacionais como parte do desafio do Bootcamp DIO. Sinta-se livre para utilizar e adaptar conforme suas necessidades.

## 🚀 Próximos Passos e Melhorias Futuras

- [ ] Adicionar aba para dependentes
- [ ] Incluir seção de bens e direitos
- [ ] Implementar cálculo de imposto estimado
- [ ] Criar relatório de resumo automático
- [ ] Adicionar mais validações de CPF e outros documentos
- [ ] Incluir gráficos de visualização de dados
- [ ] Exportar dados para formato PDF

## ⭐ Agradecimentos

Agradeço à **DIO** e ao **Banco Santander** pela oportunidade de aprendizado através deste bootcamp, e por disponibilizarem conteúdo de qualidade que permite o desenvolvimento de projetos práticos como este.

---

<div align="center">
  <strong>Desenvolvido com 💜 por Marcus Vasconcellos</strong><br>
  <em>Bootcamp Santander Excel com IA - DIO 2025</em>
</div>