# 🤝 Guia de Contribuição

Obrigado por considerar contribuir com o **Controle IR 2025**! 🎉

Este documento fornece diretrizes para contribuir com o projeto de forma eficiente e organizada.

---

## 📜 Índice

- [Código de Conduta](#código-de-conduta)
- [Como Posso Contribuir?](#como-posso-contribuir)
- [Reportando Bugs](#reportando-bugs)
- [Sugerindo Melhorias](#sugerindo-melhorias)
- [Contribuindo com Código](#contribuindo-com-código)
- [Processo de Pull Request](#processo-de-pull-request)
- [Padrões de Código](#padrões-de-código)
- [Documentação](#documentação)

---

## 🤝 Código de Conduta

### Nosso Compromisso

Este projeto segue um código de conduta para garantir um ambiente acolhedor e respeitoso para todos.

### Comportamentos Esperados

- ✅ Usar linguagem acolhedora e inclusiva
- ✅ Respeitar diferentes pontos de vista
- ✅ Aceitar críticas construtivas
- ✅ Focar no que é melhor para a comunidade
- ✅ Mostrar empatia com outros membros

### Comportamentos Inaceitáveis

- ❌ Linguagem ou imagens sexualizadas
- ❌ Comentários insultuosos ou depreciativos
- ❌ Assédio público ou privado
- ❌ Publicar informações privadas de terceiros
- ❌ Qualquer conduta inapropriada em ambiente profissional

---

## 💡 Como Posso Contribuir?

### Tipos de Contribuições

Você pode contribuir de várias formas:

1. **🐛 Reportando bugs** - Encontrou um erro?
2. **💡 Sugerindo melhorias** - Tem uma ideia legal?
3. **📝 Melhorando documentação** - Docs podem sempre melhorar
4. **💻 Contribuindo código** - VBA, fórmulas, macros
5. **✅ Testando** - Ajude a validar as planilhas
6. **🎨 Design** - Melhore a interface e layout
7. **🌎 Traduzindo** - Ajude em outras línguas

---

## 🐛 Reportando Bugs

### Antes de Reportar

1. **Verifique se já foi reportado:**
   - Procure nas [issues existentes](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues)

2. **Certifique-se de que é um bug:**
   - Teste em ambiente limpo
   - Verifique a documentação

### Como Reportar um Bug

Crie uma [nova issue](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues/new) com:

**Título claro e descritivo:**
```
🐛 Fórmula de cálculo retorna valor incorreto na aba Dashboard
```

**Conteúdo da issue:**

```markdown
## Descrição do Bug
Breve descrição do problema...

## Passos para Reproduzir
1. Abra a planilha Controle_IR_2025.xlsx
2. Vá para aba 'Dashboard'
3. Insira valor X no campo Y
4. Observe o erro Z

## Comportamento Esperado
O que deveria acontecer...

## Comportamento Atual
O que realmente acontece...

## Screenshots
(Se aplicável, adicione capturas de tela)

## Ambiente
- **Excel**: 2019 / 2021 / Microsoft 365
- **Sistema Operacional**: Windows 11 / macOS Sonoma
- **Versão da Planilha**: v1.0.0

## Informações Adicionais
Qualquer outra informação relevante...
```

---

## 💡 Sugerindo Melhorias

### Antes de Sugerir

1. **Verifique a lista de sugestões existentes**
2. **Confirme que faz sentido para o projeto**
3. **Pense na implementação**

### Como Sugerir uma Melhoria

Crie uma [nova issue](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues/new) com:

```markdown
## 💡 Sugestão de Melhoria

### Problema Atual
Descreva o problema ou limitação atual...

### Solução Proposta
Descreva sua sugestão de melhoria...

### Alternativas Consideradas
Outras soluções que você pensou...

### Benefícios
- Benefício 1
- Benefício 2

### Desafios
- Desafio 1 (se houver)

### Exemplos
(Mockups, código de exemplo, etc.)
```

---

## 💻 Contribuindo com Código

### Configuração do Ambiente

1. **Fork o repositório:**
   ```bash
   # Via interface web ou
   gh repo fork celloweb-ai/controle-ir-2025-excel-dio-challenge --clone
   ```

2. **Clone seu fork:**
   ```bash
   git clone https://github.com/SEU-USUARIO/controle-ir-2025-excel-dio-challenge.git
   cd controle-ir-2025-excel-dio-challenge
   ```

3. **Adicione o repositório original como upstream:**
   ```bash
   git remote add upstream https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge.git
   ```

4. **Instale ferramentas necessárias:**
   - Microsoft Excel 2016+
   - Editor VBA (integrado no Excel)
   - Git

### Workflow de Desenvolvimento

1. **Crie uma branch para sua feature:**
   ```bash
   git checkout -b feature/minha-nova-funcionalidade
   # ou
   git checkout -b fix/correcao-de-bug
   ```

2. **Faça suas alterações:**
   - Edite as planilhas Excel
   - Modifique scripts VBA
   - Atualize documentação

3. **Commit suas mudanças:**
   ```bash
   git add .
   git commit -m "feat: adiciona nova funcionalidade X"
   ```

4. **Mantenha sua branch atualizada:**
   ```bash
   git fetch upstream
   git rebase upstream/main
   ```

5. **Push para seu fork:**
   ```bash
   git push origin feature/minha-nova-funcionalidade
   ```

---

## 🔄 Processo de Pull Request

### Antes de Enviar

- [ ] Código testado e funcionando
- [ ] Documentação atualizada
- [ ] Commits seguem o padrão do projeto
- [ ] Branch atualizada com main

### Criando o Pull Request

1. **Acesse seu fork no GitHub**
2. **Clique em "New Pull Request"**
3. **Preencha o template:**

```markdown
## 📝 Descrição

Breve descrição das mudanças...

## 🎯 Tipo de Mudança

- [ ] 🐛 Bug fix (correção)
- [ ] ✨ Nova funcionalidade
- [ ] 📝 Documentação
- [ ] 🎨 Melhoria de UI/UX
- [ ] ♻️ Refatoração
- [ ] ⚡ Melhoria de performance

## ✅ Checklist

- [ ] Código testado
- [ ] Documentação atualizada
- [ ] Sem erros ou warnings
- [ ] Commits organizados
- [ ] Screenshots (se aplicável)

## 🖼️ Screenshots

(Adicione capturas de tela se houver mudanças visuais)

## 📝 Notas Adicionais

(Qualquer informação extra para os revisores)
```

### Após Enviar

- Aguarde feedback dos mantenedores
- Responda a comentários e sugestões
- Faça ajustes se solicitado
- Agradeça o tempo dos revisores! 🙏

---

## 📜 Padrões de Código

### Commits

Siga o padrão [Conventional Commits](https://www.conventionalcommits.org/):

```
tipo(escopo): descrição curta

[corpo opcional]

[rodapé opcional]
```

**Tipos:**
- `feat`: Nova funcionalidade
- `fix`: Correção de bug
- `docs`: Mudanças na documentação
- `style`: Formatação, espaços
- `refactor`: Refatoração de código
- `test`: Testes
- `chore`: Tarefas gerais

**Exemplos:**
```bash
feat(dashboard): adiciona gráfico de evolução patrimonial
fix(formulas): corrige cálculo de IR na aba Resumo
docs(readme): atualiza instruções de instalação
```

### Código VBA

```vba
' =========================================================================
' Nome do Módulo: NomeDoModulo
' Descrição: O que este módulo faz
' Autor: Seu Nome
' Data: YYYY-MM-DD
' =========================================================================

' Comente funções complexas
Function CalcularImpostoRenda(rendimento As Double) As Double
    ' Variáveis com nomes descritivos
    Dim baseCalculo As Double
    Dim aliquota As Double
    
    ' Lógica clara e comentada
    baseCalculo = rendimento - CalcularDeducoes()
    aliquota = ObterAliquota(baseCalculo)
    
    CalcularImpostoRenda = baseCalculo * aliquota
End Function
```

### Fórmulas Excel

- Use **nomes definidos** ao invés de referências diretas
- Comente fórmulas complexas com comenários de célula
- Organize em grupos lógicos
- Evite fórmulas muito longas (quebre em células auxiliares)

---

## 📚 Documentação

### Atualizando Docs

Quando modificar funcionalidades, atualize:

1. **README.md** - Se mudar uso básico
2. **docs/ESTRUTURA.md** - Se mudar estrutura
3. **docs/FORMULAS.md** - Se adicionar/modificar fórmulas
4. **docs/EXEMPLOS.md** - Se adicionar casos de uso
5. **docs/INSTALACAO.md** - Se mudar requisitos

### Estilo de Escrita

- Use **linguagem clara e simples**
- **Exemplos práticos** sempre que possível
- **Screenshots** para processos visuais
- **Tabelas** para comparar opções
- **Emojis** para facilitar leitura 🎉

---

## 🎖️ Reconhecimento

Contribuidores serão reconhecidos:

- 🏆 No README.md (seção Contributors)
- 🎉 Nas notas de release
- 💜 No hall da fama do projeto

---

## ❓ Dúvidas?

- 💬 [Abra uma Discussion](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/discussions)
- 🐛 [Crie uma Issue](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues)
- 📧 Entre em contato com os mantenedores

---

## 🚀 Próximos Passos

Após ler este guia:

1. ✅ Explore o [código do projeto](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge)
2. ✅ Leia a [documentação](docs/)
3. ✅ Escolha uma [issue para trabalhar](https://github.com/celloweb-ai/controle-ir-2025-excel-dio-challenge/issues)
4. ✅ Faça sua primeira contribuição! 🎉

---

**Obrigado por contribuir com o Controle IR 2025!** 💜

*Desenvolvido como parte do Desafio DIO Bootcamp*
