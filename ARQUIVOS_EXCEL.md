# 📁 Arquivos Excel do Projeto

## ⚠️ Aviso Importante

Devido a limitações da API do GitHub para upload de arquivos binários grandes via interface programada, os arquivos Excel originais não puderam ser carregados completos diretamente no repositório.

## 🚀 Como Obter os Arquivos Completos

Você tem **3 opções** para obter os arquivos Excel completos:

### Opção 1: Download dos Arquivos Originais (Recomendado)

Os arquivos originais estão disponíveis nos links oficiais da DIO:

1. **projeto_completo.xlsx**
   - Link: https://hermes.dio.me/files/assets/9b65e108-7ed6-456c-bb6b-e66827b322aa.xlsx
   - Tamanho: ~76 KB
   - Descrição: Planilha principal com todas as abas (TITULAR, INFORMES, NOTAS, TABELAS)

2. **bancos_apoio.xlsx**
   - Link: https://hermes.dio.me/files/assets/84cd23cc-7cb3-4219-9805-f99069dbbb33.xlsx
   - Tamanho: ~17 KB
   - Descrição: Planilha auxiliar com lista de 51 bancos brasileiros

3. **script_de_alinhamentos.txt**
   - Link: https://hermes.dio.me/files/assets/71335e12-6408-4a16-b42a-aef00053c0ff.txt
   - Já disponível neste repositório
   - Código VBA para alinhamento de ícones

### Opção 2: Download via GitHub Release

🔗 Acesse a seção [Releases](https://github.com/celloweb-ai/organizador-declaracao-imposto-renda/releases) deste repositório para baixar os arquivos.

*(Nota: Release será criado em breve com os arquivos completos)*

### Opção 3: Clonar Repositório e Executar Script

Se você tem os arquivos Excel localmente:

```bash
# Clone o repositório
git clone https://github.com/celloweb-ai/organizador-declaracao-imposto-renda.git
cd organizador-declaracao-imposto-renda

# Coloque os arquivos Excel no diretório
# - projeto_completo.xlsx
# - bancos_apoio.xlsx

# Execute o script Python para atualizar
python update_excel_files.py
```

## 📊 Estrutura dos Arquivos Excel

### projeto_completo.xlsx

Este é o arquivo principal do organizador e contém 4 abas:

#### 1. 👤 TITULAR
- Dados pessoais do declarante
- Campos: nome, CPF, nascimento, título de eleitor, cônjuge
- Endereço completo e contatos
- Indicações de alterações e status

#### 2. 🏦 INFORMES
- Registro de até 3 contas bancárias
- Lista suspensa com validação de bancos
- Cálculo automático do total
- Campo para anexar documentos PDF

#### 3. 💰 NOTAS
- Registro de entradas de receitas mês a mês
- Campos: data, categoria e valor
- Formato de moeda brasileira

#### 4. 📊 TABELAS
- Tabela de apoio com lista de bancos
- 51 instituições financeiras brasileiras
- Utilizada para validação de dados nas outras abas

### bancos_apoio.xlsx

Arquivo auxiliar contendo:
- Lista completa de 51 bancos brasileiros
- Código e nome de cada banco
- Utilizado para referência e validações

**Bancos incluídos:**
- Banco do Brasil (1)
- Caixa Econômica Federal (104)
- Bradesco (237)
- Santander (33)
- Itaú Unibanco (341)
- Nubank (260)
- Inter (77)
- C6 Bank (336)
- PagBank (290)
- PicPay (380)
- E outros 41 bancos

## 🛠️ Especificações Técnicas

### projeto_completo.xlsx
- **Tamanho**: 76.832 bytes (75 KB)
- **Formato**: Microsoft Excel (.xlsx)
- **Versão**: Excel 2016+
- **Abas**: 4 (TITULAR, INFORMES, NOTAS, TABELAS)
- **Fórmulas**: Cálculos automáticos, validações
- **Recursos**: Hiperlinks, listas suspensas, formatação condicional

### bancos_apoio.xlsx
- **Tamanho**: 17.351 bytes (17 KB)
- **Formato**: Microsoft Excel (.xlsx)
- **Versão**: Excel 2016+
- **Registros**: 51 bancos brasileiros
- **Estrutura**: Tabela simples com código e nome

## ❓ Perguntas Frequentes

### Por que os arquivos Excel não estão no repositório?

Devido a limitações da API do GitHub ao fazer upload programado de arquivos binários grandes, os arquivos precisam ser baixados dos links originais ou adicionados manualmente ao repositório via interface web ou git CLI.

### Como adicionar os arquivos manualmente ao repositório?

Se você tem os arquivos e deseja contribuir:

```bash
# Clone o repositório
git clone https://github.com/celloweb-ai/organizador-declaracao-imposto-renda.git
cd organizador-declaracao-imposto-renda

# Copie os arquivos Excel para o diretório
cp /caminho/para/projeto_completo.xlsx .
cp /caminho/para/bancos_apoio.xlsx .

# Adicione ao git
git add projeto_completo.xlsx bancos_apoio.xlsx
git commit -m "feat: Adiciona arquivos Excel completos"
git push origin main
```

### Os arquivos contêm macros?

Os arquivos Excel principais não contêm macros VBA integradas. O script VBA (`script_de_alinhamentos.txt`) é fornecido separadamente para ser adicionado manualmente pelo usuário, caso necessário.

### Posso usar os arquivos em versões antigas do Excel?

Recomenda-se Excel 2016 ou superior. Versões anteriores podem ter compatibilidade limitada com alguns recursos, como validações avançadas e formatações.

## 📞 Suporte

Se você encontrar problemas para baixar ou usar os arquivos:

1. Verifique os links da DIO (Opção 1)
2. Consulte a documentação completa no [README.md](README.md)
3. Abra uma [issue](https://github.com/celloweb-ai/organizador-declaracao-imposto-renda/issues) neste repositório

## 🔗 Links Úteis

- [README Principal](README.md)
- [Script de Alinhamento VBA](script_de_alinhamentos.txt)
- [Página do Desafio DIO](https://web.dio.me/track/santander-excel-com-inteligencia-artificial-2-semestre)
- [Bootcamp Santander/DIO](https://www.dio.me/)

---

<div align="center">
  <strong>💜 Desenvolvido por Marcus Vasconcellos</strong><br>
  <em>Bootcamp Santander Excel com IA - DIO 2025</em>
</div>