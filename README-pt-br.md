[English version](README.md)

[English version](README.md)

# 📊 AutPlanilhas — Automação de Relatórios de Vendas

Uma ferramenta Python de nível profissional que consolida CSVs de vendas mensais, enriquece-os com dados de clientes e gera relatórios Excel refinados com gráficos e KPIs.

![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

---

## ✨ Funcionalidades

| Funcionalidade | Descrição |
|---------|-------------|
| **Ingestão de múltiplos CSVs** | Lê todos os arquivos `.csv` de uma pasta e os mescla automaticamente |
| **Limpeza inteligente de dados** | Trata moeda brasileira (`R$ 1.234,56`), múltiplos formatos de data e valores ausentes |
| **Enriquecimento de clientes** | Junta as vendas com uma lista mestre de clientes por `ID_Cliente` |
| **Cálculo de faturamento** | Calcula `Faturamento = Preço × Quantidade` por linha |
| **Resumo executivo** | KPIs: faturamento total, total de itens, ticket médio |
| **Detalhamento regional** | Faturamento por região com percentual de participação + gráfico de colunas |
| **Rankings** | Top 10 produtos e Top 10 clientes por faturamento |
| **Saída Excel profissional** | Painéis congelados, filtros automáticos, formatação de moeda, gráficos incorporados |
| **CLI e arquivo de configuração** | Uso flexível via argumentos de linha de comando ou config JSON |
| **Logs estruturados** | Logs no console + arquivo opcional com timestamps |

---

## 🚀 Início Rápido

### 1. Instale as dependências

```bash
pip install -r requirements.txt
```

### 2. Gere dados de exemplo (opcional)

```bash
python criar_dados_exemplo.py
```

Isso cria:
- `Relatórios_Mensais/janeiro.csv`
- `Relatórios_Mensais/fevereiro.csv`
- `Base_Clientes.xlsx`

### 3. Execute a automação

```bash
python automacao_relatorios.py
```

Saída: `Relatorio_Consolidado.xlsx`

---

## 📁 Estrutura do Projeto

```
AutPlanilhas/
├── automacao_relatorios.py   # Script principal de automação
├── criar_dados_exemplo.py    # Gerador de dados de exemplo
├── requirements.txt          # Dependências do Python
├── config.exemplo.json       # Exemplo de arquivo de configuração
├── Base_Clientes.xlsx        # Dados mestres de clientes (entrada)
├── Relatórios_Mensais/       # CSVs de vendas mensais (entrada)
│   ├── janeiro.csv
│   └── fevereiro.csv
└── Relatorio_Consolidado.xlsx  # Relatório gerado (saída)
```

---

## 📋 Formatos de Arquivos de Entrada

### CSV de Vendas (`Relatórios_Mensais/*.csv`)

| Coluna | Tipo | Exemplo |
|--------|------|---------|
| `ID_Cliente` | Inteiro | `101` |
| `Produto` | Texto | `Notebook Dell` |
| `Preço` | Moeda | `R$ 3.500,00` ou `3500.00` |
| `Quantidade` | Inteiro | `2` |
| `Data` | Data | `15/03/2026` ou `2026-03-15` |

> **Dica:** Os nomes das colunas são flexíveis — `preco`, `Preço`, `valor` são todos mapeados para Preço.

### Excel de Clientes (`Base_Clientes.xlsx`)

| Coluna | Tipo | Exemplo |
|--------|------|---------|
| `ID_Cliente` | Inteiro | `101` |
| `Nome` | Texto | `Tech Solutions` |
| `Região` | Texto | `Sudeste` |
| `Estado` | Texto (opcional) | `SP` |
| `Email` | Texto (opcional) | `contact@tech.com` |

---

## 📊 Relatório de Saída

O arquivo Excel gerado contém **5 planilhas**:

| Planilha | Conteúdo |
|-------|----------|
| **Resumo** | Resumo executivo com KPIs (faturamento total, itens, ticket médio) |
| **Dados Consolidados** | Todos os registros de vendas enriquecidos com dados dos clientes |
| **Faturamento por Região** | Faturamento por região + gráfico de colunas |
| **Top Produtos** | Top 10 produtos por faturamento + gráfico de barras |
| **Top Clientes** | Top 10 clientes por faturamento |

---

## ⚙️ Opções da CLI

```bash
python automacao_relatorios.py [OPÇÕES]
```

| Opção | Padrão | Descrição |
|--------|---------|-------------|
| `--input-dir` | `Relatórios_Mensais` | Pasta contendo os CSVs de vendas |
| `--clientes` | `Base_Clientes.xlsx` | Arquivo Excel mestre de clientes |
| `--output` | `Relatorio_Consolidado.xlsx` | Caminho do relatório de saída |
| `--config` | *(nenhum)* | Arquivo de configuração JSON (substitui os padrões) |
| `--log-level` | `INFO` | Nível de verbosidade dos logs: `DEBUG`, `INFO`, `WARNING`, `ERROR` |
| `--log-file` | *(nenhum)* | Caminho para salvar o arquivo de log |

### Exemplos

```bash
# Caminhos personalizados
python automacao_relatorios.py \
  --input-dir "C:\Sales\2026" \
  --clientes "C:\Data\Customers.xlsx" \
  --output "C:\Reports\Q1_2026.xlsx"

# Usando um arquivo de configuração
python automacao_relatorios.py --config config.json

# Com log em arquivo
python automacao_relatorios.py --log-file logs/automation.log --log-level DEBUG
```

---

## 🔧 Arquivo de Configuração

Crie um `config.json` com base no `config.exemplo.json`:

```json
{
  "input_dir": "Relatórios_Mensais",
  "clientes_xlsx": "Base_Clientes.xlsx",
  "output_xlsx": "Relatorio_Consolidado.xlsx",
  "log_level": "INFO",
  "log_file": "logs/automacao.log"
}
```

---

## 🏢 Uso Empresarial

### Execução Agendada (Agendador de Tarefas do Windows)

```powershell
python "C:\AutPlanilhas\automacao_relatorios.py" --config "C:\AutPlanilhas\config.json"
```

### Integração com ERP/CRM

1. Exporte os dados de vendas como CSV do seu sistema
2. Exporte a lista de clientes como Excel
3. Execute a automação
4. Distribua ou envie o relatório

---

## 🛠️ Requisitos

- Python 3.9+
- pandas ≥ 2.0
- openpyxl ≥ 3.1
- XlsxWriter ≥ 3.1

---

## 📄 Licença

Licença MIT — sinta-se à vontade para usar, modificar e distribuir.

---

## 🤝 Contribuindo

1. Faça um fork do repositório
2. Crie uma branch de funcionalidade (`git checkout -b feature/amazing-feature`)
3. Faça commit das suas alterações (`git commit -m 'Add amazing feature'`)
4. Envie para a branch (`git push origin feature/amazing-feature`)
5. Abra um Pull Request

---

## 📬 Suporte

Para dúvidas ou problemas, abra uma issue no GitHub ou entre em contato com o mantenedor.
