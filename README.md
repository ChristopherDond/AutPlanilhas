# 📊 AutPlanilhas — Sales Report Automation

A professional-grade Python tool that consolidates monthly sales CSVs, enriches them with customer data, and generates polished Excel reports with charts and KPIs.

![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue)
![License: MIT](https://img.shields.io/badge/license-MIT-green)

---

## ✨ Features

| Feature | Description |
|---------|-------------|
| **Multi-file CSV ingestion** | Reads all `.csv` files from a folder and merges them automatically |
| **Smart data cleaning** | Handles Brazilian currency (`R$ 1.234,56`), multiple date formats, and missing values |
| **Customer enrichment** | Joins sales with a customer master list by `ID_Cliente` |
| **Revenue calculation** | Computes `Revenue = Price × Quantity` per line |
| **Executive summary** | KPIs: total revenue, total items, average ticket |
| **Regional breakdown** | Revenue by region with percentage share + column chart |
| **Top rankings** | Top 10 products and Top 10 customers by revenue |
| **Professional Excel output** | Freeze panes, auto-filters, currency formatting, embedded charts |
| **CLI & config file** | Flexible usage via command-line args or JSON config |
| **Structured logging** | Console + optional file logging with timestamps |

---

## 🚀 Quick Start

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Generate sample data (optional)

```bash
python criar_dados_exemplo.py
```

This creates:
- `Relatórios_Mensais/janeiro.csv`
- `Relatórios_Mensais/fevereiro.csv`
- `Base_Clientes.xlsx`

### 3. Run the automation

```bash
python automacao_relatorios.py
```

Output: `Relatorio_Consolidado.xlsx`

---

## 📁 Project Structure

```
AutPlanilhas/
├── automacao_relatorios.py   # Main automation script
├── criar_dados_exemplo.py    # Sample data generator
├── requirements.txt          # Python dependencies
├── config.exemplo.json       # Example configuration file
├── Base_Clientes.xlsx        # Customer master data (input)
├── Relatórios_Mensais/       # Monthly sales CSVs (input)
│   ├── janeiro.csv
│   └── fevereiro.csv
└── Relatorio_Consolidado.xlsx  # Generated report (output)
```

---

## 📋 Input File Formats

### Sales CSV (`Relatórios_Mensais/*.csv`)

| Column | Type | Example |
|--------|------|---------|
| `ID_Cliente` | Integer | `101` |
| `Produto` | String | `Notebook Dell` |
| `Preço` | Currency | `R$ 3.500,00` or `3500.00` |
| `Quantidade` | Integer | `2` |
| `Data` | Date | `15/03/2026` or `2026-03-15` |

> **Tip:** Column names are flexible — `preco`, `Preço`, `valor` all map to Price.

### Customer Excel (`Base_Clientes.xlsx`)

| Column | Type | Example |
|--------|------|---------|
| `ID_Cliente` | Integer | `101` |
| `Nome` | String | `Tech Solutions` |
| `Região` | String | `Sudeste` |
| `Estado` | String (optional) | `SP` |
| `Email` | String (optional) | `contact@tech.com` |

---

## 📊 Output Report

The generated Excel file contains **5 worksheets**:

| Sheet | Contents |
|-------|----------|
| **Resumo** | Executive summary with KPIs (total revenue, items, avg ticket) |
| **Dados Consolidados** | All sales records enriched with customer data |
| **Faturamento por Região** | Revenue by region + column chart |
| **Top Produtos** | Top 10 products by revenue + bar chart |
| **Top Clientes** | Top 10 customers by revenue |

---

## ⚙️ CLI Options

```bash
python automacao_relatorios.py [OPTIONS]
```

| Option | Default | Description |
|--------|---------|-------------|
| `--input-dir` | `Relatórios_Mensais` | Folder containing sales CSVs |
| `--clientes` | `Base_Clientes.xlsx` | Customer master Excel file |
| `--output` | `Relatorio_Consolidado.xlsx` | Output report path |
| `--config` | *(none)* | JSON config file (overrides defaults) |
| `--log-level` | `INFO` | Logging verbosity: `DEBUG`, `INFO`, `WARNING`, `ERROR` |
| `--log-file` | *(none)* | Path to save log file |

### Examples

```bash
# Custom paths
python automacao_relatorios.py \
  --input-dir "C:\Sales\2026" \
  --clientes "C:\Data\Customers.xlsx" \
  --output "C:\Reports\Q1_2026.xlsx"

# Using a config file
python automacao_relatorios.py --config config.json

# With file logging
python automacao_relatorios.py --log-file logs/automation.log --log-level DEBUG
```

---

## 🔧 Configuration File

Create a `config.json` based on `config.exemplo.json`:

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

## 🏢 Enterprise Usage

### Scheduled Execution (Windows Task Scheduler)

```powershell
python "C:\AutPlanilhas\automacao_relatorios.py" --config "C:\AutPlanilhas\config.json"
```

### Integration with ERP/CRM

1. Export sales data as CSV from your system
2. Export customer list as Excel
3. Run the automation
4. Distribute or upload the report

---

## 🛠️ Requirements

- Python 3.9+
- pandas ≥ 2.0
- openpyxl ≥ 3.1
- XlsxWriter ≥ 3.1

---

## 📄 License

MIT License — feel free to use, modify, and distribute.

---

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## 📬 Support

For questions or issues, please open a GitHub issue or contact the maintainer.
