"""Gera dados de exemplo (mais realistas) para testar a automação.

- CSVs com variações de formato de moeda (R$ 1.234,56 / 1234.56)
- Datas em formatos diferentes
- Inclui um ID_Cliente inexistente para validar avisos
"""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def main() -> int:
    pasta = Path("Relatórios_Mensais")
    pasta.mkdir(exist_ok=True)

    # Janeiro
    df_jan = pd.DataFrame({
        "ID_Cliente": [101, 102, 103, 104, 999],
        "Produto": ["Notebook", "Mouse", "Teclado", "Monitor", "Headset"],
        "Preço": ["R$ 3.500,00", 89.90, "199,00", 1200.00, None],
        "Quantidade": [2, 5, 3, 1, 2],
        "Data": ["15/01/2024", "2024-01-16", "18/01/2024", "2024-01-20", "2024-01-22"],
    })
    df_jan.to_csv(pasta / "janeiro.csv", index=False, encoding="utf-8-sig")

    # Fevereiro
    df_fev = pd.DataFrame({
        "ID_Cliente": [101, 103, 102, 105, 104],
        "Produto": ["SSD", "Webcam", "Mousepad", "Cabo HDMI", "Caixa de Som"],
        "Preço": [450.00, "R$ 299,00", 49.90, "35,00", 180.00],
        "Quantidade": [3, 2, 4, 10, 2],
        "Data": ["2024-02-05", "10/02/2024", "2024-02-12", "15/02/2024", "2024-02-20"],
    })
    df_fev.to_csv(pasta / "fevereiro.csv", index=False, encoding="utf-8-sig")

    # Base de clientes
    df_clientes = pd.DataFrame({
        "ID_Cliente": [101, 102, 103, 104, 105],
        "Nome": ["Tech Solutions", "InfoShop", "DataCenter", "MegaStore", "ByteCompany"],
        "Região": ["Sudeste", "Sul", "Nordeste", "Sudeste", "Sul"],
        "Estado": ["SP", "PR", "BA", "RJ", "RS"],
        "Email": [
            "contato@tech.com",
            "vendas@infoshop.com",
            "data@center.com",
            "mega@store.com",
            "byte@company.com",
        ],
    })
    df_clientes.to_excel("Base_Clientes.xlsx", index=False)

    print("Dados de exemplo criados com sucesso!")
    print(f"- {pasta / 'janeiro.csv'}")
    print(f"- {pasta / 'fevereiro.csv'}")
    print("- Base_Clientes.xlsx")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
