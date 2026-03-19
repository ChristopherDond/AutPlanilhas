"""Automação de relatórios (nível comercial)

Consolida CSVs mensais de vendas, cruza com a base de clientes e gera um Excel final com:
- Dados consolidados
- Resumo executivo (KPIs)
- Faturamento por região + gráfico
- Top produtos e top clientes

Uso rápido:
  python automacao_relatorios.py
  python automacao_relatorios.py --input-dir "Relatórios_Mensais" --clientes "Base_Clientes.xlsx" --output "Relatorio_Consolidado.xlsx"

Observação importante:
- O faturamento é calculado como (Preço * Quantidade).
"""

from __future__ import annotations

import argparse
import json
import logging
import sys
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, Optional

import pandas as pd


# ----------------------------- Config & Logging -----------------------------

@dataclass(frozen=True)
class AppConfig:
    input_dir: Path = Path("Relatórios_Mensais")
    clientes_xlsx: Path = Path("Base_Clientes.xlsx")
    output_xlsx: Path = Path("Relatorio_Consolidado.xlsx")
    config_json: Optional[Path] = None
    log_level: str = "INFO"
    log_file: Optional[Path] = None


def _setup_logger(level: str, log_file: Optional[Path]) -> logging.Logger:
    logger = logging.getLogger("autplanilhas")
    if logger.handlers:
        return logger

    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s", datefmt="%Y-%m-%d %H:%M:%S")

    ch = logging.StreamHandler(stream=sys.stdout)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    if log_file:
        log_file.parent.mkdir(parents=True, exist_ok=True)
        fh = logging.FileHandler(log_file, encoding="utf-8")
        fh.setFormatter(fmt)
        logger.addHandler(fh)

    return logger


def _load_json_config(path: Optional[Path], logger: logging.Logger) -> Dict[str, Any]:
    if not path:
        return {}
    if not path.exists():
        raise FileNotFoundError(f"Arquivo de configuração não encontrado: {path}")
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            raise ValueError("config.json precisa ser um objeto JSON")
        logger.info("Config carregada: %s", path)
        return data
    except json.JSONDecodeError as e:
        raise ValueError(f"config.json inválido: {e}") from e


def _apply_config_overrides(cfg: AppConfig, overrides: Dict[str, Any]) -> AppConfig:
    def _p(v: Any) -> Optional[Path]:
        if v is None:
            return None
        return Path(str(v))

    return AppConfig(
        input_dir=_p(overrides.get("input_dir")) or cfg.input_dir,
        clientes_xlsx=_p(overrides.get("clientes_xlsx")) or cfg.clientes_xlsx,
        output_xlsx=_p(overrides.get("output_xlsx")) or cfg.output_xlsx,
        config_json=cfg.config_json,
        log_level=str(overrides.get("log_level") or cfg.log_level),
        log_file=_p(overrides.get("log_file")) or cfg.log_file,
    )


# ----------------------------- Data Utilities ------------------------------

_COL_ALIASES = {
    "preco": "Preço",
    "preço": "Preço",
    "valor": "Preço",
    "qtde": "Quantidade",
    "qtd": "Quantidade",
    "quantidade": "Quantidade",
    "regiao": "Região",
    "região": "Região",
    "id_cliente": "ID_Cliente",
    "idcliente": "ID_Cliente",
    "cliente_id": "ID_Cliente",
}


def _normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    new_cols = {}
    for c in df.columns:
        key = str(c).strip()
        key_norm = (
            key.replace(" ", "_")
            .replace("-", "_")
            .lower()
        )
        new_cols[c] = _COL_ALIASES.get(key_norm, key)
    return df.rename(columns=new_cols)


def _parse_money_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")

    txt = (
        s.astype(str)
        .str.replace("\u00a0", " ", regex=False)
        .str.replace("R$", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(".", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip()
    )
    return pd.to_numeric(txt, errors="coerce")


def _read_csv_robusto(path: Path) -> pd.DataFrame:
    # Tenta auto-detecção de separador; cai para ; ou , se necessário.
    encodings = ["utf-8-sig", "utf-8", "latin1"]
    last_err: Optional[Exception] = None
    for enc in encodings:
        try:
            return pd.read_csv(path, encoding=enc, sep=None, engine="python")
        except Exception as e:
            last_err = e
    raise last_err  # type: ignore[misc]


def carregar_csvs(pasta: Path, logger: logging.Logger) -> pd.DataFrame:
    """Lê todos os CSVs de uma pasta e concatena em um único DataFrame."""
    if not pasta.exists() or not pasta.is_dir():
        raise FileNotFoundError(f"Pasta de relatórios não encontrada: {pasta}")

    arquivos_csv = sorted(pasta.glob("*.csv"))
    if not arquivos_csv:
        raise FileNotFoundError(f"Nenhum arquivo CSV encontrado em: {pasta}")

    dataframes = []
    for arquivo in arquivos_csv:
        df = _read_csv_robusto(arquivo)
        df = _normalize_columns(df)
        df["Arquivo_Origem"] = arquivo.name
        dataframes.append(df)
        logger.info("Carregado CSV: %s (%s registros)", arquivo.name, len(df))

    return pd.concat(dataframes, ignore_index=True)


def limpar_dados_vendas(df: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    df = _normalize_columns(df)

    required = {"ID_Cliente", "Produto", "Preço", "Quantidade", "Data"}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError(f"CSV(s) de vendas sem colunas obrigatórias: {', '.join(missing)}")

    out = df.copy()

    out["ID_Cliente"] = pd.to_numeric(out["ID_Cliente"], errors="coerce")
    out["Produto"] = out["Produto"].astype(str).str.strip()

    out["Preço"] = _parse_money_series(out["Preço"]).fillna(0.0)
    out["Quantidade"] = pd.to_numeric(out["Quantidade"], errors="coerce").fillna(0.0)
    out.loc[out["Quantidade"] < 0, "Quantidade"] = 0

    out["Data"] = pd.to_datetime(out["Data"], errors="coerce", dayfirst=True)
    out["Mes"] = out["Data"].dt.to_period("M").astype(str)

    out["Faturamento"] = (out["Preço"] * out["Quantidade"]).round(2)

    n_invalid_dates = int(out["Data"].isna().sum())
    if n_invalid_dates:
        logger.warning("%s linhas com Data inválida (ficaram em branco no Excel).", n_invalid_dates)

    logger.info("Dados de vendas limpos: %s registros", len(out))
    return out


def carregar_base_clientes(path: Path, logger: logging.Logger) -> pd.DataFrame:
    if not path.exists():
        raise FileNotFoundError(f"Base de clientes não encontrada: {path}")
    df = pd.read_excel(path)
    df = _normalize_columns(df)

    required = {"ID_Cliente", "Nome", "Região"}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError(f"Base de clientes sem colunas obrigatórias: {', '.join(missing)}")

    df["ID_Cliente"] = pd.to_numeric(df["ID_Cliente"], errors="coerce")
    df["Nome"] = df["Nome"].astype(str).str.strip()
    df["Região"] = df["Região"].astype(str).str.strip()

    logger.info("Base de clientes carregada: %s clientes", len(df))
    return df


def cruzar_dados(vendas: pd.DataFrame, clientes: pd.DataFrame, logger: logging.Logger) -> pd.DataFrame:
    out = pd.merge(vendas, clientes, on="ID_Cliente", how="left", suffixes=("", "_Cliente"))
    out["Cliente_Encontrado"] = out["Nome"].notna() & (out["Nome"].astype(str).str.len() > 0)

    n_missing = int((~out["Cliente_Encontrado"]).sum())
    if n_missing:
        logger.warning("%s linhas sem correspondência na base de clientes (ID_Cliente não encontrado).", n_missing)

    logger.info("Dados cruzados: %s registros", len(out))
    return out


def calcular_faturamento_regiao(df: pd.DataFrame) -> pd.DataFrame:
    if "Região" not in df.columns or "Faturamento" not in df.columns:
        raise ValueError("Colunas 'Região' e 'Faturamento' são necessárias para o cálculo")

    tmp = df.copy()
    tmp["Região"] = tmp["Região"].fillna("(Sem região)").astype(str)

    fat = (
        tmp.groupby("Região", dropna=False)["Faturamento"]
        .sum()
        .reset_index()
        .rename(columns={"Faturamento": "Faturamento_Total"})
        .sort_values("Faturamento_Total", ascending=False)
    )
    total = float(fat["Faturamento_Total"].sum())
    # Fração (0..1), pronta para formatação percentual no Excel.
    fat["Participacao_%"] = (fat["Faturamento_Total"] / total) if total else 0.0
    return fat


def calcular_top(df: pd.DataFrame, col: str, value_col: str, n: int = 10) -> pd.DataFrame:
    tmp = df.copy()
    tmp[col] = tmp[col].fillna("(Vazio)").astype(str)
    res = (
        tmp.groupby(col, dropna=False)[value_col]
        .sum()
        .reset_index()
        .sort_values(value_col, ascending=False)
        .head(n)
    )
    return res


def calcular_kpis(df: pd.DataFrame) -> Dict[str, Any]:
    total_faturamento = float(df["Faturamento"].sum())
    total_itens = float(df["Quantidade"].sum())
    total_linhas = int(len(df))
    ticket_medio = (total_faturamento / total_linhas) if total_linhas else 0.0
    return {
        "Total_Faturamento": total_faturamento,
        "Total_Itens": total_itens,
        "Total_Linhas": total_linhas,
        "Ticket_Medio_Por_Linha": ticket_medio,
    }


# ----------------------------- Excel Generation ----------------------------

def _autosize_columns(worksheet, df: pd.DataFrame, max_width: int = 45):
    for i, col in enumerate(df.columns):
        series = df[col].astype(str)
        width = max(len(str(col)), int(series.map(len).max() if len(series) else 0)) + 2
        worksheet.set_column(i, i, min(width, max_width))


def gerar_excel(df_consolidado: pd.DataFrame,
               df_faturamento: pd.DataFrame,
               df_top_produtos: pd.DataFrame,
               df_top_clientes: pd.DataFrame,
               kpis: Dict[str, Any],
               caminho_saida: Path,
               logger: logging.Logger) -> None:
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(caminho_saida, engine="xlsxwriter", datetime_format="yyyy-mm-dd") as writer:
        workbook = writer.book

        fmt_title = workbook.add_format({"bold": True, "font_size": 14})
        fmt_label = workbook.add_format({"bold": True})
        fmt_money = workbook.add_format({"num_format": "R$ #,##0.00"})
        fmt_int = workbook.add_format({"num_format": "0"})
        fmt_pct = workbook.add_format({"num_format": "0.00%"})

        # Resumo
        ws = workbook.add_worksheet("Resumo")
        writer.sheets["Resumo"] = ws
        ws.write("A1", "Resumo Executivo", fmt_title)
        ws.write("A3", "Gerado em:", fmt_label)
        ws.write("B3", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

        ws.write("A5", "Total faturamento:", fmt_label)
        ws.write_number("B5", float(kpis["Total_Faturamento"]), fmt_money)
        ws.write("A6", "Total itens:", fmt_label)
        ws.write_number("B6", float(kpis["Total_Itens"]), fmt_int)
        ws.write("A7", "Total linhas:", fmt_label)
        ws.write_number("B7", int(kpis["Total_Linhas"]), fmt_int)
        ws.write("A8", "Ticket médio (por linha):", fmt_label)
        ws.write_number("B8", float(kpis["Ticket_Medio_Por_Linha"]), fmt_money)

        ws.set_column("A:A", 28)
        ws.set_column("B:B", 22)

        # Dados Consolidados
        df_consolidado.to_excel(writer, sheet_name="Dados Consolidados", index=False)
        ws_dados = writer.sheets["Dados Consolidados"]
        ws_dados.freeze_panes(1, 0)
        ws_dados.autofilter(0, 0, len(df_consolidado), len(df_consolidado.columns) - 1)
        _autosize_columns(ws_dados, df_consolidado)

        # Formatos específicos (quando existirem colunas)
        col_idx = {c: i for i, c in enumerate(df_consolidado.columns)}
        if "Preço" in col_idx:
            ws_dados.set_column(col_idx["Preço"], col_idx["Preço"], 14, fmt_money)
        if "Faturamento" in col_idx:
            ws_dados.set_column(col_idx["Faturamento"], col_idx["Faturamento"], 16, fmt_money)
        if "Quantidade" in col_idx:
            ws_dados.set_column(col_idx["Quantidade"], col_idx["Quantidade"], 12, fmt_int)

        # Faturamento por Região
        df_faturamento.to_excel(writer, sheet_name="Faturamento por Região", index=False)
        ws_reg = writer.sheets["Faturamento por Região"]
        ws_reg.freeze_panes(1, 0)
        ws_reg.autofilter(0, 0, len(df_faturamento), len(df_faturamento.columns) - 1)
        _autosize_columns(ws_reg, df_faturamento)

        if "Faturamento_Total" in df_faturamento.columns:
            c = df_faturamento.columns.get_loc("Faturamento_Total")
            ws_reg.set_column(c, c, 18, fmt_money)
        if "Participacao_%" in df_faturamento.columns:
            c = df_faturamento.columns.get_loc("Participacao_%")
            ws_reg.set_column(c, c, 16, fmt_pct)

        # Gráfico de regiões
        chart_reg = workbook.add_chart({"type": "column"})
        n = len(df_faturamento)
        if n:
            chart_reg.add_series({
                "name": "Faturamento Total",
                "categories": ["Faturamento por Região", 1, 0, n, 0],
                "values": ["Faturamento por Região", 1, 1, n, 1],
                "fill": {"color": "#4472C4"},
                "data_labels": {"value": True, "num_format": "R$ #,##0.00"},
            })
            chart_reg.set_title({"name": "Faturamento por Região"})
            chart_reg.set_x_axis({"name": "Região"})
            chart_reg.set_y_axis({"name": "Faturamento (R$)", "num_format": "R$ #,##0.00"})
            chart_reg.set_style(10)
            chart_reg.set_size({"width": 640, "height": 380})
            ws_reg.insert_chart("E2", chart_reg)

        # Top Produtos
        df_top_produtos.to_excel(writer, sheet_name="Top Produtos", index=False)
        ws_prod = writer.sheets["Top Produtos"]
        ws_prod.freeze_panes(1, 0)
        ws_prod.autofilter(0, 0, len(df_top_produtos), len(df_top_produtos.columns) - 1)
        _autosize_columns(ws_prod, df_top_produtos)
        if "Faturamento" in df_top_produtos.columns:
            c = df_top_produtos.columns.get_loc("Faturamento")
            ws_prod.set_column(c, c, 18, fmt_money)

        chart_prod = workbook.add_chart({"type": "bar"})
        n = len(df_top_produtos)
        if n:
            chart_prod.add_series({
                "name": "Faturamento",
                "categories": ["Top Produtos", 1, 0, n, 0],
                "values": ["Top Produtos", 1, 1, n, 1],
                "fill": {"color": "#70AD47"},
                "data_labels": {"value": True, "num_format": "R$ #,##0.00"},
            })
            chart_prod.set_title({"name": "Top Produtos (Faturamento)"})
            chart_prod.set_x_axis({"name": "Faturamento (R$)", "num_format": "R$ #,##0.00"})
            chart_prod.set_y_axis({"name": "Produto"})
            chart_prod.set_style(10)
            chart_prod.set_size({"width": 640, "height": 420})
            ws_prod.insert_chart("D2", chart_prod)

        # Top Clientes
        df_top_clientes.to_excel(writer, sheet_name="Top Clientes", index=False)
        ws_cli = writer.sheets["Top Clientes"]
        ws_cli.freeze_panes(1, 0)
        ws_cli.autofilter(0, 0, len(df_top_clientes), len(df_top_clientes.columns) - 1)
        _autosize_columns(ws_cli, df_top_clientes)
        if "Faturamento" in df_top_clientes.columns:
            c = df_top_clientes.columns.get_loc("Faturamento")
            ws_cli.set_column(c, c, 18, fmt_money)

    logger.info("Arquivo Excel gerado: %s", caminho_saida)


# ---------------------------------- CLI -----------------------------------

def _parse_args(argv: Optional[Iterable[str]] = None) -> AppConfig:
    p = argparse.ArgumentParser(description="Consolida relatórios mensais e gera Excel profissional.")
    p.add_argument("--input-dir", default=str(AppConfig.input_dir), help="Pasta com CSVs mensais")
    p.add_argument("--clientes", default=str(AppConfig.clientes_xlsx), help="Arquivo Excel da base de clientes")
    p.add_argument("--output", default=str(AppConfig.output_xlsx), help="Arquivo Excel de saída")
    p.add_argument("--config", default=None, help="Arquivo JSON de configuração (opcional)")
    p.add_argument("--log-level", default=AppConfig.log_level, help="DEBUG, INFO, WARNING, ERROR")
    p.add_argument("--log-file", default=None, help="Caminho para salvar log (opcional)")

    a = p.parse_args(list(argv) if argv is not None else None)

    cfg = AppConfig(
        input_dir=Path(a.input_dir),
        clientes_xlsx=Path(a.clientes),
        output_xlsx=Path(a.output),
        config_json=Path(a.config) if a.config else None,
        log_level=str(a.log_level),
        log_file=Path(a.log_file) if a.log_file else None,
    )
    return cfg


def main(argv: Optional[Iterable[str]] = None) -> int:
    cfg = _parse_args(argv)
    logger = _setup_logger(cfg.log_level, cfg.log_file)

    try:
        overrides = _load_json_config(cfg.config_json, logger)
        cfg = _apply_config_overrides(cfg, overrides)
        logger.setLevel(getattr(logging, cfg.log_level.upper(), logging.INFO))

        logger.info("Iniciando automação")

        # Fallback para ambientes onde a pasta foi criada sem acento.
        if not cfg.input_dir.exists():
            alt = Path("Relatorios_Mensais")
            if alt.exists() and alt.is_dir():
                logger.warning("Pasta '%s' não encontrada; usando '%s'.", cfg.input_dir, alt)
                cfg = AppConfig(
                    input_dir=alt,
                    clientes_xlsx=cfg.clientes_xlsx,
                    output_xlsx=cfg.output_xlsx,
                    config_json=cfg.config_json,
                    log_level=cfg.log_level,
                    log_file=cfg.log_file,
                )

        logger.info("Entrada: %s | Clientes: %s | Saída: %s", cfg.input_dir, cfg.clientes_xlsx, cfg.output_xlsx)

        vendas_raw = carregar_csvs(cfg.input_dir, logger)
        vendas = limpar_dados_vendas(vendas_raw, logger)
        clientes = carregar_base_clientes(cfg.clientes_xlsx, logger)

        consolidado = cruzar_dados(vendas, clientes, logger)
        fat_reg = calcular_faturamento_regiao(consolidado)

        # Rankings
        top_prod = calcular_top(consolidado, "Produto", "Faturamento", n=10)
        if "Nome" in consolidado.columns:
            top_cli = calcular_top(consolidado, "Nome", "Faturamento", n=10)
        else:
            top_cli = pd.DataFrame({"Nome": [], "Faturamento": []})

        kpis = calcular_kpis(consolidado)

        gerar_excel(
            df_consolidado=consolidado,
            df_faturamento=fat_reg,
            df_top_produtos=top_prod,
            df_top_clientes=top_cli,
            kpis=kpis,
            caminho_saida=cfg.output_xlsx,
            logger=logger,
        )

        logger.info("Concluído com sucesso")
        return 0

    except Exception as e:
        logger.error("Falha na automação: %s", e)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
