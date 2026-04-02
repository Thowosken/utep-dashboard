"""
clean_excel.py
Lê um arquivo .xlsx/.xls, remove colunas completamente vazias e salva o resultado.

Se nenhum arquivo for informado, usa o nome do dia atual no formato DDMMzmm4.xls
(ex: hoje 02/04 -> 0204zmm4.xls).

Uso básico:
    python clean_excel.py                           # abre 0204zmm4.xls automaticamente
    python clean_excel.py arquivo.xls
    python clean_excel.py arquivo.xls --output resultado.xlsx

Relatório ZMM4 completo:
    python clean_excel.py --delete-rows 6 --filter-numeric Pedido
    python clean_excel.py --delete-rows 6 --filter-numeric Pedido --output limpo.xlsx
"""

import argparse
import sys
from datetime import date
from pathlib import Path

import pandas as pd


def default_filename() -> Path:
    """Gera o nome de arquivo padrão baseado na data atual: DDMMzmm4.xls"""
    today = date.today()
    name = today.strftime("%d%m") + "zmm4.xls"
    return Path(name)


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    """Strip espaços e converte células em branco/nan-string para pd.NA."""
    df = df.apply(
        lambda col: col.str.strip() if pd.api.types.is_string_dtype(col) else col
    )
    return df.replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA, "None": pd.NA})


def clean_excel(
    input_path: str,
    output_path: str | None = None,
    sheet_name: str | int = 0,
    threshold: float = 1.0,
    skip_rows: int = 0,
    filter_numeric: str | None = None,
) -> pd.DataFrame:
    """
    Processa o xlsx conforme os parâmetros:

    1. Pula as primeiras `skip_rows` linhas (a linha seguinte vira cabeçalho).
    2. Remove colunas cuja fração de vazios >= `threshold`.
    3. Se `filter_numeric` for informado, mantém apenas as linhas onde essa
       coluna contém um valor numérico (remove textos, vazios, cabeçalhos extras).
    4. Salva o resultado.

    Args:
        input_path:      Caminho do arquivo de entrada.
        output_path:     Caminho de saída. Se None, sobrescreve o original.
        sheet_name:      Nome ou índice da aba (padrão: primeira aba).
        threshold:       Fração de vazios para remover coluna (1.0 = 100%).
        skip_rows:       Nº de linhas a apagar no topo antes do cabeçalho.
        filter_numeric:  Nome da coluna cujos valores devem ser numéricos.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        sys.exit(f"Arquivo não encontrado: {input_path}")

    # --- 1. Leitura ---
    df = pd.read_excel(input_path, sheet_name=sheet_name, header=None, dtype=str)

    print(f"Arquivo : {input_path.name}")
    print(f"Linhas totais no arquivo: {len(df)}")

    # --- 1b. Apaga as primeiras N linhas e promove a linha seguinte a cabeçalho ---
    if skip_rows > 0:
        df = df.iloc[skip_rows:].reset_index(drop=True)  # apaga linhas 0..N-1
        df.columns = df.iloc[0]                           # linha N vira cabeçalho
        df = df.iloc[1:].reset_index(drop=True)           # remove a linha de cabeçalho dos dados
        df.columns.name = None
        print(f"Linhas após apagar as {skip_rows} primeiras: {len(df)}")

    # --- 2. Normaliza espaços / strings vazias ---
    df = _normalize(df)

    # --- 3. Remove colunas vazias ---
    cols_before = list(df.columns)
    empty_fraction = df.isna().mean()
    cols_to_drop = empty_fraction[empty_fraction >= threshold].index.tolist()
    df.drop(columns=cols_to_drop, inplace=True)

    removed_cols = [c for c in cols_before if c not in df.columns]
    print(f"\nColunas antes : {len(cols_before)}")
    print(f"Colunas depois: {len(df.columns)}")
    if removed_cols:
        print(f"Removidas ({len(removed_cols)}):")
        for c in removed_cols:
            print(f"  - {c}")
    else:
        print("Nenhuma coluna removida.")

    # --- 4. Filtra linhas com valor numérico na coluna indicada ---
    if filter_numeric:
        if filter_numeric not in df.columns:
            # Tenta busca case-insensitive
            match = [c for c in df.columns if c.strip().lower() == filter_numeric.lower()]
            if match:
                filter_numeric = match[0]
            else:
                available = ", ".join(df.columns.tolist())
                sys.exit(
                    f"Coluna '{filter_numeric}' não encontrada.\n"
                    f"Colunas disponíveis: {available}"
                )

        rows_before = len(df)

        def is_numeric(val):
            if pd.isna(val):
                return False
            try:
                float(str(val).replace(",", "."))
                return True
            except ValueError:
                return False

        mask = df[filter_numeric].apply(is_numeric)
        df = df[mask].reset_index(drop=True)

        print(f"\nFiltro numérico em '{filter_numeric}':")
        print(f"  Linhas antes : {rows_before}")
        print(f"  Linhas depois: {len(df)}")
        print(f"  Removidas    : {rows_before - len(df)}")

    # --- 5. Salva ---
    out = Path(output_path) if output_path else input_path
    df.to_excel(out, index=False)
    print(f"\nSalvo em: {out}")

    return df


def main():
    parser = argparse.ArgumentParser(
        description="Limpa um arquivo .xlsx: remove colunas vazias e filtra linhas."
    )
    parser.add_argument(
        "input", nargs="?", default=None,
        help=(
            "Caminho do arquivo .xls/.xlsx de entrada. "
            "Se omitido, usa DDMMzmm4.xls com a data de hoje "
            f"(hoje seria: {default_filename()})"
        ),
    )
    parser.add_argument(
        "--output", "-o", default=None,
        help="Caminho do arquivo de saída (opcional; padrão: sobrescreve o original)",
    )
    parser.add_argument(
        "--sheet", "-s", default=0,
        help="Nome ou índice da aba (padrão: 0 = primeira aba)",
    )
    parser.add_argument(
        "--threshold", "-t", type=float, default=1.0,
        help=(
            "Fração de vazios para remover coluna. "
            "1.0 = 100%% vazio (padrão). 0.9 = 90%% vazio."
        ),
    )
    parser.add_argument(
        "--delete-rows", type=int, default=0,
        dest="skip_rows",
        help="Nº de linhas a apagar no topo antes do cabeçalho (padrão: 0)",
    )
    parser.add_argument(
        "--filter-numeric", default=None,
        dest="filter_numeric",
        metavar="COLUNA",
        help=(
            "Nome da coluna: mantém apenas linhas onde o valor é numérico. "
            "Ex: --filter-numeric Pedido"
        ),
    )
    args = parser.parse_args()

    input_path = Path(args.input) if args.input else default_filename()
    if not args.input:
        print(f"Nenhum arquivo informado. Usando: {input_path}")

    sheet = args.sheet
    try:
        sheet = int(sheet)
    except ValueError:
        pass

    clean_excel(
        input_path=input_path,
        output_path=args.output,
        sheet_name=sheet,
        threshold=args.threshold,
        skip_rows=args.skip_rows,
        filter_numeric=args.filter_numeric,
    )


if __name__ == "__main__":
    main()
