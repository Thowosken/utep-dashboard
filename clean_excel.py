"""
clean_excel.py
Lê um arquivo .xlsx, remove colunas completamente vazias e salva o resultado.

Uso:
    python clean_excel.py arquivo.xlsx
    python clean_excel.py arquivo.xlsx --output resultado.xlsx
    python clean_excel.py arquivo.xlsx --sheet "Planilha1"
    python clean_excel.py arquivo.xlsx --threshold 0.9   # remove colunas com >= 90% vazios
"""

import argparse
import sys
from pathlib import Path

import pandas as pd


def clean_empty_columns(
    input_path: str,
    output_path: str | None = None,
    sheet_name: str | int = 0,
    threshold: float = 1.0,
) -> pd.DataFrame:
    """
    Lê o xlsx, remove colunas vazias e salva.

    Args:
        input_path:  Caminho do arquivo de entrada.
        output_path: Caminho de saída. Se None, sobrescreve o arquivo original.
        sheet_name:  Nome ou índice da aba (padrão: primeira aba).
        threshold:   Fração mínima de valores vazios para remover a coluna.
                     1.0 = remove só se 100% vazio (padrão).
                     0.9 = remove se >= 90% dos valores são vazios.

    Returns:
        DataFrame limpo.
    """
    input_path = Path(input_path)
    if not input_path.exists():
        sys.exit(f"Arquivo não encontrado: {input_path}")

    df = pd.read_excel(input_path, sheet_name=sheet_name, dtype=str)

    total_rows = len(df)
    cols_before = list(df.columns)

    # Normaliza células: espaços em branco e strings "nan" viram NaN real
    # pd.api.types.is_string_dtype cobre tanto object quanto StringDtype
    df = df.apply(
        lambda col: col.str.strip() if pd.api.types.is_string_dtype(col) else col
    )
    df = df.replace({"": pd.NA, "nan": pd.NA, "NaN": pd.NA, "None": pd.NA})

    # Calcula a fração de valores vazios por coluna
    empty_fraction = df.isna().mean()
    cols_to_drop = empty_fraction[empty_fraction >= threshold].index.tolist()

    df.drop(columns=cols_to_drop, inplace=True)
    cols_after = list(df.columns)

    removed = [c for c in cols_before if c not in cols_after]

    print(f"Arquivo:    {input_path.name}")
    print(f"Linhas:     {total_rows}")
    print(f"Colunas antes:  {len(cols_before)}")
    print(f"Colunas depois: {len(cols_after)}")
    if removed:
        print(f"Removidas ({len(removed)}):")
        for c in removed:
            print(f"  - {c}")
    else:
        print("Nenhuma coluna removida.")

    out = Path(output_path) if output_path else input_path
    df.to_excel(out, index=False)
    print(f"\nSalvo em: {out}")

    return df


def main():
    parser = argparse.ArgumentParser(
        description="Remove colunas vazias de um arquivo .xlsx"
    )
    parser.add_argument("input", help="Caminho do arquivo .xlsx de entrada")
    parser.add_argument(
        "--output", "-o", default=None, help="Caminho do arquivo de saída (opcional)"
    )
    parser.add_argument(
        "--sheet", "-s", default=0,
        help="Nome ou índice da aba (padrão: 0 = primeira aba)"
    )
    parser.add_argument(
        "--threshold", "-t", type=float, default=1.0,
        help=(
            "Fração de valores vazios para remover a coluna. "
            "1.0 = 100%% vazio (padrão). 0.9 = 90%% vazio."
        ),
    )
    args = parser.parse_args()

    # Converte sheet para int se for número
    sheet = args.sheet
    try:
        sheet = int(sheet)
    except ValueError:
        pass

    clean_empty_columns(
        input_path=args.input,
        output_path=args.output,
        sheet_name=sheet,
        threshold=args.threshold,
    )


if __name__ == "__main__":
    main()
