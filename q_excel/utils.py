# q_excel/utils.py

from io import BytesIO
from pathlib import Path
from typing import Any, Dict, List

import pandas as pd


def get_file_suffix(filename_or_path: str) -> str:
    """
    Finder filtypen fra et filnavn eller en filsti.

    Eksempler:
        "test.xlsx" -> ".xlsx"
        "data.csv" -> ".csv"
    """

    return Path(filename_or_path).suffix.lower()


def bytes_to_memory_file(file_bytes: bytes) -> BytesIO:
    """
    Laver bytes om til et fil-lignende memory-objekt.

    BytesIO gør, at pandas kan læse filen,
    uden at filen først skal gemmes på disk.
    """

    if not file_bytes:
        raise ValueError("file_bytes er tom. Filen blev ikke hentet korrekt.")

    return BytesIO(file_bytes)


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """
    Rydder op i en DataFrame.

    DataFrame betyder en tabel i pandas.

    Gør:
    - tomme værdier bliver til ""
    - kolonnenavne bliver trimmet for mellemrum
    - kolonnenavne bliver tekst
    """

    df = df.fillna("")
    df.columns = [str(column).strip() for column in df.columns]

    return df


def dataframe_to_dicts(df: pd.DataFrame) -> List[Dict[str, Any]]:
    """
    Konverterer en DataFrame til en liste med dicts.

    Hver række i Excel bliver til én dict.
    """

    return df.to_dict(orient="records")


def print_first_dict_rows(
    rows: List[Dict[str, Any]],
    number_of_rows: int = 5,
) -> None:
    """
    Printer de første rækker fra en liste med dicts.
    """

    print("\n==============================")
    print(f"📋 Første {number_of_rows} rækker som dict")
    print("==============================")

    for row_number, row in enumerate(rows[:number_of_rows], start=1):
        print(f"\n--- Række {row_number} ---")
        print(row)