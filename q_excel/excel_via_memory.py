# q_excel/memory.py

from typing import Any, Dict, List, Union

import pandas as pd

from q_excel.utils import (
    bytes_to_memory_file,
    clean_dataframe,
    dataframe_to_dicts,
    get_file_suffix,
)


def read_file_from_memory(
    filename: str,
    file_bytes: bytes,
    has_header: bool = True,
    sheet_name: Union[int, str] = 0,
    csv_encoding: str = "utf-8",
) -> pd.DataFrame:
    """
    Læser en fil direkte fra memory.

    filename:
        Filnavn, fx "data.xlsx" eller "data.csv".

    file_bytes:
        Rå filindhold fra fx SharePoint.

    has_header:
        True betyder at første række er kolonnenavne.

    sheet_name:
        Excel-ark. 0 betyder første ark.

    csv_encoding:
        Tegnsæt til CSV, fx "utf-8".

    Returnerer:
        DataFrame.
    """

    file_suffix = get_file_suffix(filename)

    if file_suffix == ".xlsx":
        return read_excel_from_memory(
            file_bytes=file_bytes,
            has_header=has_header,
            sheet_name=sheet_name,
        )

    if file_suffix == ".csv":
        return read_csv_from_memory(
            file_bytes=file_bytes,
            has_header=has_header,
            encoding=csv_encoding,
        )

    raise ValueError(f"Unsupported file type: {file_suffix}")


def read_file_from_memory_as_dicts(
    filename: str,
    file_bytes: bytes,
    has_header: bool = True,
    sheet_name: Union[int, str] = 0,
    csv_encoding: str = "utf-8",
) -> List[Dict[str, Any]]:
    """
    Læser en fil fra memory og returnerer data som dicts.

    Dette er den funktion, din proces typisk skal kalde.

    Returnerer:
        List[Dict[str, Any]]
    """

    df = read_file_from_memory(
        filename=filename,
        file_bytes=file_bytes,
        has_header=has_header,
        sheet_name=sheet_name,
        csv_encoding=csv_encoding,
    )

    return dataframe_to_dicts(df)


def read_excel_from_memory(
    file_bytes: bytes,
    has_header: bool = True,
    sheet_name: Union[int, str] = 0,
) -> pd.DataFrame:
    """
    Læser Excel-fil direkte fra memory.

    Denne funktion opretter ingen mapper og gemmer ingen filer.
    """

    memory_file = bytes_to_memory_file(file_bytes)

    if has_header:
        df = pd.read_excel(
            memory_file,
            dtype=str,
            engine="openpyxl",
            sheet_name=sheet_name,
        )
    else:
        df = pd.read_excel(
            memory_file,
            header=None,
            dtype=str,
            engine="openpyxl",
            sheet_name=sheet_name,
        )
        df.columns = [f"Column_{i + 1}" for i in range(len(df.columns))]

    return clean_dataframe(df)


def read_csv_from_memory(
    file_bytes: bytes,
    has_header: bool = True,
    encoding: str = "utf-8",
) -> pd.DataFrame:
    """
    Læser CSV-fil direkte fra memory.

    Denne funktion opretter ingen mapper og gemmer ingen filer.
    """

    memory_file = bytes_to_memory_file(file_bytes)

    if has_header:
        df = pd.read_csv(
            memory_file,
            dtype=str,
            encoding=encoding,
        )
    else:
        df = pd.read_csv(
            memory_file,
            header=None,
            dtype=str,
            encoding=encoding,
        )
        df.columns = [f"Column_{i + 1}" for i in range(len(df.columns))]

    return clean_dataframe(df)