# q_excel/edit_excel.py

import os
from pathlib import Path
from typing import Any, Dict, List, Optional, Union

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import range_boundaries

from q_excel.utils import (
    bytes_to_memory_file,
    clean_dataframe,
    dataframe_to_dicts,
    get_file_suffix,
)


# -----------------------------
# BASE PATH (CONTAINER READY)
# -----------------------------
BASE_DIR = Path(os.getenv("APP_BASE_DIR", Path.cwd()))

DOWNLOAD_DIR = BASE_DIR / "downloads"
PROCESSED_DIR = BASE_DIR / "processed"

DOWNLOAD_DIR.mkdir(parents=True, exist_ok=True)
PROCESSED_DIR.mkdir(parents=True, exist_ok=True)


# -----------------------------
# FILE LOADER - DISK
# -----------------------------
def read_file(
    file_path: str,
    has_header: bool = True,
) -> pd.DataFrame:
    """
    Læser en fil fra disk.

    Understøtter:
    - .xlsx
    - .csv

    Returnerer:
        DataFrame
    """

    file_path_object = Path(file_path)
    file_suffix = get_file_suffix(str(file_path_object))

    if file_suffix == ".xlsx":
        return read_excel_flexible(
            file_path=file_path_object,
            has_header=has_header,
        )

    if file_suffix == ".csv":
        return read_csv_flexible(
            file_path=file_path_object,
            has_header=has_header,
        )

    raise ValueError(f"Unsupported file type: {file_suffix}")


# -----------------------------
# FILE LOADER - MEMORY
# -----------------------------
def read_file_from_memory(
    filename: str,
    file_bytes: bytes,
    has_header: bool = True,
    sheet_name: Union[int, str] = 0,
    csv_encoding: str = "utf-8",
) -> pd.DataFrame:
    """
    Læser en fil direkte fra memory.

    Bruges når processen først har hentet filen fra SharePoint,
    og ikke vil gemme filen på disk.

    Returnerer:
        DataFrame
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


# -----------------------------
# CSV HANDLER - DISK
# -----------------------------
def read_csv_flexible(
    file_path: Union[str, Path],
    has_header: bool = True,
) -> pd.DataFrame:
    """
    Læser CSV-fil fra disk.

    Returnerer:
        DataFrame
    """

    if has_header:
        df = pd.read_csv(
            file_path,
            dtype=str,
        )
    else:
        df = pd.read_csv(
            file_path,
            header=None,
            dtype=str,
        )
        df.columns = [f"Column_{i + 1}" for i in range(len(df.columns))]

    return clean_dataframe(df)


# -----------------------------
# CSV HANDLER - MEMORY
# -----------------------------
def read_csv_from_memory(
    file_bytes: bytes,
    has_header: bool = True,
    encoding: str = "utf-8",
) -> pd.DataFrame:
    """
    Læser CSV-fil fra memory.

    Returnerer:
        DataFrame
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


# -----------------------------
# EXCEL TABLE - DISK
# -----------------------------
def read_excel_flexible(
    file_path: Union[str, Path],
    has_header: bool = True,
    sheet_name: Union[int, str] = 0,
) -> pd.DataFrame:
    """
    Læser Excel-fil fra disk.

    Returnerer:
        DataFrame
    """

    if has_header:
        df = pd.read_excel(
            file_path,
            dtype=str,
            engine="openpyxl",
            sheet_name=sheet_name,
        )
    else:
        df = pd.read_excel(
            file_path,
            header=None,
            dtype=str,
            engine="openpyxl",
            sheet_name=sheet_name,
        )
        df.columns = [f"Column_{i + 1}" for i in range(len(df.columns))]

    return clean_dataframe(df)


# -----------------------------
# EXCEL TABLE - MEMORY
# -----------------------------
def read_excel_from_memory(
    file_bytes: bytes,
    has_header: bool = True,
    sheet_name: Union[int, str] = 0,
) -> pd.DataFrame:
    """
    Læser Excel-fil direkte fra memory.

    Denne funktion bruger BytesIO,
    så pandas kan læse Excel-filen uden disk.

    Returnerer:
        DataFrame
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


# -----------------------------
# EXCEL CELL HANDLER
# -----------------------------
def excel_edit(
    file_path: str,
    mode: str,
    from_ref: Optional[str] = None,
    to_ref: Optional[str] = None,
    value=None,
    sheet_name: Optional[str] = None,
    save: bool = False,
    save_path: Optional[str] = None,
):
    """
    Læser, kopierer, skriver eller sletter celler i en Excel-fil.

    Denne funktion arbejder med filsti på disk,
    fordi openpyxl gemmer ændringer i en workbook.

    Modes:
        read
        copy
        write
        clear
    """

    wb = load_workbook(file_path, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.active

    print(f"Bruger ark: '{ws.title}'")

    # ---------- READ ----------
    if mode == "read":
        if not from_ref:
            raise ValueError("read kræver from_ref")

        if ":" in from_ref:
            min_col, min_row, max_col, max_row = range_boundaries(from_ref)

            data = []
            for row in ws.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
            ):
                data.append([cell.value for cell in row])

            return data

        return ws[from_ref].value

    # ---------- COPY ----------
    if mode == "copy":
        if not from_ref or not to_ref:
            raise ValueError("copy kræver både from_ref og to_ref")

        if ":" in from_ref and ":" in to_ref:
            min_c1, min_r1, max_c1, max_r1 = range_boundaries(from_ref)
            min_c2, min_r2, max_c2, max_r2 = range_boundaries(to_ref)

            rows1 = max_r1 - min_r1 + 1
            cols1 = max_c1 - min_c1 + 1
            rows2 = max_r2 - min_r2 + 1
            cols2 = max_c2 - min_c2 + 1

            if rows1 != rows2 or cols1 != cols2:
                raise ValueError("Ranges skal have samme størrelse")

            for i, row in enumerate(
                ws.iter_rows(
                    min_row=min_r1,
                    max_row=max_r1,
                    min_col=min_c1,
                    max_col=max_c1,
                )
            ):
                for j, cell in enumerate(row):
                    ws.cell(
                        row=min_r2 + i,
                        column=min_c2 + j,
                    ).value = cell.value
        else:
            ws[to_ref] = ws[from_ref].value

    # ---------- WRITE ----------
    elif mode == "write":
        if not to_ref:
            raise ValueError("write kræver to_ref")

        if ":" in to_ref:
            min_col, min_row, max_col, max_row = range_boundaries(to_ref)

            for row in ws.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
            ):
                for cell in row:
                    cell.value = value
        else:
            ws[to_ref] = value

    # ---------- CLEAR ----------
    elif mode == "clear":
        if not to_ref:
            raise ValueError("clear kræver to_ref")

        if ":" in to_ref:
            min_col, min_row, max_col, max_row = range_boundaries(to_ref)

            for row in ws.iter_rows(
                min_row=min_row,
                max_row=max_row,
                min_col=min_col,
                max_col=max_col,
            ):
                for cell in row:
                    cell.value = None
        else:
            ws[to_ref] = None

    else:
        raise ValueError("Invalid mode. Brug read, copy, write eller clear.")

    # ---------- SAVE ----------
    if save:
        if not save_path:
            raise ValueError("save_path skal angives når save=True")

        save_path_object = Path(save_path)

        # Relative paths gemmes i processed-mappen
        if not save_path_object.is_absolute():
            save_path_object = PROCESSED_DIR / save_path_object

        save_path_object.parent.mkdir(parents=True, exist_ok=True)

        print(f"Gemmer fil til: {save_path_object}")

        wb.save(save_path_object)

    return None


# -----------------------------
# SAFE WRAPPER - DISK
# -----------------------------
def safe_edit_file(
    file_path: str,
    mode: Optional[str] = None,
    from_ref: Optional[str] = None,
    to_ref: Optional[str] = None,
    value=None,
    sheet_name: Optional[str] = None,
    save: bool = False,
    save_path: Optional[str] = None,
    has_header: bool = True,
):
    """
    Sikker wrapper til filer på disk.

    Bruges hvis filen allerede ligger i downloads-mappen.

    Hvis mode er "read":
        Returnerer DataFrame for xlsx/csv.

    Hvis mode er "copy", "write" eller "clear":
        Arbejder med Excel-celler.
    """

    file_path_object = Path(file_path)

    # Container-safe input
    if not file_path_object.is_absolute():
        file_path_object = DOWNLOAD_DIR / file_path_object

    file_suffix = get_file_suffix(str(file_path_object))

    if file_suffix == ".csv":
        df = read_csv_flexible(
            file_path=file_path_object,
            has_header=has_header,
        )

        if mode == "read":
            return df

        raise ValueError("CSV understøtter ikke cell manipulation")

    if file_suffix == ".xlsx":
        if mode == "read":
            return read_excel_flexible(
                file_path=file_path_object,
                has_header=has_header,
                sheet_name=sheet_name if sheet_name else 0,
            )

        if not mode:
            raise ValueError("mode skal angives for Excel-filer")

        return excel_edit(
            file_path=str(file_path_object),
            mode=mode,
            from_ref=from_ref,
            to_ref=to_ref,
            value=value,
            sheet_name=sheet_name,
            save=save,
            save_path=save_path,
        )

    raise ValueError(f"Unsupported file type: {file_suffix}")