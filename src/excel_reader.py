import re
import csv
from pathlib import Path
from .config_loader import Config


def _normalize_header(text: str) -> str:
    """Normalize header text for matching: collapse whitespace, strip."""
    if text is None:
        return ""
    return re.sub(r"\s+", " ", str(text)).strip()


class SheetNotFoundError(Exception):
    pass


class ExcelReadError(Exception):
    pass


def _find_sheet_name_xlrd(file_path: str, config: Config) -> tuple[str, int]:
    """Find the correct sheet in a .xls file using xlrd. Returns (name, index)."""
    import xlrd
    wb = xlrd.open_workbook(file_path)
    sheet_names = wb.sheet_names()

    mode = config.sheet_detection.mode

    if mode == "by_name":
        target = config.sheet_detection.pattern
        if target in sheet_names:
            return target, sheet_names.index(target)
        raise SheetNotFoundError(f"Sheet '{target}' 不存在。可用 Sheet: {sheet_names}")

    elif mode == "by_pattern":
        pattern = re.compile(config.sheet_detection.pattern, re.IGNORECASE)
        for i, name in enumerate(sheet_names):
            if pattern.search(name):
                return name, i
        raise SheetNotFoundError(
            f"没有匹配模式 '{config.sheet_detection.pattern}' 的 Sheet。可用 Sheet: {sheet_names}"
        )

    elif mode == "by_index":
        idx = int(config.sheet_detection.pattern)
        if 0 <= idx < len(sheet_names):
            return sheet_names[idx], idx
        raise SheetNotFoundError(f"Sheet 索引 {idx} 无效。共 {len(sheet_names)} 个 Sheet")

    elif mode == "first":
        return sheet_names[0], 0

    raise SheetNotFoundError(f"未知的 sheet_detection.mode: {mode}")


def _find_sheet_name_openpyxl(file_path: str, config: Config):
    """Find the correct sheet in a .xlsx file using openpyxl. Returns (name, index)."""
    import openpyxl
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    wb.close()

    mode = config.sheet_detection.mode

    if mode == "by_name":
        target = config.sheet_detection.pattern
        if target in sheet_names:
            return target, sheet_names.index(target)
        raise SheetNotFoundError(f"Sheet '{target}' 不存在。可用 Sheet: {sheet_names}")

    elif mode == "by_pattern":
        pattern = re.compile(config.sheet_detection.pattern, re.IGNORECASE)
        for i, name in enumerate(sheet_names):
            if pattern.search(name):
                return name, i
        raise SheetNotFoundError(
            f"没有匹配模式 '{config.sheet_detection.pattern}' 的 Sheet。可用 Sheet: {sheet_names}"
        )

    elif mode == "by_index":
        idx = int(config.sheet_detection.pattern)
        if 0 <= idx < len(sheet_names):
            return sheet_names[idx], idx
        raise SheetNotFoundError(f"Sheet 索引 {idx} 无效。共 {len(sheet_names)} 个 Sheet")

    elif mode == "first":
        return sheet_names[0], 0

    raise SheetNotFoundError(f"未知的 sheet_detection.mode: {mode}")


def read_excel(file_path: str, config: Config) -> list[list]:
    """Read an Excel file and return data rows as list of lists.

    Supports .xls, .xlsx, .xlsm, .csv formats.
    Returns data rows from configured start row, stopping when the
    stop_on_empty_column is empty.
    """
    file_path = str(file_path)
    ext = Path(file_path).suffix.lower()

    if ext == ".csv":
        return _read_csv(file_path, config)
    elif ext in (".xls",):
        return _read_xls(file_path, config)
    elif ext in (".xlsx", ".xlsm"):
        return _read_xlsx(file_path, config)
    else:
        raise ExcelReadError(f"不支持的文件格式: {ext}。支持: .xls, .xlsx, .xlsm, .csv")


def read_excel_header(file_path: str, config: Config) -> list[str]:
    """Read the header row from an Excel file and return normalized header names.

    Returns a list of normalized header strings (whitespace collapsed, stripped).
    """
    file_path = str(file_path)
    ext = Path(file_path).suffix.lower()

    if ext == ".csv":
        return _read_csv_header(file_path, config)
    elif ext in (".xls",):
        return _read_xls_header(file_path, config)
    elif ext in (".xlsx", ".xlsm"):
        return _read_xlsx_header(file_path, config)
    else:
        raise ExcelReadError(f"不支持的文件格式: {ext}。支持: .xls, .xlsx, .xlsm, .csv")


def _read_csv_header(file_path: str, config: Config) -> list[str]:
    with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        all_rows = list(reader)
    header_row = config.data.header_row
    if header_row < len(all_rows):
        return [_normalize_header(c) for c in all_rows[header_row]]
    return []


def _read_xls_header(file_path: str, config: Config) -> list[str]:
    import xlrd
    wb = xlrd.open_workbook(file_path)
    _, sheet_idx = _find_sheet_name_xlrd(file_path, config)
    sheet = wb.sheet_by_index(sheet_idx)
    header_row = config.data.header_row
    if header_row < sheet.nrows:
        return [_normalize_header(sheet.cell_value(header_row, c)) for c in range(sheet.ncols)]
    return []


def _read_xlsx_header(file_path: str, config: Config) -> list[str]:
    import openpyxl
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    _, sheet_idx = _find_sheet_name_openpyxl(file_path, config)
    ws = wb[wb.sheetnames[sheet_idx]]
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()
    header_row = config.data.header_row
    if header_row < len(all_rows):
        return [_normalize_header(c) for c in all_rows[header_row]]
    return []


def _read_csv(file_path: str, config: Config) -> list[list]:
    with open(file_path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        all_rows = list(reader)

    start = config.data.data_start_row
    stop_col_name = config.data.stop_on_empty_column

    # Find column index for stop condition
    header_row = all_rows[config.data.header_row] if config.data.header_row < len(all_rows) else []
    stop_col_idx = _find_column_index(header_row, stop_col_name)

    result = []
    for row in all_rows[start:]:
        if stop_col_idx is not None and stop_col_idx < len(row):
            if not row[stop_col_idx].strip():
                break
        result.append(list(row))
    return result


def _read_xls(file_path: str, config: Config) -> list[list]:
    import xlrd

    wb = xlrd.open_workbook(file_path)
    sheet_name, sheet_idx = _find_sheet_name_xlrd(file_path, config)
    sheet = wb.sheet_by_index(sheet_idx)

    header_row = config.data.header_row
    start = config.data.data_start_row
    stop_col_name = config.data.stop_on_empty_column

    # Read header row to find stop column index
    headers = []
    if header_row < sheet.nrows:
        headers = [str(sheet.cell_value(header_row, c)) for c in range(sheet.ncols)]
    stop_col_idx = _find_column_index(headers, stop_col_name)

    result = []
    for r in range(start, sheet.nrows):
        row = [sheet.cell_value(r, c) for c in range(sheet.ncols)]

        # Stop if the designated column is empty
        if stop_col_idx is not None and stop_col_idx < len(row):
            val = row[stop_col_idx]
            if val is None or (isinstance(val, str) and val.strip() == "") or (isinstance(val, float) and val == 0.0):
                break

        result.append(row)

    return result


def _read_xlsx(file_path: str, config: Config) -> list[list]:
    import openpyxl

    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    sheet_name, sheet_idx = _find_sheet_name_openpyxl(file_path, config)
    ws = wb[sheet_name]

    header_row = config.data.header_row
    start = config.data.data_start_row
    stop_col_name = config.data.stop_on_empty_column

    # Read all rows into memory (read_only mode uses generator)
    all_rows = list(ws.iter_rows(values_only=True))
    wb.close()

    # Find stop column index from header
    headers = []
    if header_row < len(all_rows):
        headers = [str(c) if c is not None else "" for c in all_rows[header_row]]
    stop_col_idx = _find_column_index(headers, stop_col_name)

    result = []
    for row in all_rows[start:]:
        row_list = list(row)

        # Stop condition
        if stop_col_idx is not None and stop_col_idx < len(row_list):
            val = row_list[stop_col_idx]
            if val is None or (isinstance(val, str) and val.strip() == "") or (isinstance(val, (int, float)) and float(val) == 0.0):
                break

        result.append(row_list)

    return result


def _find_column_index(headers: list[str], col_name: str) -> int | None:
    """Find the index of a column by name (normalized match)."""
    if not col_name:
        return None
    target = _normalize_header(col_name)
    for i, h in enumerate(headers):
        if _normalize_header(h) == target:
            return i
    return None
