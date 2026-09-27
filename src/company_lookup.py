"""Look up company bank account info from the company data Excel file."""
import os
import re
import sys
from pathlib import Path


class CompanyLookupError(Exception):
    pass


def _normalize_name(name: str) -> str:
    """Normalize a company name for fuzzy matching.

    Converts full-width characters (Chinese punctuation, letters, digits) to
    half-width equivalents so that names like '玉湖冷链食品(广州)有限公司' and
    '玉湖冷链食品（广州）有限公司' match correctly.
    """
    if not name:
        return ""

    result = []
    for ch in name:
        code = ord(ch)
        # Full-width digits ０-９ (U+FF10-U+FF19) → half-width 0-9
        if 0xFF10 <= code <= 0xFF19:
            result.append(chr(code - 0xFF10 + ord('0')))
        # Full-width uppercase Ａ-Ｚ (U+FF21-U+FF3A) → A-Z
        elif 0xFF21 <= code <= 0xFF3A:
            result.append(chr(code - 0xFF21 + ord('A')))
        # Full-width lowercase ａ-ｚ (U+FF41-U+FF5A) → a-z
        elif 0xFF41 <= code <= 0xFF5A:
            result.append(chr(code - 0xFF41 + ord('a')))
        # Full-width parentheses （ ）→ ( )
        elif code == 0xFF08:  # （
            result.append('(')
        elif code == 0xFF09:  # ）
            result.append(')')
        # Full-width comma ，→ ,
        elif code == 0xFF0C:
            result.append(',')
        # Full-width period ．→ .
        elif code == 0xFF0E:
            result.append('.')
        # Full-width space → half-width space
        elif code == 0x3000:
            result.append(' ')
        else:
            result.append(ch)
    return ''.join(result)


def _find_company_data_file() -> str | None:
    """Find the company data Excel file. Searches: cwd and exe dir."""
    search_dirs = [os.getcwd()]

    # exe directory (for one-file PyInstaller, sys.executable is the exe path)
    try:
        exe_dir = os.path.dirname(sys.executable)
        if exe_dir and os.path.isdir(exe_dir):
            search_dirs.append(exe_dir)
    except Exception:
        pass

    for d in search_dirs:
        path = os.path.join(d, "公司数据.xlsx")
        if os.path.exists(path):
            return path
    return None


def lookup_company(company_name: str) -> dict[str, str] | None:
    """Look up a company by name and return its account info.

    Returns dict with keys: company_name, location, account, bank
    Returns None if company not found.
    """
    if not company_name or not company_name.strip():
        return None

    file_path = _find_company_data_file()
    if not file_path:
        raise CompanyLookupError("找不到公司数据文件: 公司数据.xlsx，请放在当前目录或 exe 同目录")

    import openpyxl
    wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active

    name_normalized = _normalize_name(company_name.strip())
    result = None
    all_names = []  # for debug output if no match

    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            break
        row_name = str(row[0]).strip()
        row_normalized = _normalize_name(row_name)
        all_names.append(row_name)

        if row_normalized == name_normalized:
            result = {
                "company_name": str(row[0]) if row[0] else "",
                "location": str(row[1]) if row[1] else "",
                "account": str(row[2]) if row[2] else "",
                "bank": str(row[3]) if row[3] else "",
            }
            break

    wb.close()

    if result is None:
        print(f"调试: 查找 '{company_name.strip()}' (标准化: '{name_normalized}')")
        print(f"调试: 文件中的公司名称列表: {all_names[:5]}{'...' if len(all_names) > 5 else ''}")

    return result
