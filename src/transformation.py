import math
from datetime import datetime, timedelta
from .config_loader import TransformRule

EXCEL_EPOCH = datetime(1899, 12, 31)


def transform_excel_date(serial: float) -> str | None:
    """Convert an Excel serial date number to a date string.

    Excel serial 45944.0 -> datetime(2025, 10, 15)
    Handles the Lotus 1-2-3 bug (Feb 29, 1900 is not a real date).
    Returns None if the value is not a valid date serial.
    """
    if not serial or serial <= 0:
        return None
    try:
        serial_int = int(serial)
        if serial_int <= 60:
            delta = serial_int
        else:
            # Skip the phantom 1900-02-29 (Lotus 1-2-3 bug: Excel day 60)
            delta = serial_int - 1
        dt = EXCEL_EPOCH + timedelta(days=delta)
        return dt.strftime("%Y-%m-%d")
    except (ValueError, OverflowError):
        return None


def transform_number(value, decimal_places: int = 2, default=0.0):
    """Convert a value to a number with specified decimal places."""
    if value is None or value == "":
        return default
    try:
        v = float(value)
        if math.isnan(v) or math.isinf(v):
            return default
        if decimal_places == 0:
            return int(round(v))
        return round(v, decimal_places)
    except (ValueError, TypeError):
        return default


def transform_to_string(value) -> str:
    """Convert a float-like value to a string, stripping .0 for whole numbers."""
    if value is None or value == "":
        return ""
    try:
        v = float(value)
        if v == int(v):
            return str(int(v))
        return str(v)
    except (ValueError, TypeError):
        return str(value)


def apply_transformation(field_name: str, raw_value, rule: TransformRule) -> object:
    """Apply a transformation rule to a raw field value."""
    if raw_value is None or raw_value == "":
        return rule.default

    if rule.type == "excel_date":
        if isinstance(raw_value, (int, float)) and raw_value > 0:
            result = transform_excel_date(float(raw_value))
            return result if result else rule.default
        # If already a datetime or date string
        if isinstance(raw_value, datetime):
            return raw_value.strftime(rule.format)
        return str(raw_value)

    elif rule.type == "to_string":
        return transform_to_string(raw_value)

    elif rule.type == "number":
        return transform_number(raw_value, rule.decimal_places, rule.default)

    else:
        return str(raw_value)


def apply_transformations(
    row_dict: dict[str, object],
    transform_rules: dict[str, TransformRule],
) -> dict[str, object]:
    """Apply all configured transformations to a row dict."""
    result = dict(row_dict)
    for field_name, rule in transform_rules.items():
        if field_name in result:
            result[field_name] = apply_transformation(field_name, result[field_name], rule)
    return result
