from .config_loader import Config
from .data_model import ProductRecord
from .transformation import apply_transformations
from .excel_reader import _normalize_header


def _resolve_column_indices(
    headers: list[str],
    column_mapping: dict[str, str],
) -> dict[int, str]:
    """Resolve header-name-based column_mapping to index-based mapping.

    For each (header_name, field_name) in column_mapping, find the matching
    column index in the actual headers by normalized name comparison.
    """
    result: dict[int, str] = {}
    unmatched: list[str] = []

    for header_name, field_name in column_mapping.items():
        target = _normalize_header(header_name)
        found = False
        for idx, actual_header in enumerate(headers):
            if _normalize_header(actual_header) == target:
                result[idx] = field_name
                found = True
                break
        if not found:
            unmatched.append(header_name)

    if unmatched:
        print(f"警告: 以下表头名称在 Excel 中未找到匹配列: {', '.join(repr(h) for h in unmatched)}")
        print(f"  可用表头: {headers}")

    return result


def map_rows_to_dicts(
    raw_rows: list[list],
    config: Config,
    headers: list[str] | None = None,
) -> list[dict[str, object]]:
    """Map raw Excel rows to field-name-keyed dicts using column_mapping.

    If headers are provided, resolves header-name-based mappings to column
    indices first. Falls back to index-based mapping for backward compatibility.
    """
    # Resolve header names → column indices if using name-based mapping
    if headers:
        index_mapping = _resolve_column_indices(headers, config.column_mapping)
    else:
        # Legacy index-based mapping (backward compatibility)
        index_mapping = config.column_mapping  # type: ignore[assignment]

    result = []
    for row in raw_rows:
        row_dict: dict[str, object] = {}
        for col_idx, field_name in index_mapping.items():
            if col_idx < len(row):
                row_dict[field_name] = row[col_idx]
            else:
                row_dict[field_name] = ""
        # Apply transformations
        row_dict = apply_transformations(row_dict, config.transformations)
        result.append(row_dict)
    return result


def map_to_products(
    raw_rows: list[list],
    config: Config,
    headers: list[str] | None = None,
) -> list[ProductRecord]:
    """Map raw Excel rows to ProductRecord instances."""
    dicts = map_rows_to_dicts(raw_rows, config, headers)
    products = []
    for d in dicts:
        try:
            product = ProductRecord(
                seq_no=int(d.get("seq_no", 0)),
                contract_no=str(d.get("contract_no", "")),
                customer_name=str(d.get("customer_name", "")),
                yuhu_entity=str(d.get("yuhu_entity", "")),
                supplier_name=str(d.get("supplier_name", "")),
                product_name=str(d.get("product_name", "")),
                country=str(d.get("country", "")),
                factory_no=str(d.get("factory_no", "")),
                container_no=str(d.get("container_no", "")),
                spec=str(d.get("spec", "")),
                quantity=int(d.get("quantity", 0)),
                weight=float(d.get("weight", 0.0)),
                unit_price=float(d.get("unit_price", 0.0)),
                amount=float(d.get("amount", 0.0)),
                production_date=str(d.get("production_date", "")),
                warehouse=str(d.get("warehouse", "")),
            )
            products.append(product)
        except (ValueError, TypeError) as e:
            # Skip rows with invalid data
            print(f"警告: 跳过一行数据 - {e}")
            continue
    return products
