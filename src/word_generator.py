import copy
from dataclasses import fields
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from .config_loader import Config, OutputConfig, WordConfig
from .data_model import ContractData, ProductRecord
from .company_lookup import lookup_company

# 属于"金额"的字段：开启千分位时只格式化这些，重量/件数等保持原样
_MONEY_FIELDS = {
    "total_amount", "total_sales_amount", "prepayment", "premium_amount",
    "ex_tax_amount", "ex_tax_sales_amount", "vat", "sales_vat",
}


def _format_money(value: float, thousands: bool = False) -> str:
    return f"{value:,.2f}" if thousands else f"{value:.2f}"

# 汇总行占位符 → 合同字段。按占位符名填充，与模板里的行顺序无关
# （标准版是 合计/大写/税点，溢价版是 合计/税点/大写）
_SUMMARY_MAPPINGS = {
    "purchase": {
        "{{总件数}}": "total_quantity",
        "{{总重量}}": "total_weight",
        "{{总金额}}": "total_amount",
        "{{总金额大写}}": "total_amount_upper",
        "{{税点}}": "vat_rate_pct",
        "{{不含税金额}}": "ex_tax_amount",
        "{{增值税}}": "vat",
    },
    "sales": {
        "{{总件数}}": "total_quantity",
        "{{总重量}}": "total_weight",
        "{{总金额}}": "total_sales_amount",
        "{{总金额大写}}": "total_sales_amount_upper",
        "{{税点}}": "vat_rate_pct",
        "{{不含税金额}}": "ex_tax_sales_amount",
        "{{增值税}}": "sales_vat",
        "{{溢价金额}}": "premium_amount",
    },
}


def _replace_in_paragraph(paragraph, mapping: dict[str, str]):
    """Replace placeholders in a paragraph, handling runs split by Word."""
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return

    # Only touch paragraphs that actually contain placeholders,
    # otherwise preserve the original run structure and formatting.
    if not any(ph in full_text for ph in mapping):
        return

    for placeholder, value in mapping.items():
        full_text = full_text.replace(placeholder, value)

    if paragraph.runs:
        for run in paragraph.runs[1:]:
            run.text = ""
        paragraph.runs[0].text = full_text


def _replace_in_cell(cell, mapping: dict[str, str]):
    """Replace placeholders in all paragraphs of a table cell."""
    for paragraph in cell.paragraphs:
        _replace_in_paragraph(paragraph, mapping)


def _fill_product_cell_text(cell, text: str):
    """Set the text of a table cell, clearing existing runs."""
    for paragraph in cell.paragraphs:
        if paragraph.runs:
            for run in paragraph.runs[1:]:
                run.text = ""
            paragraph.runs[0].text = text
            return
        else:
            paragraph.add_run(text)
            return


# 产品行占位符 → 产品字段的取值方式（按名填充，与模板列顺序解耦）
def _product_row_values(product: ProductRecord, seq: int, contract_type: str,
                        unit_price_markup: float, thousands: bool = False) -> dict:
    unit_price = product.unit_price + unit_price_markup
    if contract_type == "sales":
        amount = unit_price * product.weight
    else:
        amount = product.amount
    def money(v):
        return f"{v:,.2f}" if thousands else f"{v:.2f}"
    return {
        "{{合同产品名称}}": product.product_name,
        "{{柜号}}": product.container_no,
        "{{件数}}": str(product.quantity),
        "{{重量}}": f"{product.weight:.3f}" if product.weight else "0.000",
        "{{采购单价}}": money(unit_price),
        "{{金额}}": money(amount),
        "{{规格}}": product.spec,
        "{{生产日期}}": product.production_date,
        "{{厂号}}": product.factory_no,
    }


def _fill_product_row(row, product: ProductRecord, seq: int,
                      contract_type: str = "purchase", unit_price_markup: float = 0.0,
                      thousands: bool = False):
    """按占位符名填充产品行（模板列顺序可不同），序号写在第一列。"""
    values = _product_row_values(product, seq, contract_type, unit_price_markup, thousands)
    if row.cells:
        _fill_product_cell_text(row.cells[0], str(seq))
    for cell in row.cells[1:]:
        _replace_in_cell(cell, values)


def _fill_summary_row_by_element(tr_element, mapping: dict[str, str]):
    """Replace placeholders in a summary row using raw XML.

    Handles horizontally-merged cells: only the first cell of each merge
    group is filled; subsequent "ghost" cells (same gridSpan as previous)
    are cleared to avoid duplicate text across cells.
    """
    import re
    tc_elements = tr_element.findall(qn('w:tc'))

    prev_span = None

    for tc in tc_elements:
        # Detect ghost cells from horizontal cell merging:
        # when consecutive cells share the same gridSpan > 1, only the
        # first one is the real cell; the rest are merge continuations.
        grid_span = 1
        tc_pr = tc.find(qn('w:tcPr'))
        if tc_pr is not None:
            gs = tc_pr.find(qn('w:gridSpan'))
            if gs is not None:
                grid_span = int(gs.get(qn('w:val'), '1'))

        is_ghost = (grid_span > 1 and grid_span == prev_span)
        prev_span = grid_span

        if is_ghost:
            # Clear all text in this ghost cell
            p_elements = tc.findall(qn('w:p'))
            for p in p_elements:
                r_elements = p.findall(qn('w:r'))
                for r in r_elements[1:]:
                    p.remove(r)
                if r_elements:
                    t_elements = r_elements[0].findall(qn('w:t'))
                    for t in t_elements[1:]:
                        r_elements[0].remove(t)
                    if t_elements:
                        t_elements[0].text = ''
                        t_elements[0].set(qn('xml:space'), 'preserve')
            continue

        # Normal processing for the real cell
        p_elements = tc.findall(qn('w:p'))
        for p in p_elements:
            r_elements = p.findall(qn('w:r'))
            full_text = ""
            for r in r_elements:
                t_elements = r.findall(qn('w:t'))
                for t in t_elements:
                    full_text += (t.text or "")

            if not full_text:
                continue

            # Normalize whitespace inside {{ }} for robust matching
            full_text = re.sub(r'\{\{\s*(\S+?)\s*\}\}', r'{{\1}}', full_text)

            for placeholder, value in mapping.items():
                full_text = full_text.replace(placeholder, value)

            for r in r_elements[1:]:
                p.remove(r)
            if r_elements:
                t_elements = r_elements[0].findall(qn('w:t'))
                for t in t_elements[1:]:
                    r_elements[0].remove(t)
                if t_elements:
                    t_elements[0].text = full_text
                    t_elements[0].set(qn('xml:space'), 'preserve')


def _resolve_path(template_path: str) -> str:
    """Resolve a template path that may be relative (PyInstaller-aware)."""
    import sys, os
    if os.path.isabs(template_path):
        return template_path
    base = getattr(sys, '_MEIPASS', os.getcwd())
    candidate = os.path.join(base, template_path)
    if os.path.exists(candidate):
        return candidate
    return os.path.join(os.getcwd(), template_path)


def _do_company_lookup(contract: ContractData, para_mapping: dict[str, str]):
    """Look up company info from 公司数据.xlsx and update para_mapping.

    Purchase contract: looks up supplier (account / bank / address) and
                       yuhu_entity (address).
    Sales contract:     looks up yuhu_entity (account / bank).
    """
    is_sales = contract.contract_type == "sales"

    try:
        if is_sales:
            # === Sales contract: look up yuhu_entity for account & bank ===
            _lookup_one_company(
                contract.yuhu_entity, "玉湖主体", para_mapping,
                account_key="{{玉湖主体账号}}",
                bank_key="{{玉湖主体开户行}}",
            )
        else:
            # === Purchase contract: look up supplier for account, bank & address ===
            _lookup_one_company(
                contract.supplier_name, "甲方/供应商", para_mapping,
                account_key="{{甲方收款账号}}",
                bank_key="{{甲方开户行}}",
                address_key="{{供应商地址}}",
            )
            # === Purchase contract: also look up yuhu_entity for address ===
            _lookup_one_company(
                contract.yuhu_entity, "玉湖主体", para_mapping,
                address_key="{{玉湖主体地址}}",
            )
    except Exception as e:
        print(f"  警告: 查找公司账户信息失败: {e}")
        for ph in ("{{玉湖主体账号}}", "{{玉湖主体开户行}}",
                    "{{甲方收款账号}}", "{{甲方开户行}}",
                    "{{供应商地址}}", "{{玉湖主体地址}}"):
            para_mapping.setdefault(ph, "")


def _lookup_one_company(company_name: str, role_label: str,
                         para_mapping: dict[str, str],
                         account_key: str | None = None,
                         bank_key: str | None = None,
                         address_key: str | None = None):
    """Look up a single company and fill the specified placeholder keys."""
    print(f"正在从 公司数据.xlsx 查找 {role_label}: '{company_name}'")
    info = lookup_company(company_name)

    if info is None:
        if account_key:
            para_mapping[account_key] = ""
        if bank_key:
            para_mapping[bank_key] = ""
        if address_key:
            para_mapping[address_key] = ""
        print(f"  警告: 未找到 {role_label} '{company_name}'，相关字段留空")
        return

    # Fill requested fields
    if account_key:
        para_mapping[account_key] = info["account"]
    if bank_key:
        para_mapping[bank_key] = info["bank"]
    if address_key:
        para_mapping[address_key] = info["location"]

    # Build status message
    parts = [f"已找到 {role_label} '{company_name}'"]
    warnings = []

    if account_key is not None:  # account lookup was requested
        if info["account"]:
            parts.append(f"账号: {info['account']}")
        else:
            warnings.append("账号为空")
    if bank_key is not None:
        if info["bank"]:
            parts.append(f"开户行: {info['bank']}")
        else:
            warnings.append("开户行为空")
    if address_key is not None:
        if info["location"]:
            parts.append(f"地址: {info['location']}")
        else:
            warnings.append("地址为空")

    if warnings:
        print(f"  {'; '.join(parts)}")
        print(f"  警告: {'，'.join(warnings)}，请在 公司数据.xlsx 中补充")
    else:
        print(f"  {'; '.join(parts)}")


_FILENAME_ILLEGAL_CHARS = '\\/:*?"<>|'


def _default_output_filename(contract: ContractData) -> str:
    """pattern 未配置或渲染失败时的兜底命名，与 config.yaml 的默认 pattern 风格一致。"""
    if contract.contract_type == "sales":
        return f"销售订单-{contract.customer_name}-{contract.serial or 'S'}.docx"
    return f"采购合同-{contract.supplier_name}-{contract.serial or 'P'}.docx"


def _build_output_filename(pattern: str, contract: ContractData) -> str:
    """Render the filename from output.pattern, e.g. 合同_{contract_no}_{supplier_name}.docx."""
    if pattern:
        values = {
            f.name: str(getattr(contract, f.name))
            for f in fields(contract)
            if f.name != "products"
        }
        try:
            rendered = pattern.format(**values)
        except (KeyError, IndexError, ValueError) as e:
            print(f"  警告: output.pattern 无法渲染 ({e})，改用默认文件名")
        else:
            for ch in _FILENAME_ILLEGAL_CHARS:
                rendered = rendered.replace(ch, "_")
            rendered = rendered.strip()
            if rendered:
                return rendered
    return _default_output_filename(contract)


def generate_contract(excel_path: str, config: Config, contract: ContractData) -> str:
    """Generate a Word contract document from the template and contract data.

    Returns the path to the generated file.
    """
    import sys, os

    # Select type-specific config
    thousands = False
    if contract.contract_type == "sales" and contract.is_premium and config.premium_sales:
        prem = config.premium_sales
        word_cfg = WordConfig(
            template_path=prem.template_path,
            placeholders=prem.placeholders,
            product_table_index=prem.product_table_index,
            product_template_row=prem.product_template_row,
            signature_table_index=prem.signature_table_index,
        )
        output_cfg = OutputConfig(
            pattern=f"{prem.output_prefix}-{{customer_name}}-{{serial}}.docx",
            directory=config.sales.output.directory if config.sales else config.output.directory,
        )
        cleanup_placeholders = config.sales.cleanup_placeholders if config.sales else []
        unit_price_markup = config.sales.unit_price_markup if config.sales else 0.0
        thousands = prem.use_thousands_separator
    elif contract.contract_type == "sales" and config.sales:
        word_cfg = config.sales.word
        output_cfg = config.sales.output
        cleanup_placeholders = config.sales.cleanup_placeholders
        unit_price_markup = config.sales.unit_price_markup
    else:
        word_cfg = config.word
        output_cfg = config.output
        cleanup_placeholders = []
        unit_price_markup = 0.0

    doc = Document(_resolve_path(word_cfg.template_path))

    # 1. Build placeholder mapping for paragraphs and non-product cells
    para_mapping = {}
    for field_name, placeholder in word_cfg.placeholders.items():
        value = getattr(contract, field_name, "")
        if isinstance(value, float) and field_name in _MONEY_FIELDS:
            para_mapping[placeholder] = _format_money(value, thousands) if value else ""
        else:
            para_mapping[placeholder] = str(value) if value else ""

    # 1b. Look up company account info from 公司数据.xlsx
    _do_company_lookup(contract, para_mapping)

    # 2. Replace in all paragraphs
    for paragraph in doc.paragraphs:
        _replace_in_paragraph(paragraph, para_mapping)

    # 3. Process product table
    if word_cfg.product_table_index < len(doc.tables):
        product_table = doc.tables[word_cfg.product_table_index]
        n_products = len(contract.products)
        if n_products > 0:
            _build_product_table(product_table, contract, n_products, unit_price_markup,
                                 word_cfg.product_template_row, thousands)

    # 4. Process signature table
    if word_cfg.signature_table_index < len(doc.tables):
        sig_table = doc.tables[word_cfg.signature_table_index]
        for row in sig_table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    _replace_in_paragraph(paragraph, para_mapping)

    # 5. Cleanup placeholders — replace with empty string
    for placeholder in cleanup_placeholders:
        for paragraph in doc.paragraphs:
            _replace_in_paragraph(paragraph, {placeholder: ""})

    # 6. Build filename from output.pattern (fallback: 采购合同/销售合同 + 合同号后三位)
    output_filename = _build_output_filename(output_cfg.pattern, contract)

    output_dir = output_cfg.directory
    if not os.path.isabs(output_dir):
        output_dir = os.path.join(os.getcwd(), output_dir)
    os.makedirs(output_dir, exist_ok=True)

    output_path = os.path.join(output_dir, output_filename)
    doc.save(output_path)
    return output_path


def _build_product_table(table, contract: ContractData, n_products: int,
                         unit_price_markup: float = 0.0, template_row_idx: int = 1,
                         thousands: bool = False):
    """Build the product table: clone rows for N products, fill summary rows."""
    template_tr = table.rows[template_row_idx]._tr
    parent = template_tr.getparent()

    summary_trs = []
    for r in range(template_row_idx + 1, len(table.rows)):
        summary_trs.append(table.rows[r]._tr)

    for tr in summary_trs:
        parent.remove(tr)

    # 先克隆出全部产品行（必须在填充之前，否则克隆到的是已填好的行）
    for _ in range(1, n_products):
        new_tr = copy.deepcopy(template_tr)
        parent.findall(qn('w:tr'))[-1].addnext(new_tr)

    # 再逐行填充
    for i in range(n_products):
        _fill_product_row(table.rows[template_row_idx + i], contract.products[i], i + 1,
                          contract.contract_type, unit_price_markup, thousands)

    all_trs = parent.findall(qn('w:tr'))
    last_tr = all_trs[-1]
    for tr in summary_trs:
        last_tr.addnext(tr)
        last_tr = tr

    _fill_summary_rows(parent, contract, template_row_idx, n_products, thousands)


def _fill_summary_rows(parent, contract: ContractData, template_row_idx: int, n_products: int,
                       thousands: bool = False):
    """Fill every summary row by placeholder name (row order does not matter)."""
    all_trs = parent.findall(qn('w:tr'))
    summary_start = template_row_idx + n_products

    mapping = _SUMMARY_MAPPINGS.get(contract.contract_type, _SUMMARY_MAPPINGS["purchase"])
    replacement = {}
    for placeholder, attr_name in mapping.items():
        val = getattr(contract, attr_name, "")
        if isinstance(val, float):
            # 只对金额字段用千分位；总重量等保持 "27,505.62" -> "27505.62"
            use_sep = thousands and attr_name in _MONEY_FIELDS
            replacement[placeholder] = _format_money(val, use_sep)
        else:
            replacement[placeholder] = str(val)

    # 每一行都套完整映射：行里没有的占位符自然不会被替换
    for row_idx in range(summary_start, len(all_trs)):
        _fill_summary_row_by_element(all_trs[row_idx], replacement)
