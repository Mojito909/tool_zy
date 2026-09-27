from pathlib import Path

from docx import Document

from src.config_loader import load_config
from src.data_model import ContractData, ProductRecord
from src.word_generator import generate_contract

PROJECT_ROOT = Path(__file__).resolve().parents[1]


def _products(n: int) -> list[ProductRecord]:
    return [
        ProductRecord(
            seq_no=i + 1,
            contract_no="ABCDEFGP01",
            customer_name="客户A",
            yuhu_entity="玉湖主体",
            supplier_name="供应商B",
            product_name=f"产品{i + 1}",
            container_no=f"C{i + 1}",
            spec="20kg/箱",
            quantity=10 * (i + 1),
            weight=100.0 * (i + 1),
            unit_price=10.0,
            amount=1000.0 * (i + 1),
            production_date="2026-01-0%d" % (i + 1),
            factory_no=f"SIF{i + 1}",
        )
        for i in range(n)
    ]


def _build_contract(products, contract_type="purchase"):
    return ContractData.from_products(
        products, vat_rate=0.09, contract_type=contract_type,
        unit_price_markup=0.1, prepayment_rate=0.3,
    )


def _make_template(path: Path, template_row_index: int):
    """Table 0: 表头 + 备用行 + 产品模板行 + 3 行汇总；Table 1: 空签署表。"""
    doc = Document()
    table = doc.add_table(rows=template_row_index + 4, cols=10)
    for cell in table.rows[0].cells:
        cell.text = "表头"
    for cell in table.rows[1].cells:
        cell.text = "备用行"
    for cell in table.rows[template_row_index].cells:
        cell.text = "模板"
    table.rows[template_row_index + 1].cells[0].text = "{{总件数}} {{总重量}} {{总金额}}"
    table.rows[template_row_index + 2].cells[0].text = "{{总金额大写}}"
    table.rows[template_row_index + 3].cells[0].text = "{{税点}} {{不含税金额}} {{增值税}}"
    doc.add_table(rows=2, cols=2)
    doc.save(str(path))


def test_generates_real_templates_with_correct_row_layout(tmp_path):
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.output.directory = str(tmp_path)
    config.sales.output.directory = str(tmp_path)
    config.data.stop_on_empty_column = "合同产品名称"
    products = _products(4)
    excel_path = str(tmp_path / "dummy.xls")

    for contract_type, template in (
        ("purchase", "采购合同模板.docx"),
        ("sales", "销售合同模板.docx"),
    ):
        doc = Document(str(PROJECT_ROOT / "templates" / template))
        assert len(doc.tables[0].rows) == 5  # 表头 + 模板行 + 3 行汇总

    sales = _build_contract(products, "sales")
    sales.total_sales_amount_upper = "大写占位"
    path = generate_contract(excel_path, config, sales)
    table = Document(path).tables[0]

    # 表头 + 4 个产品行 + 合计 + 大写 + 备注
    assert len(table.rows) == 8
    assert [table.rows[r].cells[0].text for r in range(1, 5)] == ["1", "2", "3", "4"]
    assert table.rows[1].cells[1].text == "产品1"
    assert table.rows[4].cells[6].text == "4040.00"  # (10 + 0.1) * 400
    assert "{{" not in table.rows[5].cells[0].text
    assert table.rows[5].cells[3].text == "100"
    assert table.rows[5].cells[4].text == "1000.00"
    assert table.rows[5].cells[6].text == "10100.00"  # 10.1 * 1000kg
    assert table.rows[6].cells[2].text == "大写占位"
    assert "9%" in table.rows[7].cells[2].text


def test_uses_per_contract_template_row_index(tmp_path):
    """采购和销售各自读取自己配置里的 template_row_index。"""
    template_row_index = 2
    template_path = tmp_path / "template.docx"
    _make_template(template_path, template_row_index)

    config = load_config(PROJECT_ROOT / "config.yaml")
    config.output.directory = str(tmp_path)
    config.sales.output.directory = str(tmp_path)

    # 采购：产品模板行在第 1 行（备用行）
    config.word.template_path = str(template_path)
    config.word.product_table_index = 0
    config.word.product_template_row = 1

    # 销售：产品模板行在第 2 行
    config.sales.word.template_path = str(template_path)
    config.sales.word.product_table_index = 0
    config.sales.word.product_template_row = template_row_index

    products = _products(2)

    purchase_path = generate_contract(
        str(tmp_path / "dummy.xls"), config, _build_contract(products, "purchase")
    )
    purchase_rows = Document(purchase_path).tables[0].rows
    assert purchase_rows[1].cells[0].text == "1"
    assert purchase_rows[2].cells[0].text == "2"

    sales = _build_contract(products, "sales")
    sales.total_sales_amount_upper = "大写占位"
    sales_path = generate_contract(str(tmp_path / "dummy.xls"), config, sales)
    table = Document(sales_path).tables[0]
    # 产品行从第 2 行开始，汇总 3 行紧随其后
    assert table.rows[0].cells[0].text == "表头"
    assert table.rows[1].cells[0].text == "备用行"
    assert table.rows[2].cells[0].text == "1"
    assert table.rows[3].cells[0].text == "2"
    assert table.rows[4].cells[0].text == "30 300.00 3030.00"  # 总件数 10+20
    assert "{{总件数}}" not in table.rows[4].cells[0].text


def _write_both(config, products, tmp_path):
    """生成采购 + 销售两份合同，返回 (采购路径, 销售路径)。"""
    config.output.directory = str(tmp_path)
    config.sales.output.directory = str(tmp_path)
    excel_path = str(tmp_path / "dummy.xls")
    purchase_path = generate_contract(
        excel_path, config, _build_contract(products, "purchase")
    )
    sales = _build_contract(products, "sales")
    sales.total_sales_amount_upper = "大写占位"
    sales_path = generate_contract(excel_path, config, sales)
    return purchase_path, sales_path


def test_output_pattern_renders_filename(tmp_path):
    """output.pattern 生效：用合同字段填充文件名。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.output.pattern = "合同_{contract_no}_{supplier_name}.docx"
    config.sales.output.pattern = "销售合同_{sales_contract_no}_{customer_name}.docx"

    purchase_path, sales_path = _write_both(config, _products(2), tmp_path)

    assert Path(purchase_path).name == "合同_ABCDEFGP01_供应商B.docx"
    assert Path(sales_path).name == "销售合同_ABCDEFGS01_客户A.docx"
    assert Path(purchase_path).exists() and Path(sales_path).exists()


def test_output_pattern_illegal_chars_replaced(tmp_path):
    """文件名中的非法字符替换为下划线，避免保存失败。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.output.pattern = "合同_{customer_name}.docx"
    products = _products(1)
    products[0].customer_name = "A/B:C?D"

    purchase_path, _ = _write_both(config, products, tmp_path)

    assert Path(purchase_path).name == "合同_A_B_C_D.docx"
    assert Path(purchase_path).exists()


def test_output_pattern_bad_field_falls_back_to_default(tmp_path, capsys):
    """pattern 引用了不存在的字段时回退到默认命名，不崩溃。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.output.pattern = "合同_{不存在的字段}.docx"

    purchase_path, _ = _write_both(config, _products(2), tmp_path)

    assert Path(purchase_path).name == "采购合同-供应商B-P01.docx"
    assert Path(purchase_path).exists()
    assert "无法渲染" in capsys.readouterr().out


def test_output_pattern_empty_falls_back_to_default(tmp_path):
    """pattern 为空时保持原有硬编码命名。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.output.pattern = ""
    config.sales.output.pattern = ""

    purchase_path, sales_path = _write_both(config, _products(2), tmp_path)

    assert Path(purchase_path).name == "采购合同-供应商B-P01.docx"
    assert Path(sales_path).name == "销售订单-客户A-S01.docx"


def test_financing_days_placeholder_filled(tmp_path):
    """销售的 {{融资时长}} 用 contract.financing_days 填充，不再残留占位符。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.output.directory = str(tmp_path)
    config.sales.output.directory = str(tmp_path)

    sales = _build_contract(_products(2), "sales")
    sales.total_sales_amount_upper = "大写占位"
    sales.financing_days = 120

    path = generate_contract(str(tmp_path / "dummy.xls"), config, sales)
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                text += "\n" + c.text

    assert "【120】" in text
    assert "{{" not in text


def test_financing_days_defaults_from_config(tmp_path):
    """未设置 financing_days 时用 config.sales.financing_days 兜底。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.output.directory = str(tmp_path)
    config.sales.output.directory = str(tmp_path)
    assert config.sales.financing_days == 90

    sales = _build_contract(_products(1), "sales")
    sales.total_sales_amount_upper = "大写占位"
    sales.financing_days = config.sales.financing_days

    path = generate_contract(str(tmp_path / "dummy.xls"), config, sales)
    assert "【90】" in "\n".join(p.text for p in Document(path).paragraphs)


def _make_premium_contract(products, prepayment_rate=0.20, days=90, rate=0.125):
    contract = ContractData.from_products(
        products, vat_rate=0.09, contract_type="sales",
        unit_price_markup=0.0, prepayment_rate=prepayment_rate,
    )
    contract.total_sales_amount_upper = "大写占位"
    contract.financing_days = days
    contract.financing_rate = rate
    contract.is_premium = True
    contract.premium_amount = contract.compute_premium_amount()
    return contract


def test_premium_template_fills_by_placeholder_name(tmp_path):
    """溢价版列序与旧版不同：规格在第 4 列、厂号第 5 列，必须按占位符名填充。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.sales.output.directory = str(tmp_path)
    products = _products(2)

    path = generate_contract(
        str(tmp_path / "dummy.xls"), config, _make_premium_contract(products)
    )
    table = Document(path).tables[0]

    # 表头 2 行 + 产品 2 行 + 合计 + 税点 + 大写
    assert len(table.rows) == 7
    assert [table.rows[r].cells[0].text for r in (2, 3)] == ["1", "2"]
    assert table.rows[2].cells[1].text == "产品1"      # 货物品名
    assert table.rows[2].cells[3].text == "20kg/箱"     # 规格
    assert table.rows[2].cells[4].text == "SIF1"        # 厂号
    assert table.rows[2].cells[5].text == "10"          # 件数
    assert table.rows[3].cells[1].text == "产品2"
    assert table.rows[3].cells[4].text == "SIF2"


def test_premium_amount_and_thousands_separator_in_document(tmp_path):
    """溢价金额写入文档；金额带千分位，重量不带。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.sales.output.directory = str(tmp_path)
    # 采购金额（决定垫款与溢价）与重量按样例设置；
    # 销售行金额 = (单价 + 加价) × 重量
    products = _products(1)
    products[0].amount = 1567820.35
    products[0].weight = 27505.62
    products[0].quantity = 1171

    contract = _make_premium_contract(products)
    assert contract.prepayment == 313564.07
    assert contract.premium_amount == 39195.51

    path = generate_contract(str(tmp_path / "dummy.xls"), config, contract)
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs)
    table_text = "\n".join(c.text for r in doc.tables[0].rows for c in r.cells)

    assert "溢价金额【39,195.51】" in text
    assert "{{" not in text and "{{" not in table_text
    assert "277,806.76" in table_text        # 金额：带千分位（10.10 × 27505.62）
    assert "27505.620" in table_text         # 重量：不带千分位


def test_premium_uses_separate_output_filename(tmp_path):
    """溢价版文件名带独立前缀，不会覆盖标准销售订单。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.sales.output.directory = str(tmp_path)
    products = _products(1)

    standard = generate_contract(
        str(tmp_path / "d.xls"), config, _build_contract(products, "sales")
    )
    premium = generate_contract(
        str(tmp_path / "d.xls"), config, _make_premium_contract(products)
    )

    assert Path(standard).name != Path(premium).name
    assert Path(premium).name.startswith("销售订单溢价版-")
    assert Path(standard).exists() and Path(premium).exists()


def test_standard_sales_has_no_premium_leftover(tmp_path):
    """标准销售模板没有 {{溢价金额}}，选旧版不受影响。"""
    config = load_config(PROJECT_ROOT / "config.yaml")
    config.sales.output.directory = str(tmp_path)

    path = generate_contract(
        str(tmp_path / "d.xls"), config, _build_contract(_products(2), "sales")
    )
    text = "\n".join(p.text for p in Document(path).paragraphs)
    assert "{{" not in text
    assert "溢价" not in text
