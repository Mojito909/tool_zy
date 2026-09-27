from src.data_model import ContractData, ProductRecord


def _product(**overrides):
    data = dict(
        seq_no=1,
        contract_no="YH2026P001",
        customer_name="客户A",
        yuhu_entity="玉湖主体",
        supplier_name="供应商B",
        weight=100.0,
        unit_price=10.0,
        amount=1000.0,
    )
    data.update(overrides)
    return ProductRecord(**data)


def test_purchase_totals_and_vat_split():
    contract = ContractData.from_products([_product()], vat_rate=0.09, contract_type="purchase")
    assert contract.total_amount == 1000.0
    assert contract.total_weight == 100.0
    assert contract.ex_tax_amount == 917.43
    assert contract.vat == 82.57
    assert contract.vat_rate_pct == "9%"


def test_sales_total_applies_markup_to_unit_price():
    contract = ContractData.from_products(
        [_product()], vat_rate=0.09, contract_type="sales",
        unit_price_markup=0.1, prepayment_rate=0.3,
    )
    assert contract.total_sales_amount == 1010.0
    assert contract.ex_tax_sales_amount == 926.61
    assert contract.sales_vat == 83.39


def test_prepayment_is_based_on_purchase_amount():
    """预付款 = 采购合同总金额 * 比例（按采购金额推算，不是销售金额）。"""
    contract = ContractData.from_products(
        [_product()], vat_rate=0.09, contract_type="sales",
        unit_price_markup=0.1, prepayment_rate=0.3,
    )
    assert contract.total_amount == 1000.0
    assert contract.prepayment == 300.0
    assert contract.prepayment_rate_pct == "30%"


def test_purchase_contract_has_no_sales_fields():
    contract = ContractData.from_products([_product()], vat_rate=0.09, contract_type="purchase")
    assert contract.total_sales_amount == 0.0
    assert contract.prepayment == 0.0
    assert contract.sales_contract_no == ""
    assert contract.master_agreement_no == ""


def test_sales_contract_no_derivation():
    """末三位首位为 P 时替换为 S；主协议号取合同号去掉后 7 位。"""
    # 注意: 代码用了 len(cn) >= 7，不足以去掉 7 位时返回值不对
    cn_prefix = "ABC"
    changed = ContractData.from_products(
        [_product(contract_no="ABCDEFGP01")], contract_type="sales", prepayment_rate=0.3,
    )
    assert changed.sales_contract_no == "ABCDEFGS01"
    assert changed.master_agreement_no == cn_prefix

    unchanged = ContractData.from_products(
        [_product(contract_no="YH2026P001")], contract_type="sales", prepayment_rate=0.3,
    )
    assert unchanged.sales_contract_no == "YH2026P001"


def test_serial_marker_matches_contract_no_suffix():
    """serial = 编号标记 + 末 2 位：采购 P、销售 S，与编号位数无关。"""
    purchase = ContractData.from_products([_product(contract_no="YH2026P001")], contract_type="purchase")
    assert purchase.serial == "P01"

    sales = ContractData.from_products(
        [_product(contract_no="YH2026P001")], contract_type="sales", prepayment_rate=0.3,
    )
    # 注意: sales_contract_no 的 P→S 只在 P 位于倒数第 3 位时替换，
    # "YH2026P001" 的 P 在倒数第 4 位，因此保持原号（既有行为，未改动）
    assert sales.sales_contract_no == "YH2026P001"
    assert sales.serial == "S01"

    # 2 位编号 / 空编号的边界
    assert ContractData.from_products([_product(contract_no="GP7")], contract_type="purchase").serial == "P7"
    assert ContractData.from_products([_product(contract_no="")], contract_type="purchase").serial == "P"
    assert ContractData.from_products([_product(contract_no="")], contract_type="sales").serial == "S"


def test_premium_amount_matches_sample_document():
    """用样例合同真实数字校验公式：1567820.35 × 80% × 12.5% ÷ 360 × 90 = 39195.51。"""
    contract = ContractData.from_products(
        [_product(amount=1567820.35)], contract_type="sales", prepayment_rate=0.20,
    )
    contract.financing_days = 90
    contract.financing_rate = 0.125
    assert contract.total_amount == 1567820.35
    assert contract.prepayment == 313564.07
    assert contract.compute_premium_amount() == 39195.51


def test_premium_amount_is_zero_without_rate_or_days():
    """费率或期限没填时溢价为 0，而不是报错。"""
    contract = ContractData.from_products(
        [_product(amount=1567820.35)], contract_type="sales", prepayment_rate=0.20,
    )
    assert contract.compute_premium_amount() == 0.0
    contract.financing_rate = 0.125
    assert contract.compute_premium_amount() == 0.0
    contract.financing_days = 90
    assert contract.compute_premium_amount() == 39195.51


def test_premium_amount_is_zero_when_prepayment_covers_total():
    """预付款已覆盖全部货款（垫款为 0）时溢价为 0。"""
    contract = ContractData.from_products(
        [_product(amount=1000.0)], contract_type="sales", prepayment_rate=1.0,
    )
    contract.financing_rate = 0.125
    contract.financing_days = 90
    assert contract.compute_premium_amount() == 0.0
