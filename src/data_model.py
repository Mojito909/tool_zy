from dataclasses import dataclass, field


def _trailing_digits(text: str, n: int = 2) -> str:
    """取字符串末尾最多 n 位数字，供文件名序号使用。"""
    digits = ""
    for ch in reversed(text or ""):
        if not ch.isdigit() or len(digits) >= n:
            break
        digits = ch + digits
    return digits


@dataclass
class ProductRecord:
    seq_no: int
    contract_no: str = ""
    customer_name: str = ""
    yuhu_entity: str = ""
    supplier_name: str = ""
    product_name: str = ""
    country: str = ""
    factory_no: str = ""
    container_no: str = ""
    spec: str = ""
    quantity: int = 0
    weight: float = 0.0
    unit_price: float = 0.0
    amount: float = 0.0
    production_date: str = ""
    warehouse: str = ""


@dataclass
class ContractData:
    products: list[ProductRecord] = field(default_factory=list)
    contract_no: str = ""
    customer_name: str = ""
    yuhu_entity: str = ""
    supplier_name: str = ""
    country: str = ""
    factory_no: str = ""
    warehouse: str = ""
    total_quantity: int = 0
    total_weight: float = 0.0
    total_amount: float = 0.0
    total_amount_upper: str = ""
    ex_tax_amount: float = 0.0
    vat: float = 0.0
    contract_type: str = "purchase"
    vat_rate: float = 0.09
    vat_rate_pct: str = ""
    total_sales_amount: float = 0.0
    total_sales_amount_upper: str = ""
    ex_tax_sales_amount: float = 0.0
    sales_vat: float = 0.0
    prepayment: float = 0.0
    prepayment_rate: float = 0.0
    prepayment_rate_pct: str = ""
    sales_contract_no: str = ""
    master_agreement_no: str = ""
    serial: str = ""          # 合同号标记 + 末 2 位，如 P01 / S01（供文件名 pattern 使用）
    financing_days: int = 0   # 融资时长（天），生成销售合同时由用户输入
    financing_rate: float = 0.0    # 费率（溢价金额计算用），生成销售合同时由用户输入
    premium_amount: float = 0.0    # 溢价金额（仅溢价版销售订单使用）
    is_premium: bool = False       # 是否使用溢价版销售订单模板

    @classmethod
    def from_products(cls, products: list[ProductRecord], vat_rate: float = 0.09,
                      contract_type: str = "purchase",
                      unit_price_markup: float = 0.0,
                      prepayment_rate: float = 0.0) -> "ContractData":
        first = products[0] if products else ProductRecord(seq_no=0)
        total_quantity = sum(p.quantity for p in products)
        total_weight = sum(p.weight for p in products)
        total_amount = sum(p.amount for p in products)
        ex_tax_amount = round(total_amount / (1 + vat_rate), 2)
        vat = round(total_amount - ex_tax_amount, 2)

        total_sales_amount = 0.0
        ex_tax_sales_amount = 0.0
        sales_vat = 0.0
        prepayment = 0.0
        sales_contract_no = ""
        master_agreement_no = ""
        serial = ""
        if contract_type == "sales":
            total_sales_amount = round(
                sum((p.unit_price + unit_price_markup) * p.weight for p in products), 2
            )
            ex_tax_sales_amount = round(total_sales_amount / (1 + vat_rate), 2)
            sales_vat = round(total_sales_amount - ex_tax_sales_amount, 2)
            prepayment = round(total_amount * prepayment_rate, 2)
            cn = first.contract_no
            if len(cn) >= 3 and cn[-3] == "P":
                sales_contract_no = cn[:-3] + "S" + cn[-2:]
            else:
                sales_contract_no = cn
            master_agreement_no = cn[:-7] if len(cn) >= 7 else cn

        # 文件名用的编号标记：采购 P + 末 2 位，销售 S + 末 2 位
        # 文件名用的编号标记：P/S + 合同号末尾数字（最多 2 位）
        if contract_type == "sales":
            no = sales_contract_no or first.contract_no
            serial = "S" + _trailing_digits(no)
        else:
            serial = "P" + _trailing_digits(first.contract_no)
        return cls(
            products=products,
            contract_no=first.contract_no,
            customer_name=first.customer_name,
            yuhu_entity=first.yuhu_entity,
            supplier_name=first.supplier_name,
            country=first.country,
            factory_no=first.factory_no,
            warehouse=first.warehouse,
            total_quantity=total_quantity,
            total_weight=round(total_weight, 2),
            total_amount=round(total_amount, 2),
            total_amount_upper="",
            ex_tax_amount=ex_tax_amount,
            vat=vat,
            contract_type=contract_type,
            vat_rate=vat_rate,
            vat_rate_pct=f"{int(vat_rate * 100)}%",
            total_sales_amount=total_sales_amount,
            total_sales_amount_upper="",
            ex_tax_sales_amount=ex_tax_sales_amount,
            sales_vat=sales_vat,
            prepayment=prepayment,
            prepayment_rate=prepayment_rate,
            prepayment_rate_pct=f"{int(prepayment_rate * 100)}%",
            sales_contract_no=sales_contract_no,
            master_agreement_no=master_agreement_no,
            serial=serial,
        )

    def compute_premium_amount(self) -> float:
        """溢价金额 = 垫款总金额（采购合同总金额 − 预付款）× 费率 ÷ 360 × 提货期限天数。"""
        advance = self.total_amount - self.prepayment
        if advance <= 0 or self.financing_rate <= 0 or self.financing_days <= 0:
            return 0.0
        return round(advance * self.financing_rate / 360 * self.financing_days, 2)
