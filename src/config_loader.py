import yaml
from pathlib import Path
from dataclasses import dataclass, field


class ConfigError(Exception):
    pass


@dataclass
class SheetDetectionConfig:
    mode: str = "by_pattern"
    pattern: str = ""


@dataclass
class DataConfig:
    header_row: int = 1
    data_start_row: int = 2
    stop_on_empty_column: str = ""


@dataclass
class TransformRule:
    type: str = "text"            # text | number | excel_date | to_string
    format: str = "%Y-%m-%d"
    decimal_places: int = 2
    default: object = ""


@dataclass
class WordConfig:
    template_path: str = "templates/采购合同模板.docx"
    placeholders: dict = field(default_factory=dict)
    product_table_index: int = 0
    product_template_row: int = 1
    signature_table_index: int = 1


@dataclass
class TaxConfig:
    vat_rate: float = 0.09


@dataclass
class OutputConfig:
    pattern: str = "合同.docx"
    directory: str = "./output"


@dataclass
class SalesConfig:
    unit_price_markup: float = 0.2
    prepayment_rate: float = 0.3
    financing_days: int = 0
    word: WordConfig = field(default_factory=WordConfig)
    output: OutputConfig = field(default_factory=OutputConfig)
    cleanup_placeholders: list[str] = field(default_factory=list)


@dataclass
class PremiumSalesConfig:
    """溢价版销售订单（新客户专用）设置。"""
    template_path: str = "templates/销售订单溢价版.docx"
    financing_rate: float = 0.05       # 费率，窗口可改
    use_thousands_separator: bool = True   # 金额是否带千分位
    placeholders: dict = field(default_factory=dict)
    product_table_index: int = 0
    product_template_row: int = 2
    signature_table_index: int = 1
    output_prefix: str = "销售订单溢价版"


@dataclass
class Config:
    sheet_detection: SheetDetectionConfig = field(default_factory=SheetDetectionConfig)
    data: DataConfig = field(default_factory=DataConfig)
    column_mapping: dict[str, str] = field(default_factory=dict)
    transformations: dict[str, TransformRule] = field(default_factory=dict)
    word: WordConfig = field(default_factory=WordConfig)
    tax: TaxConfig = field(default_factory=TaxConfig)
    output: OutputConfig = field(default_factory=OutputConfig)
    default_type: str = "purchase"
    sales: SalesConfig | None = None
    premium_sales: PremiumSalesConfig | None = None


def load_config(config_path: str | Path) -> Config:
    """Load and validate configuration from a YAML file."""
    config_path = Path(config_path)
    if not config_path.exists():
        raise ConfigError(f"配置文件不存在: {config_path}")

    with open(config_path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    if raw is None:
        raise ConfigError(f"配置文件为空: {config_path}")

    config = Config()

    # Sheet detection
    sd = raw.get("sheet_detection", {})
    config.sheet_detection = SheetDetectionConfig(
        mode=sd.get("mode", "by_pattern"),
        pattern=sd.get("pattern", ""),
    )

    # Data
    d = raw.get("data", {})
    config.data = DataConfig(
        header_row=int(d.get("header_row", 1)),
        data_start_row=int(d.get("data_start_row", 2)),
        stop_on_empty_column=str(d.get("stop_on_empty_column", "")),
    )

    # Column mapping
    cm = raw.get("column_mapping", {})
    config.column_mapping = {str(k): str(v) for k, v in cm.items()}

    # Transformations
    tx = raw.get("transformations", {})
    for field_name, rule in tx.items():
        config.transformations[field_name] = TransformRule(
            type=rule.get("type", "text"),
            format=rule.get("format", "%Y-%m-%d"),
            decimal_places=int(rule.get("decimal_places", 2)),
            default=rule.get("default", ""),
        )

    # Word
    w = raw.get("word", {})
    pt = w.get("product_table", {})
    st = w.get("signature_table", {})
    config.word = WordConfig(
        template_path=str(w.get("template_path", "templates/采购合同模板.docx")),
        placeholders=w.get("placeholders", {}),
        product_table_index=int(pt.get("table_index", 0)),
        product_template_row=int(pt.get("template_row_index", 1)),
        signature_table_index=int(st.get("table_index", 1)),
    )

    # Tax
    t = raw.get("tax", {})
    config.tax = TaxConfig(vat_rate=float(t.get("vat_rate", 0.09)))

    # Output
    o = raw.get("output", {})
    config.output = OutputConfig(
        pattern=str(o.get("pattern", "合同.docx")),
        directory=str(o.get("directory", "./output")),
    )

    # Default type
    config.default_type = str(raw.get("default_type", "purchase"))

    # Sales config
    sales_raw = raw.get("sales")
    if sales_raw:
        sw = sales_raw.get("word", {})
        s_pt = sw.get("product_table", {})
        s_st = sw.get("signature_table", {})
        sales_word = WordConfig(
            template_path=str(sw.get("template_path", "templates/销售合同模板.docx")),
            placeholders=sw.get("placeholders", {}),
            product_table_index=int(s_pt.get("table_index", config.word.product_table_index)),
            product_template_row=int(s_pt.get("template_row_index", config.word.product_template_row)),
            signature_table_index=int(s_st.get("table_index", config.word.signature_table_index)),
        )
        so = sales_raw.get("output", {})
        sales_output = OutputConfig(
            pattern=str(so.get("pattern", config.output.pattern)),
            directory=str(so.get("directory", config.output.directory)),
        )
        config.sales = SalesConfig(
            unit_price_markup=float(sales_raw.get("unit_price_markup", 0.2)),
            prepayment_rate=float(sales_raw.get("prepayment_rate", 0.3)),
            financing_days=int(sales_raw.get("financing_days", 0)),
            word=sales_word,
            output=sales_output,
            cleanup_placeholders=sales_raw.get("cleanup_placeholders", []),
        )

    # 溢价版销售订单（新客户专用）：默认继承标准销售模板的占位符，再叠加溢价专用项
    prem_raw = raw.get("premium_sales")
    if prem_raw:
        base_placeholders = dict(config.sales.word.placeholders) if config.sales else {}
        prem_placeholders = dict(base_placeholders)
        prem_placeholders.update(prem_raw.get("placeholders", {}))
        prem_pt = prem_raw.get("product_table", {})
        prem_st = prem_raw.get("signature_table", {})
        config.premium_sales = PremiumSalesConfig(
            template_path=str(prem_raw.get(
                "template_path",
                config.sales.word.template_path if config.sales
                else "templates/销售订单溢价版.docx",
            )),
            financing_rate=float(prem_raw.get("financing_rate", 0.05)),
            use_thousands_separator=bool(prem_raw.get("use_thousands_separator", True)),
            placeholders=prem_placeholders,
            product_table_index=int(prem_pt.get(
                "table_index", config.sales.word.product_table_index if config.sales else 0)),
            product_template_row=int(prem_pt.get(
                "template_row_index",
                2 if not config.sales else config.sales.word.product_template_row)),
            signature_table_index=int(prem_st.get(
                "table_index", config.sales.word.signature_table_index if config.sales else 1)),
            output_prefix=str(prem_raw.get("output_prefix", "销售订单溢价版")),
        )
        _validate_premium(config.premium_sales)
    _validate(config)
    return config


def _validate_premium(prem: PremiumSalesConfig):
    if prem.product_template_row < 0:
        raise ConfigError("premium_sales.product_table.template_row_index 不能为负数")
    if prem.financing_rate < 0:
        raise ConfigError(f"premium_sales.financing_rate 不能为负数，当前值: {prem.financing_rate}")


def _validate(config: Config):
    if config.sheet_detection.mode not in ("by_name", "by_pattern", "by_index", "first"):
        raise ConfigError(f"无效的 sheet_detection.mode: {config.sheet_detection.mode}")
    if config.data.header_row < 0:
        raise ConfigError("data.header_row 不能为负数")
    if config.data.data_start_row < 0:
        raise ConfigError("data.data_start_row 不能为负数")
    if config.word.product_template_row < 0:
        raise ConfigError("word.product_table.template_row_index 不能为负数")
    if config.tax.vat_rate < 0 or config.tax.vat_rate >= 1:
        raise ConfigError(f"tax.vat_rate 应在 0~1 之间，当前值: {config.tax.vat_rate}")
