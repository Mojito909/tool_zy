import argparse
import sys
import os
from pathlib import Path

from .config_loader import load_config, ConfigError
from .excel_reader import read_excel, read_excel_header, SheetNotFoundError, ExcelReadError
from .field_mapper import map_to_products
from .data_model import ContractData
from .chinese_currency import amount_to_chinese_upper
from .word_generator import generate_contract

def _program_dir() -> str:
    """程序所在目录：冻结成 exe 时是 exe 同目录，否则是项目根目录（src 的上一级）。"""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _pick_folder_dialog(initial_dir: str) -> tuple[str | None, bool]:
    """弹出文件夹选择窗口。

    Returns:
        (选中的文件夹, 是否弹出过窗口)。取消选择或没有图形环境时文件夹为 None。
    """
    try:
        import tkinter
        from tkinter import filedialog
    except Exception:
        # 部分 Python 未编译 tkinter（无 _tkinter），退回手动输入
        return None, False

    root = None
    try:
        root = tkinter.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        folder = filedialog.askdirectory(
            title="请选择合同保存文件夹", initialdir=initial_dir or None
        )
    except Exception as e:
        print(f"  提示: 无法弹出文件夹选择窗口 ({e})")
        return None, False
    finally:
        if root is not None:
            try:
                root.destroy()
            except Exception:
                pass
    return (folder or None), True


def main(argv: list[str] | None = None):
    parser = argparse.ArgumentParser(
        description="Excel-to-Word 合同生成工具 — 从Excel产品明细表同时生成采购和销售合同"
    )
    parser.add_argument(
        "excel", help="Excel 文件路径（支持 .xls, .xlsx, .xlsm, .csv）"
    )
    parser.add_argument(
        "--config", "-c", default=None,
        help="配置文件路径（默认: 程序同目录下的 config.yaml）"
    )
    parser.add_argument(
        "--output", "-o", default=None,
        help="输出文件夹（覆盖配置文件中的 output.directory，同一次生成的两份合同都存到这里）"
    )
    parser.add_argument(
        "--sheet", "-s", default=None,
        help="指定 Sheet 名称（覆盖配置文件中的 Sheet 检测设置）"
    )
    args = parser.parse_args(argv)

    # Resolve config path
    config_path = args.config
    if config_path is None:
        # Try exe directory, then current directory
        base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        for candidate in [
            os.path.join(os.getcwd(), "config.yaml"),
            os.path.join(base, "config.yaml"),
        ]:
            if os.path.exists(candidate):
                config_path = candidate
                break

    if config_path is None or not os.path.exists(config_path):
        print("错误: 找不到配置文件 config.yaml", file=sys.stderr)
        print("请确保 config.yaml 与程序在同一目录", file=sys.stderr)
        sys.exit(2)

    try:
        print(f"加载配置文件: {config_path}")
        config = load_config(config_path)
    except ConfigError as e:
        print(f"配置错误: {e}", file=sys.stderr)
        sys.exit(2)

    # Override sheet name if specified
    if args.sheet:
        from .config_loader import SheetDetectionConfig
        config.sheet_detection.mode = "by_name"
        config.sheet_detection.pattern = args.sheet
        print(f"使用指定 Sheet: {args.sheet}")

    # Validate sales config (needed for both-contract mode)
    if not config.sales:
        print("错误: 配置文件中缺少 sales 节，无法生成销售合同", file=sys.stderr)
        print("请在 config.yaml 中添加 sales 配置节", file=sys.stderr)
        sys.exit(2)

    # Get VAT rate from user (shared by both contracts)
    vat_rate = config.tax.vat_rate
    if sys.stdin.isatty():
        try:
            tax_input = input(f"请输入税点百分比 (如9表示9%，默认{int(config.tax.vat_rate * 100)}): ").strip()
            if tax_input:
                vat_rate = float(tax_input) / 100.0
        except (EOFError, KeyboardInterrupt, ValueError):
            vat_rate = config.tax.vat_rate
    print(f"税点: {vat_rate * 100:.0f}%")

    # Get prepayment rate for sales contract
    prepayment_rate = config.sales.prepayment_rate
    if sys.stdin.isatty():
        try:
            rate_input = input(f"请输入预付款百分比 (如30表示30%，默认{int(config.sales.prepayment_rate * 100)}): ").strip()
            if rate_input:
                prepayment_rate = float(rate_input) / 100.0
        except (EOFError, KeyboardInterrupt, ValueError):
            pass
    print(f"预付款比例: {prepayment_rate * 100:.0f}%")

    # 融资时长（天）：配置取默认值，运行时可在窗口修改
    financing_days = config.sales.financing_days
    if sys.stdin.isatty():
        try:
            days_input = input(f"请输入融资时长/天 (默认{config.sales.financing_days}): ").strip()
            if days_input:
                financing_days = int(float(days_input))
        except (EOFError, KeyboardInterrupt, ValueError):
            pass
    print(f"融资时长: {financing_days} 天")

    # 销售订单模板版本 + 费率（溢价版才需要费率）
    use_premium = False
    financing_rate = 0.0
    if config.premium_sales:
        financing_rate = config.premium_sales.financing_rate
        if sys.stdin.isatty():
            print("请选择销售订单模板版本：")
            print("  1 = 旧版（老客户，默认）")
            print("  2 = 溢价版（新客户，含溢价金额）")
            try:
                ver = input("请输入 1 或 2 (直接回车 = 1): ").strip()
                use_premium = (ver == "2")
            except (EOFError, KeyboardInterrupt):
                pass
        if use_premium and sys.stdin.isatty():
            try:
                rate_input = input(
                    f"请输入费率百分比 (如5表示5%，默认{config.premium_sales.financing_rate * 100:g}): "
                ).strip()
                if rate_input:
                    financing_rate = float(rate_input) / 100.0
            except (EOFError, KeyboardInterrupt, ValueError):
                pass
        if use_premium:
            print(f"销售订单模板: 溢价版，费率 {financing_rate * 100:g}%")
        else:
            print("销售订单模板: 旧版")

    # 输出目录：-o > 窗口菜单 > 配置文件目录（非交互，如批量/CI）
    if args.output:
        output_dir = args.output
    elif sys.stdin.isatty():
        default_dir = _program_dir()
        print("请选择合同的保存位置：")
        print(f"  1 = 程序同目录（默认）: {default_dir}")
        print("  2 = 自定义文件夹（弹出选择窗口）")
        try:
            choice = input("请输入 1 或 2 (直接回车 = 1): ").strip()
            if choice == "2":
                folder, dialog_shown = _pick_folder_dialog(default_dir)
                if folder:
                    output_dir = folder
                elif dialog_shown:
                    # 窗口里点了取消
                    print("  未选择文件夹，改用程序同目录")
                    output_dir = default_dir
                else:
                    # 没有图形环境，退回手动输入路径
                    folder = input("请输入文件夹路径 (直接回车 = 程序同目录): ").strip().strip('"')
                    output_dir = folder or default_dir
            else:
                output_dir = default_dir
        except (EOFError, KeyboardInterrupt):
            output_dir = default_dir
    else:
        output_dir = config.output.directory
    print(f"输出目录: {output_dir}")




    # Read Excel
    excel_path = args.excel
    if not os.path.exists(excel_path):
        print(f"错误: Excel 文件不存在: {excel_path}", file=sys.stderr)
        sys.exit(1)

    try:
        print(f"读取 Excel 文件: {excel_path}")
        raw_rows = read_excel(excel_path, config)
        headers = read_excel_header(excel_path, config)
        print(f"读取到 {len(raw_rows)} 行产品数据")
        print(f"表头列名: {headers}")
    except SheetNotFoundError as e:
        print(f"Sheet 未找到: {e}", file=sys.stderr)
        sys.exit(3)
    except ExcelReadError as e:
        print(f"Excel 读取错误: {e}", file=sys.stderr)
        sys.exit(4)

    if not raw_rows:
        print("错误: 未读取到任何产品数据", file=sys.stderr)
        sys.exit(5)

    # Map to products (using header-name-based column mapping)
    products = map_to_products(raw_rows, config, headers)
    print(f"解析出 {len(products)} 个产品")

    if not products:
        print("错误: 未能解析任何产品数据", file=sys.stderr)
        sys.exit(5)

    # Build purchase contract data
    purchase = ContractData.from_products(
        products, vat_rate,
        contract_type="purchase",
    )
    purchase.total_amount_upper = amount_to_chinese_upper(purchase.total_amount)

    # Build sales contract data
    sales = ContractData.from_products(
        products, vat_rate,
        contract_type="sales",
        unit_price_markup=config.sales.unit_price_markup,
        prepayment_rate=prepayment_rate,
    )
    sales.total_sales_amount_upper = amount_to_chinese_upper(sales.total_sales_amount)
    sales.financing_days = financing_days
    sales.financing_rate = financing_rate
    sales.is_premium = use_premium
    if use_premium:
        sales.premium_amount = sales.compute_premium_amount()

    # 应用输出目录到两份合同
    config.output.directory = output_dir
    config.sales.output.directory = output_dir

    print(f"采购合同: 合同号={purchase.contract_no}, "
          f"供应商={purchase.supplier_name}, "
          f"总金额={purchase.total_amount:.2f}元, "
          f"大写={purchase.total_amount_upper}")
    print(f"销售合同: 采购合同号={sales.contract_no}, "
          f"销售合同号={sales.sales_contract_no}, "
          f"主协议号={sales.master_agreement_no}, "
          f"客户={sales.customer_name}, "
          f"总销售金额={sales.total_sales_amount:.2f}元, "
          f"大写={sales.total_sales_amount_upper}, "
          f"预付款={sales.prepayment:.2f}元")

    # Generate both contracts
    errors = []
    try:
        purchase_path = generate_contract(excel_path, config, purchase)
        print(f"\n采购合同已生成: {purchase_path}")
    except Exception as e:
        print(f"\n生成采购合同时出错: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        errors.append(("采购合同", e))

    try:
        sales_path = generate_contract(excel_path, config, sales)
        print(f"销售合同已生成: {sales_path}")
    except Exception as e:
        print(f"生成销售合同时出错: {e}", file=sys.stderr)
        import traceback
        traceback.print_exc()
        errors.append(("销售合同", e))

    if errors:
        sys.exit(6)

    return 0


if __name__ == "__main__":
    sys.exit(main())
