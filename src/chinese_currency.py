CHINESE_DIGITS = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"]
CHINESE_UNITS = ["", "拾", "佰", "仟"]
CHINESE_SECTIONS = ["", "万", "亿", "万亿"]


def _convert_integer_part(integer_str: str) -> str:
    """Convert the integer part of a number to Chinese uppercase."""
    if integer_str == "0":
        return "零"

    result = []
    n = len(integer_str)
    # 连续零只补一个"零"，遇到非零数字或小节单位时清空
    pending_zero = False

    for i, ch in enumerate(integer_str):
        digit = int(ch)
        pos = (n - i - 1) % 4            # 小节内的位次
        section_idx = (n - i - 1) // 4   # 第几个小节（万/亿）
        section_has_value = any(
            int(integer_str[j]) != 0 for j in range(max(0, i - 3), i + 1)
        )

        if digit != 0:
            if pending_zero:
                result.append("零")
                pending_zero = False
            result.append(CHINESE_DIGITS[digit])
            result.append(CHINESE_UNITS[pos])
        else:
            pending_zero = True

        # 小节末尾：本小节有值时补节单位（万/亿）
        if pos == 0 and section_idx > 0 and section_has_value:
            result.append(CHINESE_SECTIONS[section_idx])
            pending_zero = False

    text = "".join(result)
    if text.endswith("零"):
        text = text[:-1]
    return text


def amount_to_chinese_upper(amount: float) -> str:
    """Convert a float amount to Chinese uppercase currency string.

    Examples:
        154549.04 -> '壹拾伍万肆仟伍佰肆拾玖元零肆分'
        1257049.41 -> '壹佰贰拾伍万柒仟零肆拾玖元肆角壹分'
        100.00 -> '壹佰元整'
    """
    amount = round(amount, 2)
    integer_part = int(amount)
    fractional_part = int(round((amount - integer_part) * 100))

    integer_str = str(integer_part)
    jiao = fractional_part // 10
    fen = fractional_part % 10

    result = _convert_integer_part(integer_str)
    result += "元"

    if jiao == 0 and fen == 0:
        result += "整"
    else:
        if jiao > 0:
            result += CHINESE_DIGITS[jiao] + "角"
        else:
            result += "零"
        if fen > 0:
            result += CHINESE_DIGITS[fen] + "分"

    return result
