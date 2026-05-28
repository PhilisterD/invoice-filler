def num_to_chinese(n: int) -> str:
    """将整数金额转换为中文大写，如 3500 -> 叁仟伍佰元整"""
    if n == 0:
        return "零元整"
    units = ["", "拾", "佰", "仟"]
    nums = "零壹贰叁肆伍陆柒捌玖"
    result = ""
    unit_idx = 0
    zero_flag = False
    while n > 0:
        digit = n % 10
        n //= 10
        if digit == 0:
            if not zero_flag and result:
                zero_flag = True
        else:
            if zero_flag:
                result = nums[0] + result
                zero_flag = False
            result = nums[digit] + units[unit_idx] + result
        unit_idx += 1
    if result[-1] in units:
        result += "元整"
    else:
        result += "元整"
    result = result.replace("零零", "零").replace("零元", "元")
    if result.startswith("壹拾"):
        result = result[1:]
    return result

def format_date(excel_date) -> str:
    """将Excel日期转换为YYYY年MM月DD日格式"""
    import pandas as pd
    import numbers
    if pd.isna(excel_date):
        from datetime import datetime
        d = datetime.now()
    elif isinstance(excel_date, numbers.Number):
        from datetime import timedelta
        d = timedelta(days=int(excel_date)) + pd.Timestamp("1899-12-30")
    else:
        d = pd.to_datetime(excel_date)
    return f"{d.year}年{d.month:02d}月{d.day:02d}日"
