# utils/sheet_name_utils.py

import re

INVALID_CHARS = r'[\\\/\?\*\[\]\:]'

def sanitize_sheet_name(name: str, max_length: int = 31) -> str:
    """
    清洗 Excel Sheet 名称
    """
    if not name:
        return "Sheet"

    # 去除非法字符
    name = re.sub(INVALID_CHARS, "_", name)

    # 去除首尾空白
    name = name.strip()

    # 限制长度
    return name[:max_length]
