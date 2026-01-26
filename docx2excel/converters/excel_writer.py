# converters/excel_writer.py

from openpyxl import Workbook


def convert_value(value: str):
    """
    将字符串转换为 Excel 数值（int / float），
    如果无法转换则返回原字符串
    """
    if value is None:
        return None

    value = value.strip()

    if value == "":
        return ""

    # 尝试转 int
    try:
        return int(value)
    except ValueError:
        pass

    # 尝试转 float
    try:
        return float(value)
    except ValueError:
        pass

    # 仍然不是数字，返回字符串
    return value


def write_single_table_to_excel(table_data, output_path: str):
    """
    写入单个表格到 Excel（只有 Sheet1）
    自动将数字写为数值类型
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for row_idx, row in enumerate(table_data, start=1):
        for col_idx, value in enumerate(row, start=1):
            converted_value = convert_value(value)
            ws.cell(
                row=row_idx,
                column=col_idx,
                value=converted_value
            )

    wb.save(output_path)
