# converters/word_reader.py

from docx import Document

def extract_tables_from_word(word_path: str):
    """
    返回：
    [
        {
            "table_title": str,      # 表格第一行（去重后）
            "table_data": List[List[str]]  # 从第二行开始的数据
        }
    ]
    """
    document = Document(word_path)
    tables = []

    for table in document.tables:
        rows = []

        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            rows.append(row_data)

        if not rows:
            continue

        # ===== 第一行：用于 Excel 文件名（去重，解决合并单元格）=====
        first_row = rows[0]
        unique_cells = list(dict.fromkeys(first_row))
        table_title = "_".join(unique_cells)

        # ===== 表格内容：从第二行开始 =====
        table_data = rows[1:]

        tables.append({
            "table_title": table_title,
            "table_data": table_data
        })

    return tables
