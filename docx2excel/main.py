# main.py

import os
import sys
import tkinter as tk
from tkinter import filedialog

from config import EXCEL_OUTPUT_DIR
from utils.file_utils import validate_word_files, ensure_output_dir
from utils.sheet_name_utils import sanitize_sheet_name
from converters.word_reader import extract_tables_from_word
from converters.excel_writer import write_single_table_to_excel


def select_word_files():
    """
    弹出文件选择对话框，支持多选
    """
    root = tk.Tk()
    root.withdraw()

    file_paths = filedialog.askopenfilenames(
        title="请选择 Word 文件",
        filetypes=[("Word 文件", "*.docx")]
    )

    return list(file_paths)


def get_input_files():
    """
    优先使用拖拽文件，其次弹窗选择
    """
    if len(sys.argv) > 1:
        # 拖拽文件到程序
        return sys.argv[1:]
    else:
        # 手动选择
        return select_word_files()


def main():
    # 1. 获取输入文件
    input_files = get_input_files()
    input_files = validate_word_files(input_files)

    if not input_files:
        print("未选择任何有效的 Word 文件")
        return

    # 2. 确保输出目录存在
    ensure_output_dir(os.path.join(EXCEL_OUTPUT_DIR, "temp.xlsx"))

    # 3. 逐个 Word 文件处理
    for word_path in input_files:
        print(f"\n正在处理：{word_path}")

        docx_base_name = os.path.splitext(
            os.path.basename(word_path)
        )[0]

        tables = extract_tables_from_word(word_path)

        if not tables:
            print("  未检测到表格，已跳过")
            continue

        for table in tables:
            raw_title = table["table_title"]
            safe_title = sanitize_sheet_name(raw_title)

            excel_name = f"{docx_base_name}_{safe_title}.xlsx"
            excel_path = os.path.join(EXCEL_OUTPUT_DIR, excel_name)

            write_single_table_to_excel(
                table["table_data"],
                excel_path
            )

            print(f"  已生成：{excel_name}")

    print("\n全部文件处理完成")


if __name__ == "__main__":
    main()
