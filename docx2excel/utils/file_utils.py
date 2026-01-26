# utils/file_utils.py

import os

def validate_word_files(paths):
    """
    校验多个 Word 文件路径
    """
    valid_files = []

    for path in paths:
        if not os.path.exists(path):
            print(f"文件不存在，已跳过：{path}")
            continue

        if not path.lower().endswith(".docx"):
            print(f"非 docx 文件，已跳过：{path}")
            continue

        valid_files.append(path)

    return valid_files


def ensure_output_dir(path: str):
    directory = os.path.dirname(path)
    if directory and not os.path.exists(directory):
        os.makedirs(directory)
