import os
import tempfile
from utils.file_utils import validate_word_files, ensure_output_dir


def test_validate_word_files():
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
        valid_path = tmp.name

    invalid_path = valid_path + ".txt"
    not_exist_path = "not_exist.docx"

    try:
        result = validate_word_files([
            valid_path,
            invalid_path,
            not_exist_path
        ])

        assert valid_path in result
        assert invalid_path not in result
        assert not_exist_path not in result

    finally:
        os.remove(valid_path)


def test_ensure_output_dir():
    with tempfile.TemporaryDirectory() as tmp_dir:
        new_path = os.path.join(tmp_dir, "subdir", "file.xlsx")

        ensure_output_dir(new_path)

        assert os.path.exists(os.path.dirname(new_path))