import time
import tempfile
import os

from converters.excel_writer import write_single_table_to_excel


def test_large_table_performance():
    table_data = [["123", "456"] for _ in range(5000)]

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        path = tmp.name

    try:
        start = time.time()

        write_single_table_to_excel(table_data, path)

        duration = time.time() - start

        print(f"\n处理5000行耗时: {duration:.2f}秒")

        # 给一个宽松标准（论文用）
        assert duration < 10

    finally:
        os.remove(path)