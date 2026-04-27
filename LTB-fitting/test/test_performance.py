#test_performance.py

import time
from unittest.mock import MagicMock
import excel_runner


def test_excel_runner_performance(monkeypatch):

    fake_app = MagicMock()
    fake_book = MagicMock()
    fake_sheet = MagicMock()

    fake_book.sheets = MagicMock()
    fake_book.sheets.__iter__.return_value = []
    fake_book.sheets.add.return_value = fake_sheet
    fake_book.sheets.__getitem__.return_value = fake_sheet

    def fake_range(*args, **kwargs):
        r = MagicMock()
        r.value = 1.0
        return r

    fake_sheet.range.side_effect = fake_range

    monkeypatch.setattr("xlwings.App", lambda visible=False: fake_app)
    monkeypatch.setattr("xlwings.Book", lambda path: fake_book)

    monkeypatch.setattr(
        "excel_runner.find_intercept_turning_point",
        lambda x: (5, 1, 4)
    )

    start = time.time()

    for _ in range(10):
        excel_runner.run_excel_process("test.xlsx")

    end = time.time()

    assert (end - start) < 5  # 10次运行小于5秒