
from unittest.mock import MagicMock
import excel_runner


def test_run_excel_process(monkeypatch):

    fake_app = MagicMock()
    fake_book = MagicMock()
    fake_sheet = MagicMock()

    # ===== sheets =====
    fake_book.sheets = MagicMock()
    fake_book.sheets.__iter__.return_value = []
    fake_book.sheets.add.return_value = fake_sheet
    fake_book.sheets.__getitem__.return_value = fake_sheet

    def fake_range_func(*args, **kwargs):
        r = MagicMock()
        r.value = 1   # 所有读取都是数字
        return r

    fake_sheet.range.side_effect = fake_range_func

    # ===== mock xlwings =====
    monkeypatch.setattr("xlwings.App", lambda visible=False: fake_app)
    monkeypatch.setattr("xlwings.Book", lambda path: fake_book)

    # ===== mock 业务函数 =====
    monkeypatch.setattr(
        "excel_runner.find_intercept_turning_point",
        MagicMock(return_value=(10, 2, 5))
    )

    # ===== 执行 =====
    result = excel_runner.run_excel_process("test.xlsx")

    # ===== 断言 =====
    assert result == "test.xlsx"
    fake_book.save.assert_called_once()
    fake_book.close.assert_called()
    fake_app.quit.assert_called()