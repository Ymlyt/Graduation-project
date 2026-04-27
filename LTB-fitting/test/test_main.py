#test_main.py

from unittest.mock import MagicMock
import main


def test_main_success(monkeypatch):

    # mock 文件选择
    monkeypatch.setattr(
        "main.select_excel_files",
        lambda: ["a.xlsx", "b.xlsx"]
    )

    # mock UI
    monkeypatch.setattr("main.show_running_tip", lambda: MagicMock())
    monkeypatch.setattr("main.close_running_tip", lambda tip: None)

    fake_win = MagicMock()
    fake_label = MagicMock()
    fake_progress = MagicMock()

    monkeypatch.setattr(
        "main.create_progress_window",
        lambda total: (fake_win, fake_label, fake_progress)
    )

    # mock Excel 处理
    monkeypatch.setattr(
        "main.run_excel_process",
        lambda x: x
    )

    # mock messagebox
    monkeypatch.setattr("tkinter.messagebox.showinfo", lambda *args, **kwargs: None)

    # mock Tk
    fake_root = MagicMock()
    monkeypatch.setattr("tkinter.Tk", lambda: fake_root)

    main.main()

    assert True  # 能跑完即可


def test_main_with_error(monkeypatch):

    monkeypatch.setattr(
        "main.select_excel_files",
        lambda: ["a.xlsx"]
    )

    monkeypatch.setattr("main.show_running_tip", lambda: MagicMock())
    monkeypatch.setattr("main.close_running_tip", lambda tip: None)

    fake_win = MagicMock()
    fake_label = MagicMock()
    fake_progress = MagicMock()

    monkeypatch.setattr(
        "main.create_progress_window",
        lambda total: (fake_win, fake_label, fake_progress)
    )

    # 模拟报错
    def raise_error(x):
        raise Exception("fail")

    monkeypatch.setattr("main.run_excel_process", raise_error)

    monkeypatch.setattr("tkinter.messagebox.showinfo", lambda *args, **kwargs: None)
    monkeypatch.setattr("tkinter.messagebox.showerror", lambda *args, **kwargs: None)

    fake_root = MagicMock()
    monkeypatch.setattr("tkinter.Tk", lambda: fake_root)

    main.main()

    assert True