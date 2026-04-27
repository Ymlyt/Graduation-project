# test_ui_helper.py

from unittest.mock import MagicMock
import ui_helper


def test_select_excel_files(monkeypatch):

    monkeypatch.setattr(
        "ui_helper.get_file_paths_from_drag_drop",
        lambda: ["a.xlsx", "b.xlsx"]
    )

    files = ui_helper.select_excel_files()

    assert files == ["a.xlsx", "b.xlsx"]


def test_show_and_close_tip(monkeypatch):

    fake_tip = MagicMock()
    fake_tip.winfo_exists.return_value = True

    monkeypatch.setattr("tkinter.Toplevel", lambda: fake_tip)

    tip = ui_helper.show_running_tip()
    ui_helper.close_running_tip(tip)

    fake_tip.destroy.assert_called()


def test_create_progress_window(monkeypatch):

    fake_win = MagicMock()
    fake_label = MagicMock()
    fake_progress = MagicMock()

    monkeypatch.setattr("tkinter.Toplevel", lambda: fake_win)
    monkeypatch.setattr("tkinter.Label", lambda *args, **kwargs: fake_label)
    monkeypatch.setattr("tkinter.ttk.Progressbar", lambda *args, **kwargs: fake_progress)

    win, label, progress = ui_helper.create_progress_window(5)

    assert win == fake_win