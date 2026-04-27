import sys
from unittest.mock import patch
from file_utils import get_file_path_from_drag_drop

def test_get_file_path_from_drag_drop_valid(monkeypatch):
    monkeypatch.setattr(sys, "argv", ["main.py", "test.xlsx"])
    with patch("os.path.exists", return_value=True):
        result = get_file_path_from_drag_drop()
        assert result == "test.xlsx"


def test_get_file_path_from_drag_drop_invalid(monkeypatch):
    monkeypatch.setattr(sys, "argv", ["main.py", "test.txt"])
    with patch("os.path.exists", return_value=True):
        result = get_file_path_from_drag_drop()
        assert result is None


def test_get_file_path_from_drag_drop_no_args(monkeypatch):
    monkeypatch.setattr(sys, "argv", ["main.py"])
    result = get_file_path_from_drag_drop()
    assert result is None
