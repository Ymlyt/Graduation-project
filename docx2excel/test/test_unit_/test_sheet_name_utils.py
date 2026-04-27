from utils.sheet_name_utils import sanitize_sheet_name


def test_normal_name():
    assert sanitize_sheet_name("TestName") == "TestName"


def test_invalid_chars():
    name = "Test:/Name*?"
    result = sanitize_sheet_name(name)

    assert ":" not in result
    assert "/" not in result
    assert "*" not in result


def test_empty_name():
    assert sanitize_sheet_name("") == "Sheet"


def test_trim_spaces():
    assert sanitize_sheet_name("  abc  ") == "abc"


def test_length_limit():
    long_name = "a" * 50
    result = sanitize_sheet_name(long_name)

    assert len(result) == 31