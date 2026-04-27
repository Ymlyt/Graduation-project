from calculator import calculate
from file_utils import get_file_path_from_drag_drop

def test_calculate_basic():
    y_list = [10, 5, 5, 5]
    rate = 2

    result = calculate(y_list, rate)

    assert len(result) == len(y_list)
    assert result[0] == 10
    assert all(isinstance(x, (int, float)) for x in result)


def test_calculate_formula_correctness():
    y_list = [10, 5, 2, 1]
    rate = 2

    a = ((10 - 5) * rate) / 5 + 1  # = 3.0
    result = calculate(y_list, rate)

    assert result[1] == 5 * a
    assert result[2] == 2 * a
    assert result[3] == 1 * a


def test_calculate_edge_case():
    y_list = [10, 1]
    rate = 1

    result = calculate(y_list, rate)
    assert len(result) == 2
