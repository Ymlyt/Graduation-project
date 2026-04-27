# test/test_output.py

from output import find_intercept_turning_point


def test_simple():
    assert 1 == 1


def test_turning_point_normal():
    data = [10, 9, 8, 11]

    loop_count, stop_index, n = find_intercept_turning_point(data)

    assert loop_count == 3


def test_turning_point_single():
    data = [10]

    loop_count, stop_index, n = find_intercept_turning_point(data)

    assert loop_count == 0