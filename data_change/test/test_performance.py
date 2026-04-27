import time
from calculator import calculate

def test_calculate_performance():
    y_list = list(range(1, 100000))
    rate = 1.5

    start = time.perf_counter()
    result = calculate(y_list, rate)
    end = time.perf_counter()

    assert len(result) == len(y_list)
    assert (end - start) < 1.0   # 1秒内完成（性能约束）
