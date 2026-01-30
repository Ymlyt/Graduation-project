def find_intercept_turning_point(intercepts_list):
    # 初始化循环计数器
    loop_count = 0
    # 初始化停止索引
    stop_index = None

    # 从第二个元素（索引1，即Intercept_4）开始遍历
    for i in range(1, len(intercepts_list)):
        loop_count += 1  # 记录循环次数
        current_intercept = intercepts_list[i]
        previous_intercept = intercepts_list[i-1]
        
        # 判断当前截距是否大于前一个截距
        if current_intercept > previous_intercept:
            stop_index = i  # 记录停止时的索引
            break  # 满足条件，停止循环
        else:
        # 如果循环正常结束（没有遇到break），则说明没找到符合条件的点
            stop_index = len(intercepts_list)  # 这种情况下，stop_index 设为最后一个元素的索引

    # 计算停止时对应的 Intercept_n 的 n 值
    # 注意：intercepts_list 的索引 0 对应 Intercept_3，所以索引 i 对应 Intercept_{i+3}
    n_value_at_stop = stop_index + 3

    return loop_count, stop_index-1, n_value_at_stop-1

def get_file_paths_from_drag_drop():
    import os, sys

    if len(sys.argv) > 1:
        paths = []
        for p in sys.argv[1:]:
            if os.path.exists(p) and p.lower().endswith(('.xlsx', '.xls')):
                paths.append(p)

        if paths:
            print("通过拖放方式获取文件：")
            for p in paths:
                print(" ", p)
            return paths

    return []
