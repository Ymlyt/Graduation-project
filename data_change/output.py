
import os,sys
def get_file_path_from_drag_drop():
    """处理通过拖放方式传递的文件路径"""
    if len(sys.argv) > 1:
        drag_file_path = sys.argv[1]
        # 验证文件是否存在且是Excel文件
        if os.path.exists(drag_file_path) and drag_file_path.lower().endswith(('.xlsx', '.xls')):
            print(f"通过拖放方式获取文件: {drag_file_path}")
            return drag_file_path
        else:
            print(f"警告: 拖放的文件无效或不是Excel格式: {drag_file_path}")
    return None
  
