import tkinter as tk
from tkinter import ttk
from tkinter import filedialog, messagebox
import sys
from output import get_file_paths_from_drag_drop

def select_excel_files():
    root = tk.Tk()
    root.withdraw()

    #  优先读取拖放的多个文件
    file_paths = get_file_paths_from_drag_drop()

    # 2️ 如果没拖文件，用文件选择框（支持多选）
    if not file_paths:
        file_paths = filedialog.askopenfilenames(
            title="选择Excel文件（可多选）",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )

        if not file_paths:
            messagebox.showinfo("提示", "未选择任何文件！程序将退出。")
            sys.exit(0)

    return list(file_paths)


def show_running_tip():
    tip = tk.Toplevel()
    tip.title('提示')
    tip.geometry('260x80')
    tip.resizable(False, False)
    label = tk.Label(tip, text='程序正在运行，请稍后...', font=("微软雅黑", 12))
    label.pack(expand=True, fill='both', padx=10, pady=20)
    tip.attributes('-topmost', True)
    tip.update()
    return tip


def close_running_tip(tip):
    if tip and tip.winfo_exists():
        tip.destroy()


def create_progress_window(total_files):
    win = tk.Toplevel()
    win.title("处理进度")
    win.geometry("400x120")
    win.resizable(False, False)

    label = tk.Label(win, text=f"正在处理文件 0 / {total_files}", font=("微软雅黑", 11))
    label.pack(pady=10)

    progress = ttk.Progressbar(
        win,
        orient="horizontal",
        length=360,
        mode="determinate",
        maximum=total_files
    )
    progress.pack(pady=10)

    win.attributes("-topmost", True)
    win.update()

    return win, label, progress