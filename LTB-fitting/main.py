import tkinter as tk
from tkinter import messagebox
from ui_helper import (
    select_excel_files,
    show_running_tip,
    close_running_tip,
    create_progress_window
)
from excel_runner import run_excel_process

def main():
    try:
        file_paths = select_excel_files()
        total = len(file_paths)

        tip = show_running_tip()
        tip.withdraw()
        progress_win, label, progress = create_progress_window(total)

        success_files = []
        failed_files = []

        
        for idx, file_path in enumerate(file_paths, start=1):
            try:
                label.config(text=f"正在处理文件 {idx} / {total}")
                progress['value'] = idx - 1
                progress_win.update()

                file_name = run_excel_process(file_path)
                success_files.append(file_name)

                progress['value'] = idx
                progress_win.update()

            except Exception as e:
                failed_files.append((file_path, str(e)))

        progress_win.destroy()
        close_running_tip(tip)

        root = tk.Tk()
        root.withdraw()

        msg = f"成功处理 {len(success_files)} 个文件"
        if failed_files:
            msg += f"\n失败 {len(failed_files)} 个文件"

        messagebox.showinfo("处理完成", msg)
        root.destroy()

    except Exception as e:
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("错误", f"发生异常：{e}")
        root.destroy()


if __name__ == "__main__":
    main()
