import sys
import tkinter as tk  
import xlwings as xw
from tkinter import messagebox, filedialog,simpledialog
from pathlib import Path

from output import get_file_path_from_drag_drop
from calculator import calculate


try:
    root = tk.Tk()
    root.withdraw() 
    file_path = get_file_path_from_drag_drop()
    
    # 如果没有拖放文件，则使用文件对话框
    if not file_path:
    

        file_path = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel Files", "*.xlsx *.xls")] # 可以设置多个类型，如：[("Excel Files", "*.xlsx"), ("All Files", "*.*")]
        )
        print("您选择的Excel文件路径是：", file_path)

        if not file_path:
            messagebox.showinfo("提示", "未选择任何文件！程序将退出。")
            sys.exit(0)
    if isinstance(file_path, (list, tuple)):
        file_path = file_path[0]  # 取第一个文件（多文件拖放时）
    
    print("最终处理的Excel文件路径是:", file_path)
    
    
    rate = simpledialog.askfloat("输入", "请输入rate的值:")
    print(f"rate的值为: {rate}")
    
    
    
            
    running_tip = tk.Toplevel()
    running_tip.title('提示')
    running_tip.geometry('260x80')
    running_tip.resizable(False, False)
    label = tk.Label(running_tip, text='程序正在运行，请稍后...', font=("微软雅黑", 12))
    label.pack(expand=True, fill='both', padx=10, pady=20)
    running_tip.update()
    # 让提示窗口始终在最前
    running_tip.attributes('-topmost', True)
    path_obj = Path(file_path)
    # 通过 .name 属性获取文件名
    file_name = path_obj.name
    
    
    app = xw.App(visible=False)
    wb = xw.Book(file_path)
    sheet1 = wb.sheets['Sheet1']
    
    y_list = sheet1.range('B2:B17').value
    n_y_list = calculate(y_list, rate)
    print("\n计算后的n_y_list:", n_y_list)
    print("\n")

    
    sheet1.range('B3').value = n_y_list[1]
    sheet1.range('B4').value = n_y_list[2]
    sheet1.range('B5').value = n_y_list[3]        
    sheet1.range('B6').value = n_y_list[4]
    sheet1.range('B7').value = n_y_list[5]
    sheet1.range('B8').value = n_y_list[6]
    sheet1.range('B9').value = n_y_list[7]
    sheet1.range('B10').value = n_y_list[8]
    sheet1.range('B11').value = n_y_list[9]
    sheet1.range('B12').value = n_y_list[10]
    sheet1.range('B13').value = n_y_list[11]
    sheet1.range('B14').value = n_y_list[12]
    sheet1.range('B15').value = n_y_list[13]
    sheet1.range('B16').value = n_y_list[14]
    sheet1.range('B17').value = n_y_list[15]
    
    sheet1.range('D2').value = y_list[0]
    sheet1.range('D3').value = y_list[1]
    sheet1.range('D4').value = y_list[2]
    sheet1.range('D5').value = y_list[3]        
    sheet1.range('D6').value = y_list[4]
    sheet1.range('D7').value = y_list[5]
    sheet1.range('D8').value = y_list[6]
    sheet1.range('D9').value = y_list[7]
    sheet1.range('D10').value = y_list[8]
    sheet1.range('D11').value = y_list[9]
    sheet1.range('D12').value = y_list[10]
    sheet1.range('D13').value = y_list[11]
    sheet1.range('D14').value = y_list[12]
    sheet1.range('D15').value = y_list[13]
    sheet1.range('D16').value = y_list[14]
    sheet1.range('D17').value = y_list[15]
    
    wb.app.calculate()
    wb.save()
    wb.close()
    app.quit()
    
    
    if running_tip.winfo_exists():
        running_tip.destroy()
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("成功", "文件%s数据修改成功！" % file_name)
    root.destroy()
    
    
    
except Exception as e:
    # 捕获异常，尝试关闭Excel进程
    if 'running_tip' in locals() and running_tip.winfo_exists():
        running_tip.destroy()
    root = tk.Tk()
    root.withdraw()
    print("发生异常：", e)
    try:
        wb.close()
        app.quit()
        messagebox.showerror("错误", f"发生异常：{e}\nExcel进程已正常关闭，请重新运行。")
        root.destroy()
    except Exception as close_e:
        print("关闭Excel进程时出错：", close_e)
        try:
            app.kill()  # 强制杀死由xlwings启动的Excel进程
            messagebox.showerror("错误", f"发生异常：{e}\n进程关闭失败，但已强制终止，请重新运行。")
        except Exception as kill_e:
            print("强制终止Excel进程时出错：", kill_e)
            messagebox.showerror("错误", f"发生异常：{e}\n进程关闭失败，请手动从任务管理器结束Excel进程后重新运行。")
        finally:
            root.destroy()
finally:
    # 防止未关闭的Excel进程残留
    try:
        if 'running_tip' in locals() and running_tip.winfo_exists():
            running_tip.destroy()
    except:
        pass
    try:
        wb.close()
    except:
        pass
    try:
        app.quit()
    except:
        pass