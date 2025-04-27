import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

# 读取并解析 Excel 文件，获取所有班级的科目
def load_data(file_path):
    xls = pd.ExcelFile(file_path)
    sheet_names = xls.sheet_names
    class_subjects = {}

    for sheet in sheet_names:
        df = xls.parse(sheet)
        # 假设科目从第四行的第三列开始
        subjects = df.iloc[3, 2:].dropna().tolist()
        class_subjects[sheet] = subjects

    return class_subjects

# 自动生成界面，基于科目设置输入框
def create_input_frame(root, class_subjects):
    input_frame = tk.Frame(root)
    input_frame.pack(pady=20)

    entries = {}
    for class_name, subjects in class_subjects.items():
        label = tk.Label(input_frame, text=f"{class_name} 科目设置")
        label.grid(row=0, column=0, padx=10, pady=5, sticky='w')

        # 为每个科目生成输入框
        for i, subject in enumerate(subjects):
            label = tk.Label(input_frame, text=subject)
            label.grid(row=i + 1, column=0, padx=10, pady=5, sticky='w')

            total_entry = tk.Entry(input_frame)
            total_entry.grid(row=i + 1, column=1)
            total_entry.insert(0, 150)  # 默认值为150
            entries[(class_name, subject)] = total_entry

    return entries

# 计算并保存结果
def calculate_and_save(entries, output_file, file_path):
    try:
        class_subjects = load_data(file_path)
        
        params = {}
        for (class_name, subject), entry in entries.items():
            total_score = int(entry.get())
            if class_name not in params:
                params[class_name] = {}
            params[class_name][subject] = total_score
        
        # 在这里进行成绩统计和汇总计算的逻辑
        # ...

        # 在此添加生成Excel的代码
        
        # 弹出成功信息
        messagebox.showinfo("成功", f"成绩汇总已保存至 {output_file}")
    except Exception as e:
        messagebox.showerror("错误", f"发生错误：{e}")

# 主窗口
def main():
    root = tk.Tk()
    root.title("成绩汇总生成工具")

    # 选择文件
    def select_file():
        file_path = filedialog.askopenfilename(title="选择成绩文件", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
        if file_path:
            # 加载并显示科目设置输入框
            class_subjects = load_data(file_path)
            entries = create_input_frame(root, class_subjects)

            # 选择输出文件位置
            def select_output():
                output_file = filedialog.asksaveasfilename(title="选择保存文件", defaultextension=".xlsx", filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))
                if output_file:
                    calculate_and_save(entries, output_file, file_path)

            # 输出文件路径按钮
            output_button = tk.Button(root, text="选择输出文件", command=select_output)
            output_button.pack(pady=20)

    # 选择文件按钮
    select_button = tk.Button(root, text="选择成绩文件", command=select_file)
    select_button.pack(pady=20)

    root.mainloop()

if __name__ == "__main__":
    main()
