import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from openpyxl import Workbook

class GradeCalculatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("成绩计算器")
        self.geometry("1000x800")
        
        # 存储科目参数
        self.subject_params = {
            "语文": {"满分": 120, "合格": 72, "优秀": 96, "良好": 84},
            "数学": {"满分": 120, "合格": 72, "优秀": 96, "良好": 84},
            "英语": {"满分": 120, "合格": 72, "优秀": 96, "良好": 84},
            "地理": {"满分": 60, "合格": 36, "优秀": 54, "良好": 48},
            "道法": {"满分": 60, "合格": 36, "优秀": 54, "良好": 48},
            "历史": {"满分": 60, "合格": 36, "优秀": 54, "良好": 48},
            "生物": {"满分": 60, "合格": 36, "优秀": 54, "良好": 48},
            "物理": {"满分": 60, "合格": 36, "优秀": 54, "良好": 48},
            "化学": {"满分": 60, "合格": 36, "优秀": 54, "良好": 48}
        }
        
        self.create_widgets()
    
    def create_widgets(self):
        # 文件选择按钮
        self.select_file_button = ttk.Button(self, text="选择Excel文件", command=self.select_file)
        self.select_file_button.pack(pady=10)
        
        # 显示结果的文本框
        self.result_text = tk.Text(self, height=40, width=120)
        self.result_text.pack(pady=10)
        
        # 设置科目参数按钮
        self.set_params_button = ttk.Button(self, text="设置科目参数", command=self.set_subject_params)
        self.set_params_button.pack(pady=10)
        
        # 导出按钮
        self.export_button = ttk.Button(self, text="导出结果", command=self.export_results)
        self.export_button.pack(pady=10)
        
        # 存储结果的DataFrame
        self.result_df = None
    
    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                self.calculate_grades(file_path)
            except Exception as e:
                messagebox.showerror("错误", f"处理文件时发生错误: {str(e)}")
    
    def calculate_grades(self, file_path):
        # 读取Excel文件
        try:
            # 跳过前两行
            df = pd.read_excel(file_path, engine='openpyxl', skiprows=2)
        except Exception as e:
            messagebox.showerror("错误", f"读取Excel文件失败: {str(e)}")
            return
        
        # 提取需要的列
        if "语文" not in df.columns or "数学" not in df.columns or "英语" not in df.columns:
            messagebox.showerror("错误", "Excel文件中缺少必要的列（语文、数学、英语等）")
            return
        
        # 计算每科的统计值
        result = {}
        subjects = ["语文", "数学", "英语", "地理", "道法", "历史", "生物"]
        
        for subject in subjects:
            # 获取参数
            total_score = self.subject_params[subject]["满分"]
            pass_score = self.subject_params[subject]["合格"]
            excellent_score = self.subject_params[subject]["优秀"]
            good_score = self.subject_params[subject]["良好"]
            
            # 计算各项指标
            scores = df[subject].dropna()
            total = scores.sum()
            count = len(scores)
            max_score = scores.max() if count > 0 else 0
            min_score = scores.min() if count > 0 else 0
            avg_score = scores.mean() if count > 0 else 0
            pass_count = len(scores[scores >= pass_score])
            excellent_count = len(scores[scores >= excellent_score])
            good_count = len(scores[(scores >= good_score) & (scores < excellent_score)])
            pass_rate = pass_count / count if count > 0 else 0
            excellent_rate = excellent_count / count if count > 0 else 0
            good_rate = good_count / count if count > 0 else 0
            avg_score_rate = avg_score / total_score if total_score > 0 else 0
            
            # 计算综合率
            composite_rate = 0.2 * avg_score_rate + 0.6 * pass_rate + 0.1 * excellent_rate + 0.1 * good_rate
            
            result[subject] = {
                "班级总分": total,
                "参加考试人数": count,
                "最高分": max_score,
                "最低分": min_score,
                "平均分": avg_score,
                "合格人数": pass_count,
                "合格率": pass_rate,
                "优秀人数": excellent_count,
                "优秀率": excellent_rate,
                "平均得分率": avg_score_rate,
                "良好人数": good_count,
                "良好率": good_rate,
                "综合率": composite_rate
            }
        
        # 创建结果DataFrame
        result_list = []
        for subject, metrics in result.items():
            row = {"学科": subject}
            row.update(metrics)
            result_list.append(row)
        
        # 添加合计行
        total_row = {"学科": "合计"}
        for metric in ["班级总分", "参加考试人数", "最高分", "最低分", "平均分", "合格人数", "优秀人数", "良好人数"]:
            if metric in result_list[0]:
                total_row[metric] = sum(item[metric] for item in result.values())
        
        for metric in ["合格率", "优秀率", "平均得分率", "良好率", "综合率"]:
            if metric in result_list[0]:
                total_row[metric] = sum(item[metric] for item in result.values()) / len(result)
        
        result_list.append(total_row)
        
        self.result_df = pd.DataFrame(result_list)
        
        # 显示结果
        self.display_results()
    
    def display_results(self):
        if self.result_df is None:
            return
        
        self.result_text.delete(1.0, tk.END)
        
        # 设置表头
        headers = ["学科", "班级总分", "参加考试人数", "最高分", "最低分", "平均分", 
                   "合格人数", "合格率", "优秀人数", "优秀率", "平均得分率", 
                   "良好人数", "良好率", "综合率"]
        
        # 格式化表头
        header_str = "{:<10} {:<12} {:<12} {:<8} {:<8} {:<8} {:<10} {:<8} {:<10} {:<8} {:<12} {:<10} {:<8} {:<10}\n".format(*headers)
        self.result_text.insert(tk.END, header_str)
        
        # 格式化表格内容
        for _, row in self.result_df.iterrows():
            row_str = "{:<10} {:<12.2f} {:<12} {:<8.2f} {:<8.2f} {:<8.2f} {:<10} {:<8.2%} {:<10} {:<8.2%} {:<12.2%} {:<10} {:<8.2%} {:<10.2%}\n".format(
                row["学科"],
                row["班级总分"],
                int(row["参加考试人数"]),
                row["最高分"],
                row["最低分"],
                row["平均分"],
                int(row["合格人数"]),
                row["合格率"],
                int(row["优秀人数"]),
                row["优秀率"],
                row["平均得分率"],
                int(row["良好人数"]),
                row["良好率"],
                row["综合率"]
            )
            self.result_text.insert(tk.END, row_str)
    
    def set_subject_params(self):
        # 创建新窗口用于设置科目参数
        param_window = tk.Toplevel(self)
        param_window.title("设置科目参数")
        param_window.geometry("800x600")
        
        frame = ttk.Frame(param_window)
        frame.pack(pady=20)
        
        # 表头
        headers = ["科目", "满分", "合格", "优秀", "良好"]
        for i, header in enumerate(headers):
            label = ttk.Label(frame, text=header, font=("Arial", 12, "bold"))
            label.grid(row=0, column=i, padx=10, pady=5)
        
        # 输入框
        entries = {}
        subjects = ["语文", "数学", "英语", "地理", "道法", "历史", "生物", "物理", "化学"]
        
        for i, subject in enumerate(subjects, start=1):
            # 科目名称
            subject_label = ttk.Label(frame, text=subject)
            subject_label.grid(row=i, column=0, padx=10, pady=5)
            
            # 满分
            total_entry = ttk.Entry(frame)
            total_entry.insert(0, str(self.subject_params[subject]["满分"]))
            total_entry.grid(row=i, column=1, padx=10, pady=5)
            entries[(subject, "满分")] = total_entry
            
            # 合格
            pass_entry = ttk.Entry(frame)
            pass_entry.insert(0, str(self.subject_params[subject]["合格"]))
            pass_entry.grid(row=i, column=2, padx=10, pady=5)
            entries[(subject, "合格")] = pass_entry
            
            # 优秀
            excellent_entry = ttk.Entry(frame)
            excellent_entry.insert(0, str(self.subject_params[subject]["优秀"]))
            excellent_entry.grid(row=i, column=3, padx=10, pady=5)
            entries[(subject, "优秀")] = excellent_entry
            
            # 良好
            good_entry = ttk.Entry(frame)
            good_entry.insert(0, str(self.subject_params[subject]["良好"]))
            good_entry.grid(row=i, column=4, padx=10, pady=5)
            entries[(subject, "良好")] = good_entry
        
        # 保存按钮
        def save_params():
            for subject in subjects:
                self.subject_params[subject]["满分"] = int(entries[(subject, "满分")].get())
                self.subject_params[subject]["合格"] = int(entries[(subject, "合格")].get())
                self.subject_params[subject]["优秀"] = int(entries[(subject, "优秀")].get())
                self.subject_params[subject]["良好"] = int(entries[(subject, "良好")].get())
            messagebox.showinfo("提示", "科目参数已保存")
            param_window.destroy()
        
        save_button = ttk.Button(param_window, text="保存", command=save_params)
        save_button.pack(pady=20)
    
    def export_results(self):
        if self.result_df is None:
            messagebox.showwarning("警告", "没有计算结果可导出")
            return
        
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            try:
                # 创建Excel工作簿
                wb = Workbook()
                ws = wb.active
                ws.title = "成绩统计表"
                
                # 转置数据
                transposed_df = self.result_df.set_index("学科").transpose()
                transposed_df.reset_index(inplace=True)
                transposed_df.rename(columns={"index": "分值/学科"}, inplace=True)
                
                # 写入转置表头
                ws.append(transposed_df.columns.tolist())
                
                # 写入转置数据
                for _, row in transposed_df.iterrows():
                    data_row = []
                    for item in row.tolist():
                        if isinstance(item, float) and item <= 1.0:  # 判断是否为率
                            data_row.append(f"{item*100:.1f}%")
                        else:
                            data_row.append(item)
                    ws.append(data_row)
                
                # 保存文件
                wb.save(file_path)
                messagebox.showinfo("提示", f"结果已导出到 {file_path}")
            except Exception as e:
                messagebox.showerror("错误", f"导出文件时发生错误: {str(e)}")

if __name__ == "__main__":
    app = GradeCalculatorApp()
    app.mainloop()
