import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import random
import string
import os
from datetime import datetime

# 中文字符集用于随机生成姓名
FAMILY_NAMES = ["赵", "钱", "孙", "李", "周", "吴", "郑", "王", "冯", "陈", 
                "褚", "卫", "蒋", "沈", "韩", "杨", "朱", "秦", "尤", "许",
                "何", "吕", "施", "张", "孔", "曹", "严", "华", "金", "魏"]

GIVEN_NAMES = ["伟", "芳", "娜", "秀英", "敏", "静", "强", "磊", "军", "洋",
               "勇", "艳", "杰", "娟", "涛", "明", "超", "秀兰", "霞", "平",
               "刚", "桂", "文", "辉", "林", "波", "中", "宇", "洪", "梅"]

# 班级列表
CLASSES = ["高一(1)班", "高一(2)班", "高一(3)班", "高一(4)班", "高一(5)班",
           "高二(1)班", "高二(2)班", "高二(3)班", "高二(4)班", "高二(5)班",
           "高三(1)班", "高三(2)班", "高三(3)班", "高三(4)班", "高三(5)班"] 

# 科目列表
SUBJECTS = [
    {"name": "德育", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "语文", "min": 0, "max": 150, "pass_score": 90, "pass_rate": 60},
    {"name": "数学", "min": 0, "max": 150, "pass_score": 90, "pass_rate": 60},
    {"name": "外语", "min": 0, "max": 150, "pass_score": 90, "pass_rate": 60},
    {"name": "物理", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "历史", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "生物", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "地理", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "化学", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "政治", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "科学", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "信息技术", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60},
    {"name": "时事政治", "min": 0, "max": 100, "pass_score": 60, "pass_rate": 60}
]

class StudentDataGenerator:
    def __init__(self, root):
        self.root = root
        self.root.title("学生信息与成绩随机生成工具")
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        
        # 存储科目设置的变量
        self.subject_vars = {}
        self.subject_min = {}
        self.subject_max = {}
        self.subject_pass = {}
        self.subject_pass_rate = {}
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 顶部设置区域
        settings_frame = ttk.LabelFrame(main_frame, text="生成设置", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 10))
        
        # 学生数量设置
        ttk.Label(settings_frame, text="学生数量:").grid(row=0, column=0, padx=5, pady=5, sticky=tk.W)
        self.student_count = tk.IntVar(value=50)
        ttk.Spinbox(settings_frame, from_=1, to=1000, textvariable=self.student_count, width=10).grid(
            row=0, column=1, padx=5, pady=5, sticky=tk.W)
        
        # 学号前缀设置
        ttk.Label(settings_frame, text="学号前缀:").grid(row=0, column=2, padx=5, pady=5, sticky=tk.W)
        self.id_prefix = tk.StringVar(value="2023")
        ttk.Entry(settings_frame, textvariable=self.id_prefix, width=10).grid(
            row=0, column=3, padx=5, pady=5, sticky=tk.W)
        
        # 班级设置
        ttk.Label(settings_frame, text="班级范围:").grid(row=0, column=4, padx=5, pady=5, sticky=tk.W)
        self.class_range = tk.StringVar(value="随机分配")
        class_combo = ttk.Combobox(settings_frame, textvariable=self.class_range, width=15, state="readonly")
        class_combo['values'] = ["随机分配"] + CLASSES
        class_combo.grid(row=0, column=5, padx=5, pady=5, sticky=tk.W)
        
        # 科目设置区域
        subjects_frame = ttk.LabelFrame(main_frame, text="科目设置", padding="10")
        subjects_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # 创建滚动区域
        canvas = tk.Canvas(subjects_frame)
        scrollbar = ttk.Scrollbar(subjects_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # 添加科目设置行
        headers = ["选择", "科目名称", "最小值", "最大值", "及格分", "及格率(%)"]
        for col, header in enumerate(headers):
            ttk.Label(scrollable_frame, text=header, font=("Arial", 9, "bold")).grid(
                row=0, column=col, padx=5, pady=5, sticky=tk.W)
        
        # 为每个科目创建设置项
        for row, subject in enumerate(SUBJECTS, 1):
            name = subject["name"]
            
            # 选择框
            var = tk.BooleanVar(value=True)
            self.subject_vars[name] = var
            ttk.Checkbutton(scrollable_frame, variable=var).grid(
                row=row, column=0, padx=5, pady=5)
            
            # 科目名称
            ttk.Label(scrollable_frame, text=name).grid(
                row=row, column=1, padx=5, pady=5, sticky=tk.W)
            
            # 最小值
            min_var = tk.IntVar(value=subject["min"])
            self.subject_min[name] = min_var
            ttk.Entry(scrollable_frame, textvariable=min_var, width=8).grid(
                row=row, column=2, padx=5, pady=5)
            
            # 最大值
            max_var = tk.IntVar(value=subject["max"])
            self.subject_max[name] = max_var
            ttk.Entry(scrollable_frame, textvariable=max_var, width=8).grid(
                row=row, column=3, padx=5, pady=5)
            
            # 及格分
            pass_var = tk.IntVar(value=subject["pass_score"])
            self.subject_pass[name] = pass_var
            ttk.Entry(scrollable_frame, textvariable=pass_var, width=8).grid(
                row=row, column=4, padx=5, pady=5)
            
            # 及格率
            pass_rate_var = tk.IntVar(value=subject["pass_rate"])
            self.subject_pass_rate[name] = pass_rate_var
            ttk.Entry(scrollable_frame, textvariable=pass_rate_var, width=8).grid(
                row=row, column=5, padx=5, pady=5)
        
        # 按钮区域
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(button_frame, text="生成数据", command=self.generate_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="导出数据", command=self.export_data).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="重置设置", command=self.reset_settings).pack(side=tk.RIGHT, padx=5)
        
        # 结果展示区域
        result_frame = ttk.LabelFrame(main_frame, text="生成结果预览", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True)
        
        # 结果表格
        table_frame = ttk.Frame(result_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        
        # 树状表格
        self.result_tree = ttk.Treeview(
            table_frame,
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        scrollbar_x.config(command=self.result_tree.xview)
        scrollbar_y.config(command=self.result_tree.yview)
        
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.result_tree.pack(fill=tk.BOTH, expand=True)
        
        # 存储生成的数据
        self.generated_data = None
    
    def generate_name(self):
        """随机生成姓名"""
        family_name = random.choice(FAMILY_NAMES)
        given_name = random.choice(GIVEN_NAMES)
        return family_name + given_name
    
    def generate_id(self, index):
        """生成学号"""
        prefix = self.id_prefix.get()
        # 生成固定长度的序号，确保学号长度一致
        return f"{prefix}{index + 1:04d}"
    
    def generate_class(self):
        """生成班级"""
        class_range = self.class_range.get()
        if class_range == "随机分配":
            return random.choice(CLASSES)
        return class_range
    
    def generate_scores_for_subject(self, subject, student_count):
        """为特定科目生成所有学生的成绩，确保及格率不低于设定值"""
        name = subject["name"]
        min_val = self.subject_min[name].get()
        max_val = self.subject_max[name].get()
        pass_score = self.subject_pass[name].get()
        pass_rate = self.subject_pass_rate[name].get() / 100
        
        # 确保参数有效
        if min_val > max_val:
            min_val, max_val = max_val, min_val
        
        # 计算最少需要多少个及格成绩
        min_pass_count = int(pass_rate * student_count)
        # 如果计算结果为0但及格率大于0，至少保证1个及格成绩
        if min_pass_count == 0 and pass_rate > 0:
            min_pass_count = 1
        
        # 先生成确保达标的及格成绩
        scores = []
        # 添加必要的及格成绩
        for _ in range(min_pass_count):
            scores.append(random.randint(pass_score, max_val))
        
        # 生成剩余的成绩，可以是及格或不及格
        remaining = student_count - min_pass_count
        for _ in range(remaining):
            # 50%概率生成及格成绩，使最终及格率可能高于设定值
            if random.random() < 0.5:
                scores.append(random.randint(pass_score, max_val))
            else:
                scores.append(random.randint(min_val, pass_score - 1))
        
        # 打乱成绩顺序
        random.shuffle(scores)
        return scores
    
    def generate_data(self):
        """生成学生数据"""
        try:
            count = self.student_count.get()
            if count <= 0:
                messagebox.showwarning("警告", "学生数量必须大于0")
                return
            
            # 收集选中的科目
            selected_subjects = [subj for subj in SUBJECTS 
                               if self.subject_vars[subj["name"]].get()]
            
            if not selected_subjects:
                messagebox.showwarning("警告", "请至少选择一个科目")
                return
            
            # 准备数据结构
            data = []
            columns = ["序号", "学号", "姓名", "班级"] + [subj["name"] for subj in selected_subjects]
            
            # 生成基本信息
            basic_info = []
            for i in range(count):
                basic_info.append({
                    "序号": i + 1,
                    "学号": self.generate_id(i),
                    "姓名": self.generate_name(),
                    "班级": self.generate_class()
                })
            
            # 为每个科目生成成绩
            subject_scores = {}
            for subj in selected_subjects:
                subject_scores[subj["name"]] = self.generate_scores_for_subject(subj, count)
            
            # 合并基本信息和成绩
            for i in range(count):
                row = basic_info[i].copy()
                for subj in selected_subjects:
                    row[subj["name"]] = subject_scores[subj["name"]][i]
                data.append(row)
            
            # 转换为DataFrame
            self.generated_data = pd.DataFrame(data)
            
            # 更新表格显示
            self.update_result_table()
            
            messagebox.showinfo("成功", f"已成功生成 {count} 条学生数据")
            
        except Exception as e:
            messagebox.showerror("错误", f"生成数据失败: {str(e)}")
    
    def update_result_table(self):
        """更新结果表格"""
        # 清空现有数据
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        
        if self.generated_data is None:
            return
        
        # 设置列
        columns = list(self.generated_data.columns)
        self.result_tree["columns"] = columns
        self.result_tree["show"] = "headings"
        
        # 设置列名和宽度
        for col in columns:
            self.result_tree.heading(col, text=col)
            width = 80 if col in ["姓名", "班级"] else 60
            self.result_tree.column(col, width=width, anchor=tk.CENTER)
        
        # 填充数据（只显示前50行，避免表格过大）
        display_data = self.generated_data.head(50)
        for _, row in display_data.iterrows():
            self.result_tree.insert("", tk.END, values=list(row))
        
        # 如果有更多数据，显示提示
        if len(self.generated_data) > 50:
            self.result_tree.insert("", tk.END, values=["...", "仅显示前50行", "...", "..."] + ["..." for _ in range(len(columns)-4)])
    
    def export_data(self):
        """导出数据"""
        if self.generated_data is None:
            messagebox.showwarning("警告", "请先生成数据")
            return
        
        # 获取保存路径
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        default_filename = f"学生成绩数据_{timestamp}.xlsx"
        
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile=default_filename,
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
        
        try:
            if file_path.endswith('.csv'):
                self.generated_data.to_csv(file_path, index=False, encoding='utf-8-sig')
            else:
                self.generated_data.to_excel(file_path, index=False)
            
            messagebox.showinfo("成功", f"数据已成功导出到:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("错误", f"导出数据失败: {str(e)}")
    
    def reset_settings(self):
        """重置设置为默认值"""
        # 重置学生数量
        self.student_count.set(50)
        self.id_prefix.set("2023")
        self.class_range.set("随机分配")
        
        # 重置科目设置
        for subject in SUBJECTS:
            name = subject["name"]
            self.subject_vars[name].set(True)
            self.subject_min[name].set(subject["min"])
            self.subject_max[name].set(subject["max"])
            self.subject_pass[name].set(subject["pass_score"])
            self.subject_pass_rate[name].set(subject["pass_rate"])
        
        # 清空结果
        self.generated_data = None
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentDataGenerator(root)
    root.mainloop()
    