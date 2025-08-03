import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.font_manager as fm
import os
import tempfile
from datetime import datetime
import seaborn as sns
import random


# 设置中文字体支持
# 优先使用本地目录的 SimSun.ttf 文件
if os.path.exists("SimSun.ttf"):
    font_prop = fm.FontProperties(fname="SimSun.ttf")
    plt.rcParams["font.family"] = font_prop.get_name()
    print(f"使用本地字体文件: SimSun.ttf")
else:
    # 查找系统中的中文字体
    font_path = None
    for font in fm.findSystemFonts():
        if 'simhei' in font.lower() or 'microsoft yahei' in font.lower() or 'simsun' in font.lower():
            font_path = font
            break

    if font_path:
        plt.rcParams["font.family"] = fm.FontProperties(fname=font_path).get_name()
        print(f"使用系统字体: {font_path}")
    else:
        print("未找到中文字体，可能导致中文显示异常")
    
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题

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
    """学生数据生成器类"""
    def __init__(self, parent):
        self.parent = parent
        self.window = None
        
        # 存储科目设置的变量
        self.subject_vars = {}
        self.subject_min = {}
        self.subject_max = {}
        self.subject_pass = {}
        self.subject_pass_rate = {}
        
        # 存储生成的数据
        self.generated_data = None
        
    def show_generator_window(self):
        """显示数据生成窗口"""
        if self.window is not None:
            self.window.lift()
            return
            
        self.window = tk.Toplevel(self.parent)
        self.window.title("学生信息与成绩随机生成工具")
        self.window.geometry("900x700")
        self.window.minsize(800, 600)
        self.window.transient(self.parent)
        
        # 当窗口关闭时清理引用
        self.window.protocol("WM_DELETE_WINDOW", self.on_window_close)
        
        self.create_widgets()
        
    def on_window_close(self):
        """窗口关闭处理"""
        self.window.destroy()
        self.window = None
        
    def create_widgets(self):
        # 创建主框架
        main_frame = ttk.Frame(self.window, padding="10")
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


class StudentGradeAnalysisSystem:
    def __init__(self, root):
        self.root = root
        self.root.title("学生成绩分析系统")
        self.root.geometry("1200x800")
        self.root.minsize(1000, 700)
        
        # 数据存储
        self.data = None
        self.analysis_results = None
        self.current_file = None
        
        # 创建数据生成器实例
        self.data_generator = StudentDataGenerator(self.root)
        
        # 创建主界面
        self.create_widgets()
        
    def create_widgets(self):
        # 创建菜单栏
        self.create_menu()
        
        # 先创建分析按钮区域（放在顶部）
        analysis_frame = ttk.Frame(self.root, padding="10")
        analysis_frame.pack(fill=tk.X, side=tk.BOTTOM)
        
        # 分析按钮
        ttk.Button(
            analysis_frame, 
            text="基本统计分析", 
            command=self.perform_basic_analysis
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            analysis_frame, 
            text="科目对比分析", 
            command=self.perform_subject_analysis
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            analysis_frame, 
            text="成绩分布分析", 
            command=self.perform_distribution_analysis
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            analysis_frame, 
            text="高级分析", 
            command=self.perform_advanced_analysis
        ).pack(side=tk.LEFT, padx=5)
        
        ttk.Button(
            analysis_frame, 
            text="导出分析结果", 
            command=self.export_analysis_results
        ).pack(side=tk.RIGHT, padx=5)
        
        # 添加生成PDF报告按钮
        ttk.Button(
            analysis_frame, 
            text="生成PDF报告", 
            command=self.generate_pdf_report
        ).pack(side=tk.RIGHT, padx=5)
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建左侧面板（数据展示）- 固定宽度
        left_frame = ttk.LabelFrame(main_frame, text="成绩数据", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        left_frame.config(width=600)  # 设置固定宽度为600像素
        left_frame.pack_propagate(False)  # 防止子组件改变父组件大小
        
        # 创建数据表格
        table_frame = ttk.Frame(left_frame)
        table_frame.pack(fill=tk.BOTH, expand=True)
        
        # 滚动条
        scrollbar_x = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL)
        scrollbar_y = ttk.Scrollbar(table_frame, orient=tk.VERTICAL)
        
        # 树状表格
        self.tree = ttk.Treeview(
            table_frame,
            yscrollcommand=scrollbar_y.set,
            xscrollcommand=scrollbar_x.set
        )
        scrollbar_x.config(command=self.tree.xview)
        scrollbar_y.config(command=self.tree.yview)
        
        scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # 创建右侧面板（分析结果）- 占据剩余空间
        right_frame = ttk.LabelFrame(main_frame, text="分析结果", padding="10")
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(10, 0))
        
        # 分析选项卡
        self.notebook = ttk.Notebook(right_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # 统计分析选项卡
        self.stats_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.stats_frame, text="统计分析")
        
        # 可视化选项卡
        self.visual_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.visual_frame, text="数据可视化")
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="导入数据", command=self.import_data)
        file_menu.add_command(label="保存数据", command=self.save_data)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="文件", menu=file_menu)
        
        # 工具菜单
        tools_menu = tk.Menu(menubar, tearoff=0)
        tools_menu.add_command(label="学生数据生成器", command=self.open_data_generator)
        menubar.add_cascade(label="工具", menu=tools_menu)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)
        help_menu.add_command(label="使用说明", command=self.show_help)
        menubar.add_cascade(label="帮助", menu=help_menu)
        
        self.root.config(menu=menubar)
    
    def open_data_generator(self):
        """打开数据生成器窗口"""
        self.data_generator.show_generator_window()
    
    def import_data(self):
        """导入数据文件"""
        file_path = filedialog.askopenfilename(
            filetypes=[
                ("Excel files", "*.xlsx;*.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        
        if not file_path:
            return
            
        try:
            # 根据文件扩展名选择合适的读取方法
            if file_path.endswith('.csv'):
                self.data = pd.read_csv(file_path)
            else:  # Excel文件
                self.data = pd.read_excel(file_path)
            
            self.current_file = file_path
            self.update_table()
            messagebox.showinfo("成功", f"数据导入成功，共 {len(self.data)} 条记录")
            
        except Exception as e:
            messagebox.showerror("错误", f"导入数据失败: {str(e)}")
    
    def save_data(self):
        """保存数据"""
        if self.data is None:
            messagebox.showwarning("警告", "没有数据可保存")
            return
            
        if self.current_file:
            try:
                if self.current_file.endswith('.csv'):
                    self.data.to_csv(self.current_file, index=False)
                else:
                    self.data.to_excel(self.current_file, index=False)
                messagebox.showinfo("成功", "数据保存成功")
            except Exception as e:
                messagebox.showerror("错误", f"保存数据失败: {str(e)}")
        else:
            self.export_data()
    
    def export_data(self):
        """导出数据"""
        if self.data is None:
            messagebox.showwarning("警告", "没有数据可导出")
            return
            
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
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
                self.data.to_csv(file_path, index=False)
            else:
                self.data.to_excel(file_path, index=False)
            self.current_file = file_path
            messagebox.showinfo("成功", "数据导出成功")
        except Exception as e:
            messagebox.showerror("错误", f"导出数据失败: {str(e)}")
    
    def update_table(self):
        """更新表格数据"""
        # 清空现有数据
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # 设置列
        self.tree["columns"] = list(self.data.columns)
        self.tree["show"] = "headings"
        
        # 设置列名
        for col in self.data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor=tk.CENTER)
        
        # 填充数据
        for idx, row in self.data.iterrows():
            self.tree.insert("", tk.END, values=list(row))
    
    def perform_basic_analysis(self):
        """执行基本统计分析"""
        if self.data is None:
            messagebox.showwarning("警告", "请先导入数据")
            return
            
        # 清空统计分析面板
        for widget in self.stats_frame.winfo_children():
            widget.destroy()
        
        # 创建滚动文本框显示统计结果
        text_frame = ttk.Frame(self.stats_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        stats_text = tk.Text(text_frame, wrap=tk.WORD, yscrollcommand=scrollbar.set, padx=10, pady=10)
        stats_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=stats_text.yview)
        
        # 基本统计信息
        stats_text.insert(tk.END, "===== 基本统计信息 =====\n\n")
        stats_text.insert(tk.END, f"总记录数: {len(self.data)}\n")
        
        # 假设存在"姓名"列
        if "姓名" in self.data.columns:
            stats_text.insert(tk.END, f"学生人数: {self.data['姓名'].nunique()}\n\n")
        
        # 识别科目列（假设非ID和姓名的列都是科目）
        subject_columns = []
        for col in self.data.columns:
            if col not in ["序号", "学号", "姓名", "班级", "总分", "排名"]:
                subject_columns.append(col)
        
        # 计算总分（如果不存在）
        if "总分" not in self.data.columns and subject_columns:
            self.data["总分"] = self.data[subject_columns].sum(axis=1)
        
        # 计算平均分（如果不存在）
        if "平均分" not in self.data.columns and subject_columns:
            self.data["平均分"] = self.data[subject_columns].mean(axis=1)
        
        # 计算排名（如果不存在）
        if "排名" not in self.data.columns and "总分" in self.data.columns:
            self.data["排名"] = self.data["总分"].rank(ascending=False, method="min").astype(int)
        
        # 更新表格
        self.update_table()
        
        # 科目统计
        stats_text.insert(tk.END, "===== 各科目统计 =====\n\n")
        for subject in subject_columns:
            stats_text.insert(tk.END, f"{subject}:\n")
            stats_text.insert(tk.END, f"  平均分: {self.data[subject].mean():.2f}\n")
            stats_text.insert(tk.END, f"  最高分: {self.data[subject].max()}\n")
            stats_text.insert(tk.END, f"  最低分: {self.data[subject].min()}\n")
            stats_text.insert(tk.END, f"  及格率: {((self.data[subject] >= 60).sum() / len(self.data) * 100):.2f}%\n")
            stats_text.insert(tk.END, f"  优秀率(>=90): {((self.data[subject] >= 90).sum() / len(self.data) * 100):.2f}%\n\n")
        
        # 总分统计
        if "总分" in self.data.columns:
            stats_text.insert(tk.END, "===== 总分统计 =====\n\n")
            stats_text.insert(tk.END, f"  平均分: {self.data['总分'].mean():.2f}\n")
            stats_text.insert(tk.END, f"  最高分: {self.data['总分'].max()}\n")
            stats_text.insert(tk.END, f"  最低分: {self.data['总分'].min()}\n\n")
        
        # 添加更多统计功能
        stats_text.insert(tk.END, "===== 各科目详细统计 =====\n\n")
        for subject in subject_columns:
            stats_text.insert(tk.END, f"{subject}:")
            stats_text.insert(tk.END, f"\n  平均分: {self.data[subject].mean():.2f}")
            stats_text.insert(tk.END, f"\n  中位数: {self.data[subject].median():.2f}")
            stats_text.insert(tk.END, f"\n  标准差: {self.data[subject].std():.2f}")
            stats_text.insert(tk.END, f"\n  最小值: {self.data[subject].min()}")
            stats_text.insert(tk.END, f"\n  最大值: {self.data[subject].max()}")
            stats_text.insert(tk.END, f"\n  25%分位数: {self.data[subject].quantile(0.25):.2f}")
            stats_text.insert(tk.END, f"\n  75%分位数: {self.data[subject].quantile(0.75):.2f}\n\n")
        
        # 保存分析结果
        self.analysis_results = {
            "basic_stats": self.data.describe(),
            "subject_columns": subject_columns
        }
        
        # 禁止编辑文本框
        stats_text.config(state=tk.DISABLED)
        
        # 切换到统计分析选项卡
        self.notebook.select(self.stats_frame)
    
    def perform_subject_analysis(self):
        """执行科目对比分析"""
        if self.data is None:
            messagebox.showwarning("警告", "请先导入数据")
            return
            
        # 清空可视化面板
        for widget in self.visual_frame.winfo_children():
            widget.destroy()
        
        # 识别科目列
        subject_columns = []
        for col in self.data.columns:
            if col not in ["序号", "学号", "姓名", "班级", "总分", "排名"]:
                subject_columns.append(col)
        
        if not subject_columns:
            messagebox.showwarning("警告", "未识别到科目列")
            return
        
        # 创建一个主框架用于放置所有图表
        main_canvas = tk.Canvas(self.visual_frame)
        scrollbar_v = ttk.Scrollbar(self.visual_frame, orient="vertical", command=main_canvas.yview)
        scrollbar_h = ttk.Scrollbar(self.visual_frame, orient="horizontal", command=main_canvas.xview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

        # 调整横向滚动条的位置，使其位于数据可视化界面范围的底部
        main_canvas.pack(side="top", fill="both", expand=True)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 存储所有图表以便导出
        self.current_figures = []
        self.current_analysis_type = 'subject'  # 标记分析类型
        
        # 1. 各科目平均分对比图
        avg_frame = ttk.LabelFrame(scrollable_frame, text="各科目平均分对比", padding="10")
        avg_frame.pack(fill=tk.X, padx=10, pady=5)
        
        fig1, ax1 = plt.subplots(1, 1, figsize=(12, 6))
        avg_scores = [self.data[subject].mean() for subject in subject_columns]
        
        # 使用更美观的颜色
        colors = plt.cm.Set3(range(len(subject_columns)))
        bars = ax1.bar(subject_columns, avg_scores, color=colors)
        
        # 添加数值标签
        for bar, score in zip(bars, avg_scores):
            height = bar.get_height()
            ax1.text(bar.get_x() + bar.get_width()/2., height + 1,
                    f'{score:.1f}', ha='center', va='bottom', fontweight='bold')
        
        ax1.set_title('各科目平均分对比', fontsize=14, fontweight='bold')
        ax1.set_ylabel('平均分')
        ax1.grid(axis='y', linestyle='--', alpha=0.7)
        
        # 自动调整x轴标签角度
        if len(subject_columns) > 5:
            plt.setp(ax1.get_xticklabels(), rotation=45, ha='right')
        
        plt.tight_layout()
        
        # 将图表嵌入框架
        canvas1 = FigureCanvasTkAgg(fig1, master=avg_frame)
        canvas1.draw()
        canvas1.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figures.append(fig1)
        
        # 2. 各科目分数分布箱线图
        box_frame = ttk.LabelFrame(scrollable_frame, text="各科目分数分布箱线图", padding="10")
        box_frame.pack(fill=tk.X, padx=10, pady=5)
        
        fig2, ax2 = plt.subplots(1, 1, figsize=(12, 6))
        
        # 准备箱线图数据
        box_data = [self.data[subject].dropna() for subject in subject_columns]
        
        # 创建箱线图
        box_plot = ax2.boxplot(box_data, labels=subject_columns, patch_artist=True)
        
        # 美化箱线图
        colors2 = plt.cm.Pastel1(range(len(subject_columns)))
        for patch, color in zip(box_plot['boxes'], colors2):
            patch.set_facecolor(color)
            patch.set_alpha(0.8)
        
        # 设置中位数线的颜色
        for median in box_plot['medians']:
            median.set_color('red')
            median.set_linewidth(2)
        
        ax2.set_title('各科目分数分布箱线图', fontsize=14, fontweight='bold')
        ax2.set_ylabel('分数')
        ax2.grid(axis='y', linestyle='--', alpha=0.7)
        
        # 自动调整x轴标签角度
        if len(subject_columns) > 5:
            plt.setp(ax2.get_xticklabels(), rotation=45, ha='right')
        
        plt.tight_layout()
        
        # 将图表嵌入框架
        canvas2 = FigureCanvasTkAgg(fig2, master=box_frame)
        canvas2.draw()
        canvas2.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figures.append(fig2)
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 保存第一个图表供导出（为了兼容性）
        if self.current_figures:
            self.current_fig = self.current_figures[0]
        
        # 切换到可视化选项卡
        self.notebook.select(self.visual_frame)
    
    def perform_distribution_analysis(self):
        """执行成绩分布分析"""
        if self.data is None:
            messagebox.showwarning("警告", "请先导入数据")
            return
            
        # 清空可视化面板
        for widget in self.visual_frame.winfo_children():
            widget.destroy()
        
        # 识别科目列
        subject_columns = []
        for col in self.data.columns:
            if col not in ["序号", "学号", "姓名", "班级", "总分", "排名"]:
                subject_columns.append(col)
        
        if not subject_columns:
            messagebox.showwarning("警告", "未识别到科目列")
            return
        
        # 创建一个主框架用于放置所有图表
        main_canvas = tk.Canvas(self.visual_frame)
        scrollbar_v = ttk.Scrollbar(self.visual_frame, orient="vertical", command=main_canvas.yview)
        scrollbar_h = ttk.Scrollbar(self.visual_frame, orient="horizontal", command=main_canvas.xview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

        # 调整横向滚动条的位置，使其位于数据可视化界面范围的底部
        main_canvas.pack(side="top", fill="both", expand=True)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 存储所有图表以便导出
        self.current_figures = []
        self.current_analysis_type = 'distribution'  # 标记分析类型
        
        # 1. 如果有总分，先显示总分分布
        if "总分" in self.data.columns:
            total_frame = ttk.LabelFrame(scrollable_frame, text="总分分布分析", padding="10")
            total_frame.pack(fill=tk.X, padx=10, pady=5)
            
            fig_total, ax_total = plt.subplots(1, 1, figsize=(12, 6))
            sns.histplot(self.data["总分"], bins=15, kde=True, ax=ax_total, color='green')
            ax_total.set_title('总分分布直方图', fontsize=14, fontweight='bold')
            ax_total.set_xlabel('总分')
            ax_total.set_ylabel('学生数量')
            ax_total.grid(axis='y', linestyle='--', alpha=0.7)
            
            # 添加统计信息
            mean_score = self.data["总分"].mean()
            max_score = self.data["总分"].max()
            min_score = self.data["总分"].min()
            ax_total.axvline(mean_score, color='red', linestyle='--', linewidth=2, label=f'平均分: {mean_score:.1f}')
            ax_total.legend()
            
            plt.tight_layout()
            
            canvas_total = FigureCanvasTkAgg(fig_total, master=total_frame)
            canvas_total.draw()
            canvas_total.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig_total)
        
        # 2. 为每个科目创建独立的分析图表
        colors = sns.color_palette('Set2', n_colors=len(subject_columns))
        
        for i, subject in enumerate(subject_columns):
            # 为每个科目创建独立的框架
            subject_frame = ttk.LabelFrame(scrollable_frame, text=f"{subject} 分析", padding="10")
            subject_frame.pack(fill=tk.X, padx=10, pady=5)
            
            # 创建包含两个子图的图表：直方图和饼图
            fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
            
            # 左侧：分数分布直方图
            sns.histplot(self.data[subject], bins=12, kde=True, ax=ax1, color=colors[i])
            ax1.set_title(f'{subject} 分数分布直方图', fontsize=12, fontweight='bold')
            ax1.set_xlabel('分数')
            ax1.set_ylabel('学生数量')
            ax1.grid(axis='y', linestyle='--', alpha=0.7)
            
            # 添加统计信息到直方图
            mean_score = self.data[subject].mean()
            pass_score = 60 if self.data[subject].max() <= 100 else 90
            ax1.axvline(mean_score, color='red', linestyle='--', linewidth=2, label=f'平均分: {mean_score:.1f}')
            ax1.axvline(pass_score, color='orange', linestyle='--', linewidth=2, label=f'及格线: {pass_score}')
            ax1.legend()
            
            # 右侧：分数段占比饼图
            subject_max = self.data[subject].max()
            if subject_max <= 100:
                bins = [0, 60, 70, 80, 90, 100]
                labels = ['不及格\n(0-59)', '及格\n(60-69)', '中等\n(70-79)', '良好\n(80-89)', '优秀\n(90-100)']
            elif subject_max <= 150:
                bins = [0, 90, 105, 120, 135, 150]
                labels = ['不及格\n(0-89)', '及格\n(90-104)', '中等\n(105-119)', '良好\n(120-134)', '优秀\n(135-150)']
            else:
                # 对于其他满分制，使用百分比分段
                bins = [0, subject_max*0.6, subject_max*0.7, subject_max*0.8, subject_max*0.9, subject_max]
                labels = ['不及格', '及格', '中等', '良好', '优秀']
            
            score_ranges = pd.cut(self.data[subject], bins=bins, labels=labels, include_lowest=True)
            range_counts = score_ranges.value_counts(normalize=True).reindex(labels) * 100
            
            # 过滤掉为0的分段
            non_zero_counts = range_counts[range_counts > 0]
            non_zero_labels = non_zero_counts.index.tolist()
            
            # 使用更美观的颜色
            pie_colors = sns.color_palette('viridis', n_colors=len(non_zero_labels))
            
            wedges, texts, autotexts = ax2.pie(non_zero_counts, labels=non_zero_labels, autopct='%1.1f%%', 
                                             startangle=90, colors=pie_colors, explode=[0.05]*len(non_zero_labels))
            
            # 美化饼图文本
            for autotext in autotexts:
                autotext.set_color('white')
                autotext.set_fontweight('bold')
                autotext.set_fontsize(10)
            
            ax2.set_title(f'{subject} 各分数段占比', fontsize=12, fontweight='bold')
            
            plt.tight_layout()
            
            # 将图表嵌入框架
            canvas = FigureCanvasTkAgg(fig, master=subject_frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig)
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 保存第一个图表供导出（为了兼容性）
        if self.current_figures:
            self.current_fig = self.current_figures[0]
        
        # 切换到可视化选项卡
        self.notebook.select(self.visual_frame)
    
    def export_analysis_results(self):
        """导出分析结果"""
        if self.data is None or self.analysis_results is None:
            messagebox.showwarning("警告", "请先进行分析")
            return
            
        # 创建导出对话框
        export_window = tk.Toplevel(self.root)
        export_window.title("导出分析结果")
        export_window.geometry("400x300")
        export_window.transient(self.root)
        export_window.grab_set()
        
        # 导出选项
        ttk.Label(export_window, text="请选择要导出的内容:").pack(pady=10)
        
        var_stats = tk.BooleanVar(value=True)
        var_chart = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(export_window, text="统计数据", variable=var_stats).pack(anchor=tk.W, padx=20, pady=5)
        ttk.Checkbutton(export_window, text="图表", variable=var_chart).pack(anchor=tk.W, padx=20, pady=5)
        
        # 格式选择
        ttk.Label(export_window, text="请选择导出格式:").pack(pady=10)
        
        frame_format = ttk.Frame(export_window)
        frame_format.pack(pady=5)
        
        var_format = tk.StringVar(value="excel")
        ttk.Radiobutton(frame_format, text="Excel", variable=var_format, value="excel").pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(frame_format, text="CSV", variable=var_format, value="csv").pack(side=tk.LEFT, padx=10)
        
        # 导出按钮
        def do_export():
            if not var_stats.get() and not var_chart.get():
                messagebox.showwarning("警告", "请至少选择一项导出内容")
                return
                
            # 获取保存路径
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"成绩分析结果_{timestamp}"
            
            if var_stats.get():
                file_ext = ".xlsx" if var_format.get() == "excel" else ".csv"
                file_path = filedialog.asksaveasfilename(
                    defaultextension=file_ext,
                    initialfile=default_filename,
                    filetypes=[
                        ("Excel files", "*.xlsx"),
                        ("CSV files", "*.csv"),
                        ("All files", "*.*")
                    ]
                )
                
                if file_path:
                    try:
                        # 创建结果表格
                        result_df = pd.DataFrame()
                        
                        # 添加基本统计
                        stats_df = self.analysis_results["basic_stats"]
                        
                        # 添加总分排名前10的学生
                        if "总分" in self.data.columns:
                            top_students = self.data.sort_values("总分", ascending=False).head(10)
                            result_df = pd.concat([result_df, pd.DataFrame(["", "总分排名前10的学生", ""], columns=["备注"])])
                            result_df = pd.concat([result_df, top_students[["姓名", "总分", "排名"]]])
                        
                        # 保存到文件
                        if file_path.endswith('.csv'):
                            result_df.to_csv(file_path, index=False)
                        else:
                            with pd.ExcelWriter(file_path) as writer:
                                stats_df.to_excel(writer, sheet_name='统计数据')
                                if "总分" in self.data.columns:
                                    top_students.to_excel(writer, sheet_name='前十名学生', index=False)
                
                    except Exception as e:
                        messagebox.showerror("错误", f"导出统计数据失败: {str(e)}")
                        return
            
            # 导出图表
            if var_chart.get() and hasattr(self, 'current_figures'):
                # 如果有多个图表，询问用户是否要分别导出
                if len(self.current_figures) > 1:
                    export_choice = messagebox.askyesnocancel(
                        "图表导出", 
                        "检测到多个图表。\n点击'是'分别导出每个图表\n点击'否'导出为单个文件\n点击'取消'跳过图表导出"
                    )
                    
                    if export_choice is None:  # 取消
                        pass
                    elif export_choice:  # 分别导出
                        base_path = filedialog.askdirectory(title="选择保存目录")
                        if base_path:
                            try:
                                for i, fig in enumerate(self.current_figures):
                                    if i == 0 and "总分" in self.data.columns:
                                        filename = f"{default_filename}_总分分布.png"
                                    elif hasattr(self, 'current_analysis_type'):
                                        # 根据分析类型命名
                                        if self.current_analysis_type == 'subject':
                                            if i == 0 or (i == 1 and "总分" not in self.data.columns):
                                                filename = f"{default_filename}_各科目平均分对比.png"
                                            elif i == 1 or (i == 2 and "总分" in self.data.columns):
                                                filename = f"{default_filename}_各科目分数分布箱线图.png"
                                            else:
                                                filename = f"{default_filename}_科目对比图表{i+1}.png"
                                        else:  # distribution analysis
                                            # 获取科目名
                                            subject_columns = [col for col in self.data.columns 
                                                             if col not in ["序号", "学号", "姓名", "班级", "总分", "排名"]]
                                            subject_idx = i - (1 if "总分" in self.data.columns else 0)
                                            if subject_idx < len(subject_columns):
                                                subject_name = subject_columns[subject_idx]
                                                filename = f"{default_filename}_{subject_name}分析.png"
                                            else:
                                                filename = f"{default_filename}_图表{i+1}.png"
                                    else:
                                        # 获取科目名（兼容性处理）
                                        subject_columns = [col for col in self.data.columns 
                                                         if col not in ["序号", "学号", "姓名", "班级", "总分", "排名"]]
                                        subject_idx = i - (1 if "总分" in self.data.columns else 0)
                                        if subject_idx < len(subject_columns):
                                            subject_name = subject_columns[subject_idx]
                                            filename = f"{default_filename}_{subject_name}分析.png"
                                        else:
                                            filename = f"{default_filename}_图表{i+1}.png"
                                    
                                    file_path = os.path.join(base_path, filename)
                                    fig.savefig(file_path, dpi=300, bbox_inches='tight')
                                
                                messagebox.showinfo("成功", f"已导出 {len(self.current_figures)} 个图表到:\n{base_path}")
                            except Exception as e:
                                messagebox.showerror("错误", f"导出图表失败: {str(e)}")
                                return
                    else:  # 导出为单个文件
                        file_path = filedialog.asksaveasfilename(
                            defaultextension=".png",
                            initialfile=default_filename,
                            filetypes=[
                                ("PNG files", "*.png"),
                                ("JPEG files", "*.jpg"),
                                ("PDF files", "*.pdf"),
                                ("All files", "*.*")
                            ]
                        )
                        
                        if file_path:
                            try:
                                # 将所有图表合并为一个大图
                                from matplotlib.backends.backend_pdf import PdfPages
                                
                                if file_path.endswith('.pdf'):
                                    with PdfPages(file_path) as pdf:
                                        for fig in self.current_figures:
                                            pdf.savefig(fig, bbox_inches='tight')
                                else:
                                    # 对于图片格式，只保存第一个图表
                                    self.current_figures[0].savefig(file_path, dpi=300, bbox_inches='tight')
                                    
                            except Exception as e:
                                messagebox.showerror("错误", f"导出图表失败: {str(e)}")
                                return
                else:
                    # 只有一个图表的情况
                    file_path = filedialog.asksaveasfilename(
                        defaultextension=".png",
                        initialfile=default_filename,
                        filetypes=[
                            ("PNG files", "*.png"),
                            ("JPEG files", "*.jpg"),
                            ("PDF files", "*.pdf"),
                            ("All files", "*.*")
                        ]
                    )
                    
                    if file_path:
                        try:
                            self.current_figures[0].savefig(file_path, dpi=300, bbox_inches='tight')
                        except Exception as e:
                            messagebox.showerror("错误", f"导出图表失败: {str(e)}")
                            return
            elif var_chart.get() and hasattr(self, 'current_fig'):
                # 兼容旧版本的单图表导出
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".png",
                    initialfile=default_filename,
                    filetypes=[
                        ("PNG files", "*.png"),
                        ("JPEG files", "*.jpg"),
                        ("PDF files", "*.pdf"),
                        ("All files", "*.*")
                    ]
                )
                
                if file_path:
                    try:
                        self.current_fig.savefig(file_path, dpi=300, bbox_inches='tight')
                    except Exception as e:
                        messagebox.showerror("错误", f"导出图表失败: {str(e)}")
                        return
        
            messagebox.showinfo("成功", "分析结果导出成功")
            export_window.destroy()
    
    def generate_pdf_report(self):
        """生成PDF报告"""
        if self.data is None or self.analysis_results is None:
            messagebox.showwarning("警告", "请先进行分析")
            return

        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        from matplotlib.backends.backend_agg import FigureCanvasAgg
        import os
        import tempfile
        import matplotlib.font_manager as fm

        # 字体设置辅助函数
        def set_pdf_font(pdf, size=12, bold=False):
            """设置PDF字体"""
            if font_registered:
                pdf.setFont("ChineseFont", size)
            else:
                if bold:
                    pdf.setFont("Helvetica-Bold", size)
                else:
                    pdf.setFont("Helvetica", size)

        # 注册中文字体
        font_registered = False
        try:
            # 优先使用本地目录的 SimSun.ttf 文件
            if os.path.exists("SimSun.ttf"):
                pdfmetrics.registerFont(TTFont("ChineseFont", "SimSun.ttf"))
                font_registered = True
                print("成功注册本地中文字体: SimSun.ttf")
                # 测试字体是否真的可用
                try:
                    test_canvas = canvas.Canvas("test.pdf")
                    test_canvas.setFont("ChineseFont", 12)
                    test_canvas.drawString(50, 750, "测试")
                    print("中文字体测试成功")
                except Exception as e:
                    print(f"中文字体测试失败: {e}")
                    font_registered = False
            else:
                # 尝试注册系统中的中文字体
                chinese_fonts = []
                for font in fm.findSystemFonts():
                    if any(name in font.lower() for name in ['simhei', 'simsun', 'microsoft yahei', 'pingfang', 'noto']):
                        chinese_fonts.append(font)
                
                if chinese_fonts:
                    # 使用找到的第一个中文字体
                    font_path = chinese_fonts[0]
                    pdfmetrics.registerFont(TTFont("ChineseFont", font_path))
                    font_registered = True
                    print(f"成功注册系统中文字体: {font_path}")
                else:
                    print("未找到中文字体，使用默认字体")
        except Exception as e:
            print(f"字体注册失败: {e}")
            font_registered = False

        # 获取保存路径
        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*")]
        )

        if not file_path:
            return

        try:
            # 创建PDF画布
            pdf = canvas.Canvas(file_path, pagesize=letter)
            width, height = letter

            # 添加标题
            set_pdf_font(pdf, 16, bold=True)
            pdf.drawString(30, height - 30, "学生成绩分析报告")

            # 添加统计数据
            y_position = height - 60
            set_pdf_font(pdf, 12)
            pdf.drawString(30, y_position, "===== 基本统计信息 =====")
            y_position -= 20
            pdf.drawString(30, y_position, f"总记录数: {len(self.data)}")

            if "姓名" in self.data.columns:
                y_position -= 20
                pdf.drawString(30, y_position, f"学生人数: {self.data['姓名'].nunique()}")

            y_position -= 40
            pdf.drawString(30, y_position, "===== 各科目详细统计 =====")
            y_position -= 20

            for subject in self.analysis_results["subject_columns"]:
                if y_position < 50:
                    pdf.showPage()
                    y_position = height - 30

                pdf.drawString(30, y_position, f"{subject}:")
                y_position -= 20
                pdf.drawString(50, y_position, f"平均分: {self.data[subject].mean():.2f}")
                y_position -= 20
                pdf.drawString(50, y_position, f"中位数: {self.data[subject].median():.2f}")
                y_position -= 20
                pdf.drawString(50, y_position, f"标准差: {self.data[subject].std():.2f}")
                y_position -= 20

            # 添加更多统计内容
            if "总分" in self.data.columns:
                y_position -= 20
                pdf.drawString(30, y_position, "===== 总分统计 =====")
                y_position -= 20
                pdf.drawString(50, y_position, f"平均分: {self.data['总分'].mean():.2f}")
                y_position -= 20
                pdf.drawString(50, y_position, f"最高分: {self.data['总分'].max()}")
                y_position -= 20
                pdf.drawString(50, y_position, f"最低分: {self.data['总分'].min()}")
                y_position -= 20
                pdf.drawString(50, y_position, f"标准差: {self.data['总分'].std():.2f}")
                y_position -= 30
                
                # 添加排名前10的学生
                pdf.drawString(30, y_position, "===== 排名前10的学生 =====")
                y_position -= 20
                top_students = self.data.sort_values("总分", ascending=False).head(10)
                for idx, (_, student) in enumerate(top_students.iterrows()):
                    if y_position < 50:
                        pdf.showPage()
                        y_position = height - 30
                        set_pdf_font(pdf, 12)  # 重新设置字体
                    
                    # 处理中文字符编码问题
                    student_name = student['姓名']
                    total_score = student['总分']
                    
                    if font_registered:
                        # 如果成功注册了中文字体，尝试使用中文姓名
                        try:
                            # 确保使用正确的字体
                            set_pdf_font(pdf, 12)
                            student_info = f"{idx+1}. {student_name}: {total_score}"
                            pdf.drawString(50, y_position, student_info)
                        except Exception as e:
                            print(f"中文姓名显示失败: {e}")
                            # 备用方案：使用学号或序号
                            if '学号' in student:
                                student_info = f"{idx+1}. Student ID {student['学号']}: {total_score}"
                            else:
                                student_info = f"{idx+1}. Student {idx+1}: {total_score}"
                            pdf.drawString(50, y_position, student_info)
                    else:
                        # 如果没有中文字体，使用备用标识
                        if '学号' in student:
                            student_info = f"{idx+1}. Student ID {student['学号']}: {total_score}"
                        else:
                            student_info = f"{idx+1}. Student {idx+1}: {total_score}"
                        pdf.drawString(50, y_position, student_info)
                    
                    y_position -= 20

            # 添加所有图表
            if hasattr(self, 'current_figures'):
                for i, fig in enumerate(self.current_figures):
                    # 检查是否需要新页面
                    if y_position < 350:
                        pdf.showPage()
                        y_position = height - 30

                    # 添加图表标题
                    set_pdf_font(pdf, 14, bold=True)
                    if i == 0 and "总分" in self.data.columns:
                        chart_title = "总分分布分析图"
                    elif hasattr(self, 'current_analysis_type'):
                        if self.current_analysis_type == 'subject':
                            if i == 0 or (i == 1 and "总分" not in self.data.columns):
                                chart_title = "各科目平均分对比图"
                            elif i == 1 or (i == 2 and "总分" in self.data.columns):
                                chart_title = "各科目分数分布箱线图"
                            else:
                                chart_title = f"科目对比图表 {i+1}"
                        else:  # distribution analysis
                            subject_columns = [col for col in self.data.columns 
                                             if col not in ["序号", "学号", "姓名", "班级", "总分", "排名"]]
                            subject_idx = i - (1 if "总分" in self.data.columns else 0)
                            if subject_idx < len(subject_columns):
                                subject_name = subject_columns[subject_idx]
                                chart_title = f"{subject_name} 分析图"
                            else:
                                chart_title = f"分析图表 {i+1}"
                    else:
                        chart_title = f"分析图表 {i+1}"
                    
                    pdf.drawString(30, y_position, chart_title)
                    y_position -= 30

                    # 将图表保存为图像（使用唯一文件名）
                    import tempfile
                    import os
                    
                    # 创建临时文件
                    temp_file = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
                    temp_file.close()
                    img_path = temp_file.name
                    
                    try:
                        # 保存图表
                        fig.savefig(img_path, format="png", dpi=150, bbox_inches='tight')

                        # 在PDF中插入图像
                        pdf.drawImage(img_path, 30, y_position - 280, width=500, height=280)
                        y_position -= 320
                        
                    finally:
                        # 清理临时文件
                        if os.path.exists(img_path):
                            os.unlink(img_path)

            # 添加分析结论
            if y_position < 150:
                pdf.showPage()
                y_position = height - 30
            
            set_pdf_font(pdf, 14, bold=True)
            pdf.drawString(30, y_position, "===== 分析结论 =====")
            y_position -= 30
            set_pdf_font(pdf, 12)
            
            # 计算整体表现
            subject_columns = self.analysis_results["subject_columns"]
            overall_avg = sum(self.data[subject].mean() for subject in subject_columns) / len(subject_columns)
            pdf.drawString(50, y_position, f"1. 各科目平均分为: {overall_avg:.2f}")
            y_position -= 20
            
            # 找出表现最好和最差的科目
            best_subject = max(subject_columns, key=lambda x: self.data[x].mean())
            worst_subject = min(subject_columns, key=lambda x: self.data[x].mean())
            pdf.drawString(50, y_position, f"2. 表现最好的科目: {best_subject} ({self.data[best_subject].mean():.2f})")
            y_position -= 20
            pdf.drawString(50, y_position, f"3. 需要改进的科目: {worst_subject} ({self.data[worst_subject].mean():.2f})")
            y_position -= 20
            
            # 整体及格率
            overall_pass_rate = sum((self.data[subject] >= 60).sum() for subject in subject_columns) / (len(subject_columns) * len(self.data)) * 100
            pdf.drawString(50, y_position, f"4. 整体及格率: {overall_pass_rate:.2f}%")
            y_position -= 30
            
            # 添加生成时间
            from datetime import datetime
            pdf.drawString(50, y_position, f"报告生成时间: {datetime.now().strftime('%Y年%m月%d日 %H:%M:%S')}")

            pdf.save()
            messagebox.showinfo("成功", "PDF报告生成成功")

        except Exception as e:
            messagebox.showerror("错误", f"生成PDF报告失败: {str(e)}")
    
    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo(
            "关于",
            "学生成绩分析系统 v2.0\n\n"
            "一个用于分析学生成绩的工具，支持数据导入、统计分析、数据可视化和结果导出。\n\n"
            "新增功能:\n- 集成学生数据生成器\n- 可生成随机学生成绩数据\n\n"
            "© 2023 学生成绩分析系统"
        )
    
    def show_help(self):
        """显示使用说明"""
        help_window = tk.Toplevel(self.root)
        help_window.title("使用说明")
        help_window.geometry("600x500")
        help_window.transient(self.root)
        
        # 创建滚动文本框
        scrollbar = ttk.Scrollbar(help_window)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        help_text = tk.Text(help_window, wrap=tk.WORD, yscrollcommand=scrollbar.set, padx=10, pady=10)
        help_text.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=help_text.yview)
        
        # 帮助内容
        help_content = """
学生成绩分析系统使用说明

1. 数据导入
   - 点击菜单栏的"文件" -> "导入数据"
   - 支持导入Excel(.xlsx, .xls)和CSV(.csv)格式的文件
   - 数据应包含学生信息（如姓名、学号等）和各科目成绩

2. 数据生成
   - 点击菜单栏的"工具" -> "学生数据生成器"
   - 可以自定义生成学生数量、学号前缀、班级范围
   - 可以选择科目和设置分数范围、及格分、及格率
   - 生成的数据可以直接导出为Excel或CSV文件

3. 基本操作
   - 导入数据后，系统会在左侧表格显示数据
   - 可以使用底部的按钮进行各种分析

4. 分析功能
   - 基本统计分析：计算各科目平均分、最高分、最低分、及格率等
   - 科目对比分析：生成各科目平均分对比图和分数分布箱线图
   - 成绩分布分析：生成总分分布直方图和分数段占比饼图

5. 结果导出
   - 点击"导出分析结果"按钮
   - 可以选择导出统计数据和图表
   - 统计数据支持Excel和CSV格式，图表支持PNG、JPG和PDF格式

6. 数据保存
   - 点击菜单栏的"文件" -> "保存数据"可以保存当前数据
   - 点击"文件" -> "导出数据"可以将数据导出为新文件
        """
        
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)
    
    def perform_advanced_analysis(self):
        """执行高级分析"""
        if self.data is None:
            messagebox.showwarning("警告", "请先导入数据")
            return
            
        # 确保中文字体设置
        if os.path.exists("SimSun.ttf"):
            font_prop = fm.FontProperties(fname="SimSun.ttf")
            plt.rcParams["font.family"] = font_prop.get_name()
        
        # 清空可视化面板
        for widget in self.visual_frame.winfo_children():
            widget.destroy()
        
        # 识别科目列
        subject_columns = []
        for col in self.data.columns:
            if col not in ["序号", "学号", "姓名", "班级", "总分", "排名"]:
                subject_columns.append(col)
        
        if not subject_columns:
            messagebox.showwarning("警告", "未识别到科目列")
            return
        
        # 创建一个主框架用于放置所有图表
        main_canvas = tk.Canvas(self.visual_frame)
        scrollbar_v = ttk.Scrollbar(self.visual_frame, orient="vertical", command=main_canvas.yview)
        scrollbar_h = ttk.Scrollbar(self.visual_frame, orient="horizontal", command=main_canvas.xview)
        scrollable_frame = ttk.Frame(main_canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: main_canvas.configure(scrollregion=main_canvas.bbox("all"))
        )

        main_canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        main_canvas.configure(yscrollcommand=scrollbar_v.set, xscrollcommand=scrollbar_h.set)

        # 调整横向滚动条的位置，使其位于数据可视化界面范围的底部
        main_canvas.pack(side="top", fill="both", expand=True)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        
        # 存储所有图表以便导出
        self.current_figures = []
        self.current_analysis_type = 'advanced'  # 标记分析类型
        
        # 1. 科目相关性热力图
        corr_frame = ttk.LabelFrame(scrollable_frame, text="科目间相关性分析", padding="10")
        corr_frame.pack(fill=tk.X, padx=10, pady=5)
        
        fig1, ax1 = plt.subplots(1, 1, figsize=(12, 8))
        
        # 计算相关性矩阵
        correlation_matrix = self.data[subject_columns].corr()
        
        # 创建热力图
        sns.heatmap(correlation_matrix, annot=True, cmap='coolwarm', center=0,
                   square=True, ax=ax1, fmt='.2f', cbar_kws={'shrink': .8})
        ax1.set_title('科目间相关性热力图', fontsize=14, fontweight='bold')
        plt.tight_layout()
        
        canvas1 = FigureCanvasTkAgg(fig1, master=corr_frame)
        canvas1.draw()
        canvas1.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figures.append(fig1)
        
        # 2. 班级成绩对比分析（如果有班级信息）
        if "班级" in self.data.columns:
            class_frame = ttk.LabelFrame(scrollable_frame, text="班级成绩对比分析", padding="10")
            class_frame.pack(fill=tk.X, padx=10, pady=5)
            
            fig2, ax2 = plt.subplots(1, 1, figsize=(12, 6))
            
            # 计算各班级平均分
            class_avg = self.data.groupby("班级")[subject_columns].mean()
            
            # 创建堆叠柱状图
            class_avg.plot(kind='bar', ax=ax2, width=0.8)
            ax2.set_title('各班级各科目平均分对比', fontsize=14, fontweight='bold')
            ax2.set_ylabel('平均分')
            ax2.set_xlabel('班级')
            ax2.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
            plt.xticks(rotation=45)
            plt.tight_layout()
            
            canvas2 = FigureCanvasTkAgg(fig2, master=class_frame)
            canvas2.draw()
            canvas2.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig2)
        
        # 3. 成绩分布密度图
        density_frame = ttk.LabelFrame(scrollable_frame, text="成绩分布密度分析", padding="10")
        density_frame.pack(fill=tk.X, padx=10, pady=5)
        
        fig3, ax3 = plt.subplots(1, 1, figsize=(12, 6))
        
        # 为每个科目绘制密度曲线
        colors = plt.cm.Set3(range(len(subject_columns)))
        for i, subject in enumerate(subject_columns):
            sns.kdeplot(data=self.data, x=subject, ax=ax3, color=colors[i], label=subject)
        
        ax3.set_title('各科目成绩分布密度图', fontsize=14, fontweight='bold')
        ax3.set_xlabel('分数')
        ax3.set_ylabel('密度')
        ax3.legend()
        ax3.grid(True, alpha=0.3)
        plt.tight_layout()
        
        canvas3 = FigureCanvasTkAgg(fig3, master=density_frame)
        canvas3.draw()
        canvas3.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        self.current_figures.append(fig3)
        
        # 4. 学生成绩雷达图（前5名学生）
        if "总分" in self.data.columns:
            radar_frame = ttk.LabelFrame(scrollable_frame, text="优秀学生成绩雷达图", padding="10")
            radar_frame.pack(fill=tk.X, padx=10, pady=5)
            
            fig4, ax4 = plt.subplots(1, 1, figsize=(10, 10), subplot_kw=dict(projection='polar'))
            
            # 获取前5名学生
            top5_students = self.data.nlargest(5, "总分")
            
            # 设置雷达图的角度
            angles = np.linspace(0, 2 * np.pi, len(subject_columns), endpoint=False)
            angles = np.concatenate((angles, [angles[0]]))  # 闭合图形
            
            colors_radar = plt.cm.Set1(range(5))
            
            for i, (_, student) in enumerate(top5_students.iterrows()):
                values = [student[subject] for subject in subject_columns]
                values += [values[0]]  # 闭合图形
                
                ax4.plot(angles, values, 'o-', linewidth=2, label=f"{student['姓名']}", color=colors_radar[i])
                ax4.fill(angles, values, alpha=0.1, color=colors_radar[i])
            
            ax4.set_xticks(angles[:-1])
            ax4.set_xticklabels(subject_columns)
            ax4.set_title('前5名学生各科目成绩雷达图', fontsize=14, fontweight='bold', pad=20)
            ax4.legend(loc='upper right', bbox_to_anchor=(1.2, 1.0))
            plt.tight_layout()
            
            canvas4 = FigureCanvasTkAgg(fig4, master=radar_frame)
            canvas4.draw()
            canvas4.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig4)
        
        # 5. 成绩趋势分析（散点图矩阵）
        scatter_frame = ttk.LabelFrame(scrollable_frame, text="科目间成绩散点图矩阵", padding="10")
        scatter_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # 选择主要科目进行散点图分析（避免图表过于复杂）
        main_subjects = subject_columns[:4] if len(subject_columns) > 4 else subject_columns
        
        if len(main_subjects) > 1:
            fig5, axes = plt.subplots(len(main_subjects), len(main_subjects), figsize=(12, 12))
            
            for i, subject1 in enumerate(main_subjects):
                for j, subject2 in enumerate(main_subjects):
                    ax = axes[i, j] if len(main_subjects) > 1 else axes
                    
                    if i == j:
                        # 对角线显示直方图
                        ax.hist(self.data[subject1], bins=15, alpha=0.7, color='skyblue')
                        ax.set_title(f'{subject1} 分布')
                    else:
                        # 非对角线显示散点图
                        ax.scatter(self.data[subject2], self.data[subject1], alpha=0.6, s=20)
                        
                        # 添加趋势线
                        z = np.polyfit(self.data[subject2], self.data[subject1], 1)
                        p = np.poly1d(z)
                        ax.plot(self.data[subject2], p(self.data[subject2]), "r--", alpha=0.8)
                    
                    if j == 0:
                        ax.set_ylabel(subject1)
                    if i == len(main_subjects) - 1:
                        ax.set_xlabel(subject2)
            
            plt.suptitle('主要科目间成绩关系散点图矩阵', fontsize=14, fontweight='bold')
            plt.tight_layout()
            
            canvas5 = FigureCanvasTkAgg(fig5, master=scatter_frame)
            canvas5.draw()
            canvas5.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            self.current_figures.append(fig5)
        
        # 绑定鼠标滚轮事件
        def _on_mousewheel(event):
            main_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        main_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # 保存第一个图表供导出（为了兼容性）
        if self.current_figures:
            self.current_fig = self.current_figures[0]
        
        # 切换到可视化选项卡
        self.notebook.select(self.visual_frame)

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentGradeAnalysisSystem(root)
    root.mainloop()
