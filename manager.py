import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import os
from datetime import datetime
import seaborn as sns

# 设置中文字体支持
plt.rcParams["font.family"] = ["SimHei", "WenQuanYi Micro Hei", "Heiti TC"]
plt.rcParams["axes.unicode_minus"] = False  # 解决负号显示问题

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
        
        # 创建主界面
        self.create_widgets()
        
    def create_widgets(self):
        # 创建菜单栏
        self.create_menu()
        
        # 创建主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # 创建左侧面板（数据展示）
        left_frame = ttk.LabelFrame(main_frame, text="成绩数据", padding="10")
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
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
        
        # 创建右侧面板（分析结果）
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
        
        # 创建分析按钮区域
        analysis_frame = ttk.Frame(self.root, padding="10")
        analysis_frame.pack(fill=tk.X)
        
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
            text="导出分析结果", 
            command=self.export_analysis_results
        ).pack(side=tk.RIGHT, padx=5)
    
    def create_menu(self):
        menubar = tk.Menu(self.root)
        
        # 文件菜单
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="导入数据", command=self.import_data)
        file_menu.add_command(label="保存数据", command=self.save_data)
        file_menu.add_separator()
        file_menu.add_command(label="退出", command=self.root.quit)
        menubar.add_cascade(label="文件", menu=file_menu)
        
        # 帮助菜单
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="关于", command=self.show_about)
        help_menu.add_command(label="使用说明", command=self.show_help)
        menubar.add_cascade(label="帮助", menu=help_menu)
        
        self.root.config(menu=menubar)
    
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
        
        # 创建图表
        fig, axes = plt.subplots(2, 1, figsize=(10, 10))
        fig.subplots_adjust(hspace=0.5)
        
        # 1. 各科目平均分对比
        avg_scores = [self.data[subject].mean() for subject in subject_columns]
        axes[0].bar(subject_columns, avg_scores, color='skyblue')
        axes[0].set_title('各科目平均分对比')
        axes[0].set_ylabel('平均分')
        axes[0].grid(axis='y', linestyle='--', alpha=0.7)
        
        # 2. 各科目分数分布箱线图
        self.data[subject_columns].boxplot(ax=axes[1], patch_artist=True)
        axes[1].set_title('各科目分数分布')
        axes[1].set_ylabel('分数')
        axes[1].grid(axis='y', linestyle='--', alpha=0.7)
        
        # 将图表嵌入Tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.visual_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 保存图表供导出
        self.current_fig = fig
        
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
        
        # 创建图表
        fig, axes = plt.subplots(1, 2, figsize=(12, 5))
        
        # 1. 总分分布直方图
        if "总分" in self.data.columns:
            sns.histplot(self.data["总分"], bins=10, kde=True, ax=axes[0], color='green')
            axes[0].set_title('总分分布')
            axes[0].set_xlabel('总分')
            axes[0].set_ylabel('学生数量')
            axes[0].grid(axis='y', linestyle='--', alpha=0.7)
        
        # 2. 各分数段人数占比（以第一个科目为例）
        subject = subject_columns[0]
        bins = [0, 60, 70, 80, 90, 100]
        labels = ['0-59', '60-69', '70-79', '80-89', '90-100']
        score_ranges = pd.cut(self.data[subject], bins=bins, labels=labels)
        range_counts = score_ranges.value_counts(normalize=True).reindex(labels) * 100
        
        axes[1].pie(range_counts, labels=labels, autopct='%1.1f%%', startangle=90, colors=sns.color_palette('pastel'))
        axes[1].set_title(f'{subject}各分数段占比')
        
        # 将图表嵌入Tkinter
        canvas = FigureCanvasTkAgg(fig, master=self.visual_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
        
        # 保存图表供导出
        self.current_fig = fig
        
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
            if var_chart.get() and hasattr(self, 'current_fig'):
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
        
        ttk.Button(export_window, text="导出", command=do_export).pack(pady=20)
    
    def show_about(self):
        """显示关于信息"""
        messagebox.showinfo(
            "关于",
            "学生成绩分析系统 v1.0\n\n"
            "一个用于分析学生成绩的工具，支持数据导入、统计分析、数据可视化和结果导出。\n\n"
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

2. 基本操作
   - 导入数据后，系统会在左侧表格显示数据
   - 可以使用底部的按钮进行各种分析

3. 分析功能
   - 基本统计分析：计算各科目平均分、最高分、最低分、及格率等
   - 科目对比分析：生成各科目平均分对比图和分数分布箱线图
   - 成绩分布分析：生成总分分布直方图和分数段占比饼图

4. 结果导出
   - 点击"导出分析结果"按钮
   - 可以选择导出统计数据和图表
   - 统计数据支持Excel和CSV格式，图表支持PNG、JPG和PDF格式

5. 数据保存
   - 点击菜单栏的"文件" -> "保存数据"可以保存当前数据
   - 点击"文件" -> "导出数据"可以将数据导出为新文件
        """
        
        help_text.insert(tk.END, help_content)
        help_text.config(state=tk.DISABLED)

if __name__ == "__main__":
    root = tk.Tk()
    app = StudentGradeAnalysisSystem(root)
    root.mainloop()
