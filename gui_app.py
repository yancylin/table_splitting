import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os

class PlateAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("表格拆分")
        self.root.geometry("800x800")
        self.root.resizable(True, True)

        # 设置图标（如果有的话）
        try:
            self.root.iconbitmap("app.ico")
        except:
            pass

        from excel_filter import ExcelFilter
        self.analyzer = ExcelFilter()
        self.data_file_path = None
        self.filter_file_path = None
        self.data_dataframe = None
        self.filter_dataframe = None

        self.setup_ui()

    def setup_ui(self):
        """设置用户界面 - 左右两列布局"""
        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        title_label = ttk.Label(main_frame, text="表格拆分工具",
                                font=("Arial", 16, "bold"))
        title_label.pack(pady=(0, 10))

        # 说明文本
        desc_text = "本工具用于拆分表格，可以根据一个表中的列，筛选出符合条件的数据"
        desc_label = ttk.Label(main_frame, text=desc_text, wraplength=600)
        desc_label.pack(pady=(0, 10))

        # 创建左右两列框架
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True)

        # 左侧操作区域
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))

        # 右侧结果显示区域
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))

        # 左侧操作控件
        # 数据文件选择区域
        data_file_frame = ttk.LabelFrame(left_frame, text="1. 选择数据源文件", padding="10")
        data_file_frame.pack(fill=tk.X, pady=(0, 10))

        self.data_file_label = ttk.Label(data_file_frame, text="未选择文件")
        self.data_file_label.pack(anchor=tk.W)

        data_file_button = ttk.Button(data_file_frame, text="选择Excel/CSV文件",
                                      command=self.select_data_file)
        data_file_button.pack(pady=(5, 0))

        # 数据列选择区域
        data_column_frame = ttk.LabelFrame(left_frame, text="2. 选择数据源筛选列", padding="10")
        data_column_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(data_column_frame, text="选择需要筛选的列:").pack(anchor=tk.W)

        self.data_column_var = tk.StringVar()
        self.data_column_combo = ttk.Combobox(data_column_frame, textvariable=self.data_column_var, state="readonly")
        self.data_column_combo.pack(fill=tk.X, pady=(5, 0))

        # 筛选条件文件选择区域
        filter_file_frame = ttk.LabelFrame(left_frame, text="3. 选择筛选条件文件", padding="10")
        filter_file_frame.pack(fill=tk.X, pady=(0, 10))

        self.filter_file_label = ttk.Label(filter_file_frame, text="未选择文件")
        self.filter_file_label.pack(anchor=tk.W)

        filter_file_button = ttk.Button(filter_file_frame, text="选择Excel/CSV文件",
                                        command=self.select_filter_file)
        filter_file_button.pack(pady=(5, 0))

        # 筛选条件列选择区域
        filter_column_frame = ttk.LabelFrame(left_frame, text="4. 选择筛选条件列", padding="10")
        filter_column_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(filter_column_frame, text="选择筛选条件列:").pack(anchor=tk.W)

        self.filter_column_var = tk.StringVar()
        self.filter_column_combo = ttk.Combobox(filter_column_frame, textvariable=self.filter_column_var,
                                                state="readonly")
        self.filter_column_combo.pack(fill=tk.X, pady=(5, 0))

        # 统计列选择区域
        stats_column_frame = ttk.LabelFrame(left_frame, text="5. 选择统计数据列（可选）", padding="10")
        stats_column_frame.pack(fill=tk.X, pady=(0, 10))

        # 新增功能：筛选表2中不存在的数据
        missing_data_frame = ttk.LabelFrame(left_frame, text="6. 筛选选项", padding="10")
        missing_data_frame.pack(fill=tk.X, pady=(0, 10))

        self.exclude_missing_var = tk.BooleanVar()
        exclude_missing_check = ttk.Checkbutton(
            missing_data_frame,
            text="排除表2中不存在的数据",
            variable=self.exclude_missing_var
        )
        exclude_missing_check.pack(anchor=tk.W)

        ttk.Label(stats_column_frame, text="选择需要统计的列:").pack(anchor=tk.W)

        self.stats_column_var = tk.StringVar()
        self.stats_column_combo = ttk.Combobox(stats_column_frame, textvariable=self.stats_column_var,
                                               state="readonly")
        self.stats_column_combo.pack(fill=tk.X, pady=(5, 0))

        # 分析按钮
        button_frame = ttk.Frame(left_frame)
        button_frame.pack(pady=20)

        analyze_button = ttk.Button(button_frame, text="开始筛选",
                                    command=self.analyze_data, state="disabled")
        analyze_button.pack(side=tk.LEFT, padx=(0, 10))
        self.analyze_button = analyze_button

        export_button = ttk.Button(button_frame, text="导出结果",
                                   command=self.export_results, state="disabled")
        export_button.pack(side=tk.LEFT)
        self.export_button = export_button

        # 右侧结果显示区域
        result_frame = ttk.LabelFrame(right_frame, text="筛选结果", padding="10")
        result_frame.pack(fill=tk.BOTH, expand=True)

        # 创建文本框和滚动条
        text_frame = ttk.Frame(result_frame)
        text_frame.pack(fill=tk.BOTH, expand=True)

        self.result_text = tk.Text(text_frame, height=20)
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.result_text.yview)
        self.result_text.configure(yscrollcommand=scrollbar.set)

        self.result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 状态栏
        self.status_var = tk.StringVar()
        self.status_var.set("就绪 - 请选择数据文件")
        status_bar = ttk.Label(main_frame, textvariable=self.status_var, relief=tk.SUNKEN)
        status_bar.pack(fill=tk.X, pady=(10, 0))

        # 存储分析结果
        self.last_result = None

    def select_data_file(self):
        """选择数据文件"""
        file_types = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]

        filename = filedialog.askopenfilename(
            title="选择数据文件",
            filetypes=file_types
        )

        if filename:
            self.data_file_path = filename
            self.data_file_label.config(text=os.path.basename(filename))
            self.status_var.set(f"已选择数据文件: {os.path.basename(filename)}")

            # 尝试读取文件并获取列名
            try:
                self.data_dataframe = self.analyzer.load_data(filename)

                # 更新列名下拉框
                columns = list(self.data_dataframe.columns)
                self.data_column_combo['values'] = columns
                self.stats_column_combo['values'] = columns

                if len(columns) > 0:
                    self.data_column_combo.current(0)

                self.result_text.delete(1.0, tk.END)
                self.result_text.insert(tk.END, f"数据文件加载成功!\n")
                self.result_text.insert(tk.END, f"数据行数: {self.data_dataframe.shape[0]}\n")
                self.result_text.insert(tk.END, f"数据列数: {self.data_dataframe.shape[1]}\n")
                self.result_text.insert(tk.END, f"可用列: {', '.join(columns)}\n\n")

                # 检查是否可以启用分析按钮
                self.check_analysis_ready()

            except Exception as e:
                messagebox.showerror("错误", f"读取文件时出错: {str(e)}")
                self.status_var.set("数据文件读取失败")

    def select_filter_file(self):
        """选择筛选条件文件"""
        file_types = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]

        filename = filedialog.askopenfilename(
            title="选择筛选条件文件",
            filetypes=file_types
        )

        if filename:
            self.filter_file_path = filename
            self.filter_file_label.config(text=os.path.basename(filename))
            self.status_var.set(f"已选择筛选条件文件: {os.path.basename(filename)}")

            # 尝试读取文件并获取列名
            try:
                self.filter_dataframe = self.analyzer.load_filter(filename)

                # 更新列名下拉框
                self.filter_column_combo['values'] = list(self.filter_dataframe.columns)
                if len(self.filter_dataframe.columns) > 0:
                    self.filter_column_combo.current(0)

                self.result_text.insert(tk.END, f"筛选条件文件加载成功!\n")
                self.result_text.insert(tk.END, f"筛选条件行数: {self.filter_dataframe.shape[0]}\n")
                self.result_text.insert(tk.END, f"筛选条件列数: {self.filter_dataframe.shape[1]}\n")
                self.result_text.insert(tk.END, f"可用列: {', '.join(self.filter_dataframe.columns)}\n\n")

                # 检查是否可以启用分析按钮
                self.check_analysis_ready()

            except Exception as e:
                messagebox.showerror("错误", f"读取筛选条件文件时出错: {str(e)}")
                self.status_var.set("筛选条件文件读取失败")

    def check_analysis_ready(self):
        """检查是否可以启用分析按钮"""
        if (self.data_file_path and self.filter_file_path and
                self.data_column_var.get() and self.filter_column_var.get()):
            self.analyze_button.config(state="normal")
        else:
            self.analyze_button.config(state="disabled")

    def analyze_data(self):
        """分析数据"""
        if not self.data_file_path:
            messagebox.showwarning("警告", "请先选择数据文件")
            return

        if not self.filter_file_path:
            messagebox.showwarning("警告", "请先选择筛选条件文件")
            return

        if not self.data_column_var.get():
            messagebox.showwarning("警告", "请选择数据源筛选列")
            return

        if not self.filter_column_var.get():
            messagebox.showwarning("警告", "请选择筛选条件列")
            return

        # 清空结果文本框
        self.result_text.delete(1.0, tk.END)
        self.result_text.insert(tk.END, "正在筛选数据，请稍候...\n")
        self.root.update()

        try:
            # 加载数据
            self.analyzer.load_data(self.data_file_path)
            self.analyzer.load_filter(self.filter_file_path)

            # 根据选项决定执行哪种筛选
            if self.exclude_missing_var.get():
                # 筛选出表2中不存在的数据
                result_df = self.analyzer.find_missing_in_filter(
                    self.data_column_var.get(),
                    self.filter_column_var.get()
                )
                operation_desc = "查找表2中不存在的数据"
            else:
                # 正常筛选（表2中存在的数据）
                result_df = self.analyzer.filter_by_column(
                    self.data_column_var.get(),
                    self.filter_column_var.get()
                )
                operation_desc = "筛选表2中存在的数据"

            # 显示结果
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"=== {operation_desc} ===\n\n")
            self.result_text.insert(tk.END, f"原始数据行数: {len(self.analyzer.data_df)}\n")
            self.result_text.insert(tk.END, f"筛选后行数: {len(result_df)}\n")
            self.result_text.insert(tk.END, f"筛选条件数量: {len(self.analyzer.filter_df)}\n\n")

            # 如果选择了统计列，则计算统计信息
            stats_column = self.stats_column_var.get()
            if stats_column and stats_column in result_df.columns:
                try:
                    total = self.analyzer.sum_column(result_df, stats_column)
                    self.result_text.insert(tk.END, f"列 '{stats_column}' 的合计: {total}\n\n")
                except Exception as e:
                    self.result_text.insert(tk.END, f"统计列 '{stats_column}' 时出错: {str(e)}\n\n")

            if len(result_df) > 0:
                # 添加基本统计信息
                self.result_text.insert(tk.END, "=== 数据统计 ===\n")
                # 统计数值型列的基本信息
                numeric_columns = result_df.select_dtypes(include=['number']).columns.tolist()
                if numeric_columns:
                    self.result_text.insert(tk.END, f"数值型列: {', '.join(numeric_columns)}\n")
                    for col in numeric_columns:
                        col_sum = result_df[col].sum()
                        col_mean = result_df[col].mean()
                        col_count = result_df[col].count()
                        self.result_text.insert(tk.END,
                                                f"  {col} - 总和: {col_sum}, 平均值: {col_mean:.2f}, 有效值数量: {col_count}\n")

                # 显示各列非空值数量
                self.result_text.insert(tk.END, "\n各列非空值统计:\n")
                for col in result_df.columns:
                    non_null_count = result_df[col].count()
                    self.result_text.insert(tk.END, f"  {col}: {non_null_count}/{len(result_df)}\n")

                self.result_text.insert(tk.END, "\n前10行数据预览:\n")
                self.result_text.insert(tk.END, str(result_df.head(10)) + "\n")

            # 存储结果并启用导出按钮
            self.last_result = result_df
            self.export_button.config(state="normal")
            self.status_var.set("筛选完成 - 可以导出结果")

        except Exception as e:
            self.result_text.delete(1.0, tk.END)
            self.result_text.insert(tk.END, f"筛选失败:\n{str(e)}")
            self.status_var.set("筛选失败")
            self.export_button.config(state="disabled")

    def export_results(self):
        """导出分析结果"""
        if self.last_result is None:
            messagebox.showwarning("警告", "请先完成筛选")
            return

        # 选择保存文件
        file_types = [
            ("Excel files", "*.xlsx"),
            ("CSV files", "*.csv")
        ]

        output_file = filedialog.asksaveasfilename(
            title="保存筛选结果",
            defaultextension=".xlsx",
            filetypes=file_types
        )

        if not output_file:
            return

        self.status_var.set("正在导出结果...")
        self.root.update()

        try:
            # 保存结果
            success = self.analyzer.export_result(self.last_result, output_file)

            if success:
                self.result_text.insert(tk.END, f"\n结果已成功导出到: {output_file}\n")
                self.status_var.set("结果导出完成")

                # 询问是否打开文件
                if messagebox.askyesno("导出完成", "结果已成功导出，是否打开文件？"):
                    os.startfile(output_file) if os.name == 'nt' else os.system(f'open "{output_file}"')
            else:
                raise Exception("导出失败")

        except Exception as e:
            messagebox.showerror("导出失败", str(e))
            self.status_var.set("导出失败")
