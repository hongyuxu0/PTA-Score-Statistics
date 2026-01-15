import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from collections import defaultdict
import chardet
import re
from datetime import datetime
import traceback  # 用于打印详细异常信息


class ScoreStatisticsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("通用成绩统计工具（适配PTA成绩单）")
        self.root.geometry("900x800")

        # 全局变量
        self.file_paths = []  # 选中的文件路径（原文件）
        self.processed_paths = []  # 处理后的文件路径（带_统计后后缀）
        self.file_info = {}  # 存储文件信息：{原路径: {columns: [], encoding: '', processed_path: '', full_score: ''}}
        self.score_segments = []  # 分数段规则：[(min_rate, max_rate, score), ...]（改为得分率）
        self.segment_count = tk.IntVar(value=0)  # 分数段数量
        self.skip_header = tk.IntVar(value=0)  # 表头跳过行数（你的文件填0）
        self.score_column_var = tk.StringVar()  # 选中的总分列
        self.name_column_var = tk.StringVar()  # 选中的姓名列
        self.id_column_var = tk.StringVar()  # 选中的学号列

        # 创建UI界面
        self.create_widgets()

    def create_widgets(self):
        # 1. 文件上传区域
        frame_file = ttk.LabelFrame(self.root, text="1. 文件/文件夹上传与设置")
        frame_file.pack(fill="x", padx=10, pady=5)

        ttk.Button(frame_file, text="选择单个/多个文件", command=self.select_files).pack(side="left", padx=5, pady=5)
        ttk.Button(frame_file, text="选择文件夹（批量导入）", command=self.select_folder).pack(side="left", padx=5,
                                                                                             pady=5)
        self.file_label = ttk.Label(frame_file, text="未选择文件")
        self.file_label.pack(side="left", padx=10, pady=5)

        # 表头跳过行数
        ttk.Label(frame_file, text="表头跳过行数：").pack(side="left", padx=10, pady=5)
        ttk.Entry(frame_file, textvariable=self.skip_header, width=5).pack(side="left", pady=5)
        ttk.Label(frame_file, text="（你的文件填0）").pack(side="left", pady=5)

        # 核心列选择
        frame_columns = ttk.Frame(frame_file)
        frame_columns.pack(side="left", padx=10, pady=5)

        ttk.Label(frame_columns, text="姓名列：").grid(row=0, column=0, padx=2, pady=2)
        self.name_combobox = ttk.Combobox(frame_columns, textvariable=self.name_column_var, state="readonly", width=15)
        self.name_combobox.grid(row=0, column=1, padx=2, pady=2)

        ttk.Label(frame_columns, text="学号列：").grid(row=1, column=0, padx=2, pady=2)
        self.id_combobox = ttk.Combobox(frame_columns, textvariable=self.id_column_var, state="readonly", width=15)
        self.id_combobox.grid(row=1, column=1, padx=2, pady=2)

        ttk.Label(frame_columns, text="总分列：").grid(row=2, column=0, padx=2, pady=2)
        self.score_combobox = ttk.Combobox(frame_columns, textvariable=self.score_column_var, state="readonly",
                                           width=15)
        self.score_combobox.grid(row=2, column=1, padx=2, pady=2)

        # 2. 分数段设置区域（改为得分率）
        frame_segment = ttk.LabelFrame(self.root, text="2. 得分率-分数段规则设置（适配不同满分）")
        frame_segment.pack(fill="x", padx=10, pady=5)

        ttk.Label(frame_segment, text="分数段数量：").grid(row=0, column=0, padx=5, pady=5)
        ttk.Entry(frame_segment, textvariable=self.segment_count, width=8).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame_segment, text="生成设置项", command=self.create_segment_entries).grid(row=0, column=2, padx=5,
                                                                                               pady=5)

        # 分数段表头（改为得分率）
        ttk.Label(frame_segment, text="起始得分率（0-1）").grid(row=1, column=0, padx=10, pady=2)
        ttk.Label(frame_segment, text="结束得分率（0-1）").grid(row=1, column=1, padx=10, pady=2)
        ttk.Label(frame_segment, text="对应分值").grid(row=1, column=2, padx=10, pady=2)

        self.segment_frame = ttk.Frame(frame_segment)
        self.segment_frame.grid(row=2, column=0, columnspan=3, padx=5, pady=5)

        # 3. 操作按钮区域
        frame_operate = ttk.Frame(self.root)
        frame_operate.pack(fill="x", padx=10, pady=10)

        ttk.Button(frame_operate, text="开始统计（添加统计分）", command=self.start_statistics).pack(side="left", padx=10)
        ttk.Button(frame_operate, text="汇总所有文件总分", command=self.summary_total_score).pack(side="left", padx=10)
        ttk.Button(frame_operate, text="清空日志", command=self.clear_log).pack(side="left", padx=10)

        # 4. 日志显示区域
        frame_log = ttk.LabelFrame(self.root, text="3. 运行日志")
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = tk.Text(frame_log, height=20, font=("Arial", 10))
        scrollbar_y = ttk.Scrollbar(frame_log, orient="vertical", command=self.log_text.yview)
        scrollbar_x = ttk.Scrollbar(frame_log, orient="horizontal", command=self.log_text.xview)
        self.log_text.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        self.log_text.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        # 初始化日志
        self.log("=== 通用成绩统计工具启动完成 ===")
        self.log("提示：1. 选择文件 2. 确认列名 3. 配置得分率-分数段 4. 开始统计 5. 汇总总分")
        self.log("注意：分数段需填写「得分率范围」（例：0,0.59,1；0.6,1.0,2）")

    def custom_ask_float(self, title, prompt):
        """
        自定义浮点数输入弹窗（替代askfloat/askstring，兼容低版本Tkinter）
        :param title: 弹窗标题
        :param prompt: 提示文本
        :return: 输入的浮点数，用户取消则返回None
        """
        result = None
        top = tk.Toplevel(self.root)
        top.title(title)
        top.geometry("400x150")
        top.transient(self.root)  # 置顶
        top.grab_set()  # 模态窗口

        # 提示标签
        label = ttk.Label(top, text=prompt, wraplength=350)
        label.pack(pady=10)

        # 输入框
        entry_var = tk.StringVar()
        entry = ttk.Entry(top, textvariable=entry_var, width=20)
        entry.pack(pady=5)
        entry.focus()

        # 确认按钮事件
        def on_confirm():
            nonlocal result
            val = entry_var.get().strip()
            try:
                num = float(val)
                if num <= 0:
                    messagebox.showwarning("警告", "数值必须大于0！")
                    return
                result = num
                top.destroy()
            except ValueError:
                messagebox.showwarning("警告", "请输入有效的数字（如80、100）！")

        # 取消按钮事件
        def on_cancel():
            nonlocal result
            result = None
            top.destroy()

        # 按钮区域
        btn_frame = ttk.Frame(top)
        btn_frame.pack(pady=10)
        ttk.Button(btn_frame, text="确认", command=on_confirm).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="取消", command=on_cancel).pack(side="left", padx=10)

        # 等待窗口关闭
        self.root.wait_window(top)
        return result

    def detect_encoding(self, file_path):
        with open(file_path, 'rb') as f:
            raw_data = f.read(4096)
        result = chardet.detect(raw_data)
        return result['encoding']

    def try_read_csv(self, file_path, skip_rows, is_full_read=False):
        """
        增强版CSV读取，优先使用utf-8-sig，兼容更多编码
        """
        # 调整编码优先级：优先utf-8-sig，解决中文解码问题
        encodings = [
            'utf-8-sig', 'utf-8', 'gbk', 'gb2312', 'latin-1',
            self.detect_encoding(file_path)  # 最后尝试检测的编码
        ]
        # 去重并过滤None
        encodings = list(filter(None, list(dict.fromkeys(encodings))))

        for enc in encodings:
            try:
                if is_full_read:
                    df = pd.read_csv(file_path, encoding=enc, skiprows=skip_rows, on_bad_lines='skip')
                else:
                    df = pd.read_csv(file_path, encoding=enc, skiprows=skip_rows, nrows=0)
                self.log(f"使用编码 {enc} 成功读取文件：{os.path.basename(file_path)}")
                return df, enc
            except Exception as e:
                self.log(f"尝试编码 {enc} 读取失败：{str(e)[:100]}")  # 只打印前100字符，避免日志过长
                continue
        raise ValueError(f"无法读取文件（所有编码尝试失败）：{file_path}")

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="选择成绩文件",
            filetypes=[("数据文件", "*.csv *.xlsx"), ("CSV文件", "*.csv"), ("Excel文件", "*.xlsx")]
        )
        if not files:
            return

        new_files = [f for f in files if f not in self.file_paths]
        if not new_files:
            messagebox.showinfo("提示", "所选文件已全部添加！")
            return

        self.file_paths.extend(new_files)
        self.file_label.config(text=f"已选 {len(self.file_paths)} 个文件")
        self.log(f"成功添加文件：{[os.path.basename(f) for f in new_files]}")

        first_file = new_files[0]
        skip_rows = self.skip_header.get()

        try:
            if first_file.endswith('.csv'):
                df, encoding = self.try_read_csv(first_file, skip_rows)
            else:
                df = pd.read_excel(first_file, skiprows=skip_rows, nrows=0)
                encoding = "excel"

            columns = df.columns.tolist()
            # 初始化file_info，避免后续KeyError
            self.file_info[first_file] = {"columns": columns, "encoding": encoding, "processed_path": "",
                                          "full_score": 0}

            if not self.name_combobox['values']:
                self.name_combobox['values'] = columns
                self.id_combobox['values'] = columns
                self.score_combobox['values'] = columns

            self.auto_match_columns(columns)

        except Exception as e:
            self.log(f"读取文件列名失败：{str(e)}")
            self.log(f"详细异常：{traceback.format_exc()}")
            messagebox.showerror("错误", f"读取文件结构失败：{str(e)}")

    def select_folder(self):
        folder = filedialog.askdirectory(title="选择成绩文件所在文件夹")
        if not folder:
            return

        file_types = ('.csv', '.xlsx')
        folder_files = [
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(file_types)
        ]
        if not folder_files:
            messagebox.showwarning("提示", "该文件夹下没有CSV/XLSX文件！")
            return

        new_files = [f for f in folder_files if f not in self.file_paths]
        if not new_files:
            messagebox.showinfo("提示", "文件夹内的文件已全部添加！")
            return

        self.file_paths.extend(new_files)
        self.file_label.config(text=f"已选 {len(self.file_paths)} 个文件")
        self.log(f"成功添加文件夹下的文件：{[os.path.basename(f) for f in new_files]}")

        first_file = new_files[0]
        skip_rows = self.skip_header.get()

        try:
            if first_file.endswith('.csv'):
                df, encoding = self.try_read_csv(first_file, skip_rows)
            else:
                df = pd.read_excel(first_file, skiprows=skip_rows, nrows=0)
                encoding = "excel"

            columns = df.columns.tolist()
            # 初始化file_info
            self.file_info[first_file] = {"columns": columns, "encoding": encoding, "processed_path": "",
                                          "full_score": 0}

            if not self.name_combobox['values']:
                self.name_combobox['values'] = columns
                self.id_combobox['values'] = columns
                self.score_combobox['values'] = columns

            self.auto_match_columns(columns)

        except Exception as e:
            self.log(f"读取文件夹文件列名失败：{str(e)}")
            self.log(f"详细异常：{traceback.format_exc()}")
            messagebox.showerror("错误", f"读取文件夹文件结构失败：{str(e)}")

    def auto_match_columns(self, columns):
        for col in columns:
            col_str = str(col).strip()
            if not self.name_column_var.get() and (col_str == "Unnamed: 1" or "姓名" in col_str):
                self.name_column_var.set(col_str)
            if not self.id_column_var.get() and (col_str == "Unnamed: 2" or "学号" in col_str):
                self.id_column_var.set(col_str)
            if not self.score_column_var.get() and (col_str == "Unnamed: 4" or "总分" in col_str):
                self.score_column_var.set(col_str)

        self.log(
            f"自动匹配列名：姓名列={self.name_column_var.get()}, 学号列={self.id_column_var.get()}, 总分列={self.score_column_var.get()}")

    def create_segment_entries(self):
        for widget in self.segment_frame.winfo_children():
            widget.destroy()

        count = self.segment_count.get()
        if count <= 0:
            messagebox.showwarning("警告", "分数段数量必须大于0！")
            return

        self.segment_vars = []
        for i in range(count):
            min_var = tk.StringVar()
            max_var = tk.StringVar()
            score_var = tk.StringVar()
            self.segment_vars.append((min_var, max_var, score_var))

            ttk.Entry(self.segment_frame, textvariable=min_var, width=10, justify="center").grid(row=i, column=0,
                                                                                                 padx=10, pady=3)
            ttk.Entry(self.segment_frame, textvariable=max_var, width=10, justify="center").grid(row=i, column=1,
                                                                                                 padx=10, pady=3)
            ttk.Entry(self.segment_frame, textvariable=score_var, width=10, justify="center").grid(row=i, column=2,
                                                                                                   padx=10, pady=3)

        self.log(f"生成 {count} 个得分率-分数段设置项，请填写（例：0,0.59,1；0.6,1.0,2）")

    def parse_segment_rules(self):
        self.score_segments = []
        if not hasattr(self, 'segment_vars'):
            return False

        for i, (min_var, max_var, score_var) in enumerate(self.segment_vars):
            try:
                min_rate = float(min_var.get().strip())
                max_rate = float(max_var.get().strip())
                assign_score = float(score_var.get().strip())

                if not (0 <= min_rate <= 1) or not (0 <= max_rate <= 1):
                    messagebox.showwarning("警告", f"第{i + 1}个分数段：得分率需在0-1之间！")
                    return False
                if min_rate > max_rate:
                    messagebox.showwarning("警告", f"第{i + 1}个分数段：起始得分率不能大于结束得分率！")
                    return False

                self.score_segments.append((min_rate, max_rate, assign_score))
            except ValueError:
                messagebox.showerror("错误", f"第{i + 1}个分数段：请输入有效的数字！")
                return False

        self.log(f"得分率-分数段规则：{self.score_segments}")
        return True

    def extract_full_score(self, score_column_name, file_name):
        """
        从总分列名提取满分，兼容更多格式
        """
        # 扩展正则匹配规则
        patterns = [
            r'总分\((\d+\.?\d*)\)',  # 总分(80)、总分(80.0)
            r'总分\((\d+\.?\d*),',  # 总分(80,排名)
            r'总分[^\d]*(\d+\.?\d*)',  # 总分80、总分_80分
        ]

        for pattern in patterns:
            match = re.search(pattern, score_column_name)
            if match:
                full_score = float(match.group(1))
                self.log(f"从列名「{score_column_name}」提取到{file_name}的满分：{full_score}")
                return full_score

        # 正则匹配失败，使用自定义弹窗输入
        full_score = self.custom_ask_float("输入满分",
                                           f"无法从列名「{score_column_name}」提取{file_name}的满分，请输入（如80、100）：")
        if full_score is None:  # 用户取消
            raise ValueError("用户取消满分输入")
        return full_score

    def clean_score(self, score_value):
        if pd.isna(score_value) or score_value == '' or str(score_value).strip() in ['未开考', '缺考']:
            return 0.0
        try:
            num_str = re.search(r'(\d+\.?\d*)', str(score_value))
            if num_str:
                return float(num_str.group(1))
            else:
                return 0.0
        except:
            return 0.0

    def get_assigned_score_by_rate(self, rate):
        """按得分率匹配分数段"""
        rate = max(0.0, min(1.0, rate))
        for min_rate, max_rate, assign_score in self.score_segments:
            if min_rate <= rate <= max_rate:
                return assign_score
        return 0.0

    def start_statistics(self):
        if not self.file_paths:
            messagebox.showwarning("警告", "请先选择文件！")
            return

        required_cols = [self.name_column_var.get(), self.id_column_var.get(), self.score_column_var.get()]
        if not all(required_cols):
            messagebox.showwarning("警告", "请选择姓名列、学号列、总分列！")
            return

        if not self.parse_segment_rules():
            return

        self.processed_paths = []
        for file_path in self.file_paths:
            try:
                filename = os.path.basename(file_path)
                skip_rows = self.skip_header.get()
                self.log(f"开始处理文件：{filename}（跳过{skip_rows}行）")

                # 初始化当前文件的file_info，避免KeyError
                if file_path not in self.file_info:
                    self.file_info[file_path] = {"columns": [], "encoding": "", "processed_path": "", "full_score": 0}

                # 读取文件
                if file_path.endswith('.csv'):
                    df, encoding = self.try_read_csv(file_path, skip_rows, is_full_read=True)
                    self.file_info[file_path]["encoding"] = encoding  # 现在赋值不会KeyError
                else:
                    df = pd.read_excel(file_path, skiprows=skip_rows)
                    self.file_info[file_path]["encoding"] = "excel"

                # 提取该文件的满分
                full_score = self.extract_full_score(self.score_column_var.get(), filename)
                self.file_info[file_path]["full_score"] = full_score
                self.log(f"文件 {filename} 的总分满分为：{full_score}")

                # 检查核心列
                for col in required_cols:
                    if col not in df.columns:
                        raise ValueError(f"缺少核心列：{col}")

                # 计算得分率+统计分
                df['实际得分'] = df[self.score_column_var.get()].apply(self.clean_score)
                df['得分率'] = df['实际得分'] / full_score
                df['统计分'] = df['得分率'].apply(self.get_assigned_score_by_rate)
                self.log(f"文件 {filename}：统计分计算完成，共 {len(df)} 条数据")

                # 保存文件
                save_path = file_path.replace('.csv', '_统计后.csv') if file_path.endswith(
                    '.csv') else file_path.replace('.xlsx', '_统计后.xlsx')
                if file_path.endswith('.csv'):
                    # 保存时用utf-8-sig，避免中文乱码
                    df.to_csv(save_path, index=False, encoding='utf-8-sig')
                else:
                    df.to_excel(save_path, index=False)

                self.file_info[file_path]["processed_path"] = save_path
                self.processed_paths.append(save_path)
                self.log(f"文件 {filename} 已保存为：{os.path.basename(save_path)}")

            except Exception as e:
                error_msg = f"处理 {filename} 失败：{str(e)}"
                self.log(error_msg)
                self.log(f"详细异常堆栈：{traceback.format_exc()}")
                messagebox.showerror("处理错误", error_msg)

        messagebox.showinfo("完成", "所有文件统计分添加完成！")
        self.log("=== 统计分添加任务全部结束 ===")

    def summary_total_score(self):
        if not self.processed_paths:
            messagebox.showwarning("警告", "请先执行「开始统计」生成处理后的文件！")
            return

        required_cols = [self.name_column_var.get(), self.id_column_var.get()]
        if not all(required_cols):
            messagebox.showwarning("警告", "请选择姓名列、学号列！")
            return

        total_score_dict = defaultdict(float)
        file_count = 0

        for processed_path in self.processed_paths:
            try:
                filename = os.path.basename(processed_path)
                skip_rows = self.skip_header.get()
                self.log(f"汇总文件：{filename}（跳过{skip_rows}行）")

                # 汇总阶段也使用增强版的try_read_csv读取文件，自动适配编码
                if processed_path.endswith('.csv'):
                    df, _ = self.try_read_csv(processed_path, skip_rows, is_full_read=True)
                else:
                    df = pd.read_excel(processed_path, skiprows=skip_rows)

                if '统计分' not in df.columns:
                    raise ValueError("文件中无「统计分」列，请先执行「开始统计」！")

                file_count += 1
                self.log(f"汇总 {filename}：{len(df)} 条数据")

                name_col = self.name_column_var.get()
                id_col = self.id_column_var.get()

                for _, row in df.iterrows():
                    name = str(row[name_col]).strip() if not pd.isna(row[name_col]) else "未知姓名"
                    student_id = str(row[id_col]).strip() if not pd.isna(row[id_col]) else "未知学号"
                    stat_score = float(row['统计分']) if not pd.isna(row['统计分']) else 0.0

                    total_score_dict[(name, student_id)] += stat_score

            except Exception as e:
                error_msg = f"汇总 {filename} 失败：{str(e)}"
                self.log(error_msg)
                self.log(f"详细异常堆栈：{traceback.format_exc()}")
                messagebox.showerror("汇总错误", error_msg)

        if not total_score_dict:
            messagebox.showwarning("警告", "无有效数据可汇总！")
            return

        summary_data = []
        for (name, student_id), total_score in total_score_dict.items():
            summary_data.append({
                "姓名": name,
                "学号": student_id,
                "参与文件数": file_count,
                "统计分总分": round(total_score, 2)
            })

        summary_df = pd.DataFrame(summary_data)
        summary_df = summary_df.sort_values("统计分总分", ascending=False).reset_index(drop=True)

        summary_dir = os.path.dirname(self.processed_paths[0])
        summary_path = os.path.join(summary_dir, "成绩统计总分汇总.csv")
        # 汇总文件用utf-8-sig保存，避免中文乱码
        summary_df.to_csv(summary_path, index=False, encoding='utf-8-sig')

        self.log(f"汇总完成！共 {len(summary_df)} 名学生，汇总文件已保存至：{summary_path}")
        messagebox.showinfo("汇总完成", f"成功汇总 {file_count} 个文件！\n汇总文件路径：\n{summary_path}")

    def log(self, msg):
        time_str = datetime.now().strftime("[%H:%M:%S]")
        self.log_text.insert(tk.END, f"{time_str} {msg}\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        self.log("日志已清空")


if __name__ == "__main__":
    print("=== 通用成绩统计工具 ===")
    print("运行前请确保已安装依赖：pip install pandas openpyxl chardet")
    print("你的文件表头跳过行数设为0，列名会自动匹配Unnamed:1（姓名）、Unnamed:2（学号）、Unnamed:4（总分）")
    print("分数段需填写「得分率范围」（例：0,0.59,1；0.6,1.0,2）")

    root = tk.Tk()
    app = ScoreStatisticsApp(root)
    root.mainloop()
