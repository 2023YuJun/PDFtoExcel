import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import config
import PDFtoExcel
import time

class WinForm:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        self.center_window(585, 260)  # 设置窗体大小并居中
        self.root.resizable(True, True)  # 允许调整大小
        self.root.minsize(width=480, height=180)
        self.not_settings = True

        # 加载默认文件路径
        self.file_paths = config.get_file_paths()
        self.pdf_path = self.file_paths.get("PDF_path", "")
        self.excel_path = self.file_paths.get("Excel_path", "")

        # 创建界面元素
        self.create_widgets()

    def center_window(self, width=600, height=400):
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def create_widgets(self):
        # 使用 Grid 布局
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)
        self.root.grid_rowconfigure(2, weight=1)

        # PDF 文件选择
        tk.Label(self.root, text="选择 PDF 文件:", anchor="w").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.pdf_entry = tk.Entry(self.root, width=50)
        self.pdf_entry.insert(0, self.pdf_path)
        self.pdf_entry.grid(row=0, column=1, padx=10, pady=10, sticky="w")
        tk.Button(self.root, text="浏览文件", command=self.select_pdf).grid(row=0, column=2, padx=10, pady=10, sticky="e")

        # Excel 文件保存路径选择
        tk.Label(self.root, text="选择 Excel 保存路径:", anchor="w").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        self.excel_entry = tk.Entry(self.root, width=50)
        self.excel_entry.insert(0, self.excel_path)
        self.excel_entry.grid(row=1, column=1, padx=10, pady=10, sticky="w")
        tk.Button(self.root, text="浏览文件", command=self.select_excel).grid(row=1, column=2, padx=10, pady=10, sticky="e")

        # 高级设置按钮
        self.advanced_settings_button = tk.Button(self.root, text="高级设置", command=self.open_advanced_settings)
        self.advanced_settings_button.grid(row=2, column=0, padx=10, pady=10, sticky="sw")

        # 转换按钮
        self.convert_button = tk.Button(self.root, text="开始转换", command=self.convert_pdf_to_excel)
        self.convert_button.grid(row=2, column=2, padx=10, pady=10, sticky="se")

    def select_pdf(self):
        """选择 PDF 文件"""
        initial_dir = os.path.dirname(self.pdf_path) if self.pdf_path else ""
        file_path = filedialog.askopenfilename(initialdir=initial_dir, filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.pdf_entry.delete(0, tk.END)
            self.pdf_entry.insert(0, file_path)

    def select_excel(self):
        """选择 Excel 文件保存路径"""
        initial_dir = os.path.dirname(self.excel_path) if self.excel_path else ""
        file_path = filedialog.asksaveasfilename(initialdir=initial_dir, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if file_path:
            self.excel_entry.delete(0, tk.END)
            self.excel_entry.insert(0, file_path)

    def convert_pdf_to_excel(self):
        """调用 PDF 转 Excel 的功能"""
        pdf_path = self.pdf_entry.get()
        excel_path = self.excel_entry.get()
        if not pdf_path or not excel_path:
            messagebox.showerror("错误", "请先选择 PDF 文件和 Excel 保存路径！")
            return
        
        # 自动保存路径
        config.set_pdf_path(pdf_path)
        config.set_excel_path(excel_path)

        # 如果用户未点击高级设置，恢复默认设置
        if self.not_settings:
            config.reset_table_settings()

        self.convert_button.config(text="转换中...", state=tk.DISABLED)
        self.root.update()

        # 调用 PDF 转 Excel 的功能
        start_time = time.time()
        result = PDFtoExcel.pdf_to_excel(pdf_path, excel_path)
        end_time = time.time()
        elapsed_time = end_time - start_time
        self.convert_button.config(text="开始转换", state=tk.NORMAL)
        messagebox.showinfo("转换完成", result+f"\n总耗时 {elapsed_time:.2f} 秒。")

    def open_advanced_settings(self):
        """打开高级设置窗体"""
        self.not_settings = False
        self.advanced_settings_form = AdvancedSettingsForm(self.root, self.pdf_path)




class AdvancedSettingsForm:
    def __init__(self, parent, pdf_path):
        self.parent = parent
        self.pdf_path = pdf_path
        self.settings_window = tk.Toplevel(self.parent)
        self.settings_window.title("高级设置")
        self.settings_window.geometry("400x500")
        self.settings_window.resizable(False, False)  # 允许调整大小
        self.settings_window.transient(self.parent)  # 设置为父窗体的子窗体
        self.settings_window.grab_set()  # 独占焦点
        self.center_window(self.settings_window, 400, 500)  # 居中显示
        self.create_widgets()

    def center_window(self, window, width=400, height=500):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    def create_widgets(self):
        settings = config.get_table_settings()

        # 配置网格布局权重，使控件随窗体大小调整
        self.settings_window.grid_columnconfigure(0, weight=1)
        self.settings_window.grid_columnconfigure(1, weight=2)  # 第二列权重更大，控件会扩展
        self.settings_window.grid_columnconfigure(2, weight=1)
        for i in range(9):  # 根据控件行数调整
            self.settings_window.grid_rowconfigure(i, weight=1)

        # 垂直线策略
        self.vertical_strategy_label = ttk.Label(self.settings_window, text="垂直线策略:")
        self.vertical_strategy_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.vertical_strategy = ttk.Combobox(self.settings_window, values=["lines", "lines_strict", "text", "explicit"])
        self.vertical_strategy.set(settings["vertical_strategy"])
        self.vertical_strategy.grid(row=0, column=1, padx=10, pady=10, sticky="we")
        self.create_tooltip(self.vertical_strategy_label, """
                                                            默认值："lines"
                                                            描述：指定如何检测垂直线（表格的列分隔符）。可选值：
                                                            "lines"：使用原有页面上的垂直线作为分隔符。
                                                            "lines_strict"：仅使用 明确的垂直线 ，忽略文本对齐。
                                                            "text"：根据文本对齐方式推断垂直线。
                                                            "explicit"：仅使用 明确的垂直线 中指定的线。
                                                            """)

        # 水平线策略
        self.horizontal_strategy_label = ttk.Label(self.settings_window, text="水平线策略:")
        self.horizontal_strategy_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.horizontal_strategy = ttk.Combobox(self.settings_window, values=["lines", "lines_strict", "text", "explicit"])
        self.horizontal_strategy.set(settings["horizontal_strategy"])
        self.horizontal_strategy.grid(row=1, column=1, padx=10, pady=10, sticky="we")
        self.create_tooltip(self.horizontal_strategy_label, """
                                                            默认值："lines"
                                                            描述：指定如何检测水平线（表格的行分隔符）。可选值：
                                                            "lines"：使用页面上的水平线作为分隔符。
                                                            "lines_strict"：仅使用 明确的水平线 ，忽略文本对齐。
                                                            "text"：根据文本对齐方式推断水平线。
                                                            "explicit"：仅使用 明确的水平线 中指定的线。
                                                            """)

        # 明确的垂直线
        self.explicit_vertical_lines_label = ttk.Label(self.settings_window, text="明确的垂直线:")
        self.explicit_vertical_lines_label.grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.explicit_vertical_lines = ttk.Entry(self.settings_window)
        self.explicit_vertical_lines.insert(0, "|".join(map(str, settings["explicit_vertical_lines"])))
        self.explicit_vertical_lines.grid(row=2, column=1, padx=10, pady=10, sticky="we")
        ttk.Button(self.settings_window, text="获取", command=self.get_vertical_lines).grid(row=2, column=2, padx=10, pady=10, sticky="w")
        self.create_tooltip(self.explicit_vertical_lines_label, """
                                                                默认值：[]
                                                                描述：明确指定垂直线的位置，填入的元素只可以是数字，表示线条的 x 坐标，
                                                                坐标原点为PDF文档页面的左下角，单位为页面为 适合页面 缩放时的像数值，
                                                                建议修改此内容时，先通过获取按钮得到默认识别到的垂直线后再修改，输入格式为x1|x2|x3，
                                                                注意：使用此参数建议将垂直线策略修改为"explicit"，提高识别的准确性
                                                                """)


        # 明确的水平线
        self.explicit_horizontal_lines_label = ttk.Label(self.settings_window, text="明确的水平线:")
        self.explicit_horizontal_lines_label.grid(row=3, column=0, padx=10, pady=10, sticky="e")
        self.explicit_horizontal_lines = ttk.Entry(self.settings_window)
        self.explicit_horizontal_lines.insert(0, "|".join(map(str, settings["explicit_horizontal_lines"])))
        self.explicit_horizontal_lines.grid(row=3, column=1, padx=10, pady=10, sticky="we")
        ttk.Button(self.settings_window, text="获取", command=self.get_horizontal_lines).grid(row=3, column=2, padx=10, pady=10, sticky="w")
        self.create_tooltip(self.explicit_horizontal_lines_label, """
                                                                默认值：[]
                                                                描述：明确指定水平线的位置，填入的元素只可以是数字，表示线条的 y 坐标，
                                                                坐标原点为PDF文档页面的左下角，单位为页面为 适合页面 缩放时的像数值，
                                                                建议修改此内容时，先通过获取按钮得到默认识别到的水平线后再修改，输入格式为y1|y2|y3
                                                                注意：使用此参数建议将水平线策略修改为"explicit"，提高识别的准确性
                                                                """)
        # 捕捉距离
        self.snap_tolerance_label = ttk.Label(self.settings_window, text="捕捉距离:")
        self.snap_tolerance_label.grid(row=4, column=0, padx=10, pady=10, sticky="e")
        self.snap_tolerance = ttk.Entry(self.settings_window)
        self.snap_tolerance.insert(0, settings["snap_tolerance"])
        self.snap_tolerance.grid(row=4, column=1, padx=10, pady=10, sticky="we")
        self.create_tooltip(self.snap_tolerance_label, """
                                                        默认值：3
                                                        描述：如果两条平行线之间的距离小于此值，它们会被“捕捉”到同一水平或垂直位置。
                                                        """)

        # 合并距离
        self.join_tolerance_label = ttk.Label(self.settings_window, text="合并距离:")
        self.join_tolerance_label.grid(row=5, column=0, padx=10, pady=10, sticky="e")
        self.join_tolerance = ttk.Entry(self.settings_window)
        self.join_tolerance.insert(0, settings["join_tolerance"])
        self.join_tolerance.grid(row=5, column=1, padx=10, pady=10, sticky="we")
        self.create_tooltip(self.join_tolerance_label,  """
                                                        默认值：3
                                                        描述：如果两条线段在同一无限延长线上，且它们的端点距离小于此值，则会被合并为一条线。
                                                        """)

        # 文本距离
        self.text_tolerance_label = ttk.Label(self.settings_window, text="文本距离:")
        self.text_tolerance_label.grid(row=6, column=0, padx=10, pady=10, sticky="e")
        self.text_tolerance = ttk.Entry(self.settings_window)
        self.text_tolerance.insert(0, settings["text_tolerance"])
        self.text_tolerance.grid(row=6, column=1, padx=10, pady=10, sticky="we")
        self.create_tooltip(self.text_tolerance_label, """
                                                            默认值：3
                                                            描述：当使用文本策略时，字符之间的最大距离（像素）。如果字符间距小于此值，则会被视为同一单词。
                                                            """)

        # 保存设置按钮
        ttk.Button(self.settings_window, text="恢复默认", command=self.reset_table_settings).grid(row=7, column=0, pady=10, sticky="ew")
        ttk.Button(self.settings_window, text="保存设置", command=self.save_settings).grid(row=7, column=2, pady=10,sticky="ew")

        # 添加关闭事件处理
        self.settings_window.protocol("WM_DELETE_WINDOW", self.on_closing)

    def get_vertical_lines(self):
        """获取明确的垂直线"""
        if self.pdf_path == "":
            messagebox.showerror("错误", "请先选择PDF文件")
            return
        try:
            vertical_lines, _ = PDFtoExcel.get_lines(self.pdf_path)
            lines_str = "|".join(f"{line['x0']:.2f}" for line in vertical_lines)
            self.explicit_vertical_lines.delete(0, tk.END)
            self.explicit_vertical_lines.insert(0, lines_str)
            self.vertical_strategy.set("explicit")
        except Exception as e:
            messagebox.showerror("错误", f"获取垂直线失败：{e}")

    def get_horizontal_lines(self):
        """获取明确的水平线"""
        if self.pdf_path == "":
            messagebox.showerror("错误", "请先选择PDF文件")
            return
        try:
            _, horizontal_lines = PDFtoExcel.get_lines(self.pdf_path)
            lines_str = "|".join(f"{line['y0']:.2f}" for line in horizontal_lines)
            self.explicit_horizontal_lines.delete(0, tk.END)
            self.explicit_horizontal_lines.insert(0, lines_str)
            self.horizontal_strategy.set("explicit")
        except Exception as e:
            messagebox.showerror("错误", f"获取水平线失败：{e}")

    def save_settings(self):
        """保存高级设置"""
        try:
            vertical_strategy = self.vertical_strategy.get()
            horizontal_strategy = self.horizontal_strategy.get()
            snap_tolerance = int(self.snap_tolerance.get())
            join_tolerance = int(self.join_tolerance.get())
            text_tolerance = int(self.text_tolerance.get())
            explicit_vertical_lines_str = self.explicit_vertical_lines.get()
            explicit_horizontal_lines_str = self.explicit_horizontal_lines.get()

            explicit_vertical_lines = (
                [float(x.strip()) for x in explicit_vertical_lines_str.split("|") if x.strip()]
                if explicit_vertical_lines_str.strip()
                else []
            )
            explicit_horizontal_lines = (
                [float(x.strip()) for x in explicit_horizontal_lines_str.split("|") if x.strip()]
                if explicit_horizontal_lines_str.strip()
                else []
            )

            # 验证输入
            if vertical_strategy not in ["lines", "lines_strict", "text", "explicit"]:
                raise ValueError("垂直线策略无效")
            if horizontal_strategy not in ["lines", "lines_strict", "text", "explicit"]:
                raise ValueError("水平线策略无效")
            if not isinstance(snap_tolerance, int) or snap_tolerance < 0:
                raise ValueError("捕捉距离必须为非负整数")
            if not isinstance(join_tolerance, int) or join_tolerance < 0:
                raise ValueError("合并距离必须为非负整数")
            if not isinstance(text_tolerance, int) or text_tolerance < 0:
                raise ValueError("文本距离必须为非负整数")

            # 保存设置
            config.set_vertical_strategy(vertical_strategy)
            config.set_horizontal_strategy(horizontal_strategy)
            config.set_snap_tolerance(snap_tolerance)
            config.set_join_tolerance(join_tolerance)
            config.set_text_tolerance(text_tolerance)
            config.set_explicit_vertical_lines(explicit_vertical_lines)
            config.set_explicit_horizontal_lines(explicit_horizontal_lines)

            self.settings_window.destroy()
        except ValueError as e:
            messagebox.showerror("错误", f"保存设置失败：{e}")
        except Exception as e:
            messagebox.showerror("错误", f"未知错误：{e}")

    def reset_table_settings(self):
        """恢复默认设置"""
        config.reset_table_settings()
        settings = config.get_table_settings()
        # 更新控件的值
        self.vertical_strategy.set(settings["vertical_strategy"])
        self.horizontal_strategy.set(settings["horizontal_strategy"])
        self.explicit_vertical_lines.delete(0, tk.END)
        self.explicit_vertical_lines.insert(0, "|".join(map(str, settings["explicit_vertical_lines"])))
        self.explicit_horizontal_lines.delete(0, tk.END)
        self.explicit_horizontal_lines.insert(0, "|".join(map(str, settings["explicit_horizontal_lines"])))
        self.snap_tolerance.delete(0, tk.END)
        self.snap_tolerance.insert(0, settings["snap_tolerance"])
        self.join_tolerance.delete(0, tk.END)
        self.join_tolerance.insert(0, settings["join_tolerance"])
        self.text_tolerance.delete(0, tk.END)
        self.text_tolerance.insert(0, settings["text_tolerance"])

    def create_tooltip(self, widget, text):
        """为控件添加鼠标悬浮提示"""
        tooltip = tk.Toplevel(self.settings_window)
        tooltip.overrideredirect(True)  # 去掉标题栏
        tooltip.withdraw()  # 隐藏窗口
        tooltip_label = tk.Label(tooltip, text=text, wraplength=200, justify="left", bg="white", borderwidth=1, relief="solid")
        tooltip_label.pack()

        def show_tooltip(event):
            x, y = widget.winfo_pointerxy()
            tooltip.deiconify()
            tooltip.geometry(f"+{x + 10}+{y + 10}")

        def hide_tooltip(event):
            tooltip.withdraw()

        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)

    def on_closing(self):
        """处理窗体关闭事件"""
        if messagebox.askokcancel("确认", "是否保存设置？\n如不保存将恢复默认设置。"):
            self.save_settings()
        else:
            config.reset_table_settings()
        self.settings_window.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = WinForm(root)

    root.mainloop()