import os
import json
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import config
import PDFtoExcel

class WinForm:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Excel Converter")
        self.center_window(585, 260)  # 设置窗体大小并居中
        self.root.resizable(True, True)  # 允许调整大小
        self.root.minsize(width=480, height=180)
        # self.root.iconbitmap("app_icon.ico")  # 设置窗体图标

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
        self.create_tooltip(self.advanced_settings_button, "点击打开高级设置，可修改表格解析参数。")

        # 转换按钮
        tk.Button(self.root, text="开始转换", command=self.convert_pdf_to_excel).grid(row=2, column=2, padx=10, pady=10, sticky="se")

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

        # 调用 PDF 转 Excel 的功能
        result = PDFtoExcel.pdf_to_excel(pdf_path, excel_path)
        messagebox.showinfo("转换完成", result)

    def open_advanced_settings(self):
        """打开高级设置窗体"""
        self.advanced_settings_form = AdvancedSettingsForm(self.root)

    def create_tooltip(self, widget, text):
        """为控件添加鼠标悬浮提示"""
        tooltip = tk.Toplevel(self.root)
        tooltip.withdraw()  # 隐藏窗口
        tooltip_label = tk.Label(tooltip, text=text, wraplength=200, justify="left", bg="white", borderwidth=1,
                                 relief="solid")
        tooltip_label.pack()

        def show_tooltip(event):
            x, y = widget.winfo_pointerxy()
            tooltip.deiconify()
            tooltip.geometry(f"+{x + 15}+{y + 15}")

        def hide_tooltip(event):
            tooltip.withdraw()

        widget.bind("<Enter>", show_tooltip)
        widget.bind("<Leave>", hide_tooltip)


class AdvancedSettingsForm:
    def __init__(self, parent):
        self.parent = parent
        self.settings_window = tk.Toplevel(self.parent)
        self.settings_window.title("高级设置")
        self.settings_window.geometry("200x400")
        self.settings_window.resizable(False, False)
        self.settings_window.transient(self.parent)  # 设置为父窗体的子窗体
        self.settings_window.grab_set()  # 独占焦点
        self.center_window(self.settings_window, 200, 400)  # 居中显示
        self.create_widgets()

    def center_window(self, window, width=200, height=400):
        screen_width = window.winfo_screenwidth()
        screen_height = window.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    def create_widgets(self):
        settings = config.get_table_settings()

        ttk.Label(self.settings_window, text="垂直线策略:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
        self.vertical_strategy = ttk.Combobox(self.settings_window, values=["lines", "lines_strict", "text", "explicit"])
        self.vertical_strategy.set(settings["vertical_strategy"])
        self.vertical_strategy.grid(row=0, column=1, padx=10, pady=10, sticky="we")

        ttk.Label(self.settings_window, text="水平线策略:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
        self.horizontal_strategy = ttk.Combobox(self.settings_window, values=["lines", "lines_strict", "text", "explicit"])
        self.horizontal_strategy.set(settings["horizontal_strategy"])
        self.horizontal_strategy.grid(row=1, column=1, padx=10, pady=10, sticky="we")

        ttk.Label(self.settings_window, text="捕捉距离:").grid(row=2, column=0, padx=10, pady=10, sticky="e")
        self.snap_tolerance = ttk.Entry(self.settings_window)
        self.snap_tolerance.insert(0, settings["snap_tolerance"])
        self.snap_tolerance.grid(row=2, column=1, padx=10, pady=10, sticky="we")

        ttk.Label(self.settings_window, text="合并距离:").grid(row=3, column=0, padx=10, pady=10, sticky="e")
        self.join_tolerance = ttk.Entry(self.settings_window)
        self.join_tolerance.insert(0, settings["join_tolerance"])
        self.join_tolerance.grid(row=3, column=1, padx=10, pady=10, sticky="we")

        ttk.Label(self.settings_window, text="文本距离:").grid(row=4, column=0, padx=10, pady=10, sticky="e")
        self.text_tolerance = ttk.Entry(self.settings_window)
        self.text_tolerance.insert(0, settings["text_tolerance"])
        self.text_tolerance.grid(row=4, column=1, padx=10, pady=10, sticky="we")

        ttk.Label(self.settings_window, text="保留空白字符:").grid(row=5, column=0, padx=10, pady=10, sticky="e")
        self.keep_blank_chars = tk.BooleanVar(value=settings["keep_blank_chars"])
        ttk.Checkbutton(self.settings_window, variable=self.keep_blank_chars).grid(row=5, column=1, padx=10, pady=10, sticky="w")

        ttk.Button(self.settings_window, text="保存设置", command=self.save_settings).grid(row=6, column=0, columnspan=2, pady=10)

    def save_settings(self):
        config.set_vertical_strategy(self.vertical_strategy.get())
        config.set_horizontal_strategy(self.horizontal_strategy.get())
        config.set_snap_tolerance(int(self.snap_tolerance.get()))
        config.set_join_tolerance(int(self.join_tolerance.get()))
        config.set_text_tolerance(int(self.text_tolerance.get()))
        config.set_keep_blank_chars(self.keep_blank_chars.get())
        messagebox.showinfo("提示", "高级设置已保存！")
        self.settings_window.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = WinForm(root)
    root.mainloop()