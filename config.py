import json
import os

default_table_settings = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "explicit_vertical_lines": [],
    "explicit_horizontal_lines": [],
    "snap_tolerance": 5,
    "join_tolerance": 5,
    "text_tolerance": 3
}

default_file_paths = {
    "PDF_path": "",
    "Excel_path": ""
}

table_settings_json = "./table_settings.json"
file_path_json = "./file_default_path.json"

def load_settings(file_path: str, default_settings: dict) -> dict:
    """
    加载设置文件，如果文件不存在则初始化为默认设置。
    """
    try:
        with open(file_path, "r") as file:
            return json.load(file)
    except FileNotFoundError:
        save_settings(file_path, default_settings)
        return default_settings
    except Exception as e:
        print(f"加载设置时发生错误：{e}")
        return default_settings

def save_settings(file_path: str, settings: dict):
    """
    保存设置到文件。
    """
    try:
        with open(file_path, "w") as file:
            json.dump(settings, file, indent=4)
    except Exception as e:
        print(f"保存设置时发生错误：{e}")

def get_table_settings() -> dict:
    """
    获取当前表格设置。
    """
    return load_settings(table_settings_json, default_table_settings)

def set_setting(key: str, value):
    """
    设置表格配置项。
    """
    table_settings = get_table_settings()
    table_settings[key] = value
    save_settings(table_settings_json, table_settings)

def reset_table_settings():
    """
    恢复默认表格设置。
    """
    save_settings(table_settings_json, default_table_settings)

def get_file_paths() -> dict:
    """
    获取文件路径设置。
    """
    return load_settings(file_path_json, default_file_paths)

def set_file_path(key: str, value: str):
    """
    设置文件路径。
    """
    file_paths = get_file_paths()
    file_paths[key] = value
    save_settings(file_path_json, file_paths)

# 动态设置表格配置项的函数
def set_vertical_strategy(strategy: str):
    set_setting("vertical_strategy", strategy)

def set_horizontal_strategy(strategy: str):
    set_setting("horizontal_strategy", strategy)

def set_explicit_vertical_lines(lines: list):
    set_setting("explicit_vertical_lines", lines)

def set_explicit_horizontal_lines(lines: list):
    set_setting("explicit_horizontal_lines", lines)

def set_snap_tolerance(tolerance: int):
    set_setting("snap_tolerance", tolerance)

def set_join_tolerance(tolerance: int):
    set_setting("join_tolerance", tolerance)

def set_text_tolerance(tolerance: int):
    set_setting("text_tolerance", tolerance)

# 动态设置文件路径的函数
def set_pdf_path(pdf_path: str):
    set_file_path("PDF_path", pdf_path)

def set_excel_path(excel_path: str):
    set_file_path("Excel_path", excel_path)