import pdfplumber
import xlsxwriter
import multiprocessing
import time

def process_page_range(pdf_path, start_page, end_page):
    """
    处理指定页面范围的任务。
    """
    all_data = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_num in range(start_page, end_page):
            page = pdf.pages[page_num]
            text = page.extract_text()
            textdata = []
            tabledata = page.extract_table()
            if text:
                lines = text.split("\n")
                for line in lines:
                    line = line.split(" ")
                    textdata.append(line)

            j = 0
            for i in range(len(textdata)):
                if textdata[i][0] != tabledata[j][0]:
                    all_data.append(textdata[i])
                else:
                    all_data.append(tabledata[j])
                    j += 1
    return all_data


def pdf_to_excel(pdf_path, excel_path):
    """
    使用多进程将 PDF 文件中的表格内容提取并保存为 Excel 文件。
    """
    try:
        with pdfplumber.open(pdf_path) as pdf:
            total_pages = len(pdf.pages)
            pool_size = max(1, multiprocessing.cpu_count() // 2)  # 设置进程池大小为CPU核数的一半
            pages_per_process = total_pages // pool_size + 1  # 每个进程处理的页面数

            # 创建进程池并分配任务
            with multiprocessing.Pool(processes=pool_size) as pool:
                tasks = [(pdf_path, i * pages_per_process, min((i + 1) * pages_per_process, total_pages)) for i in
                         range(pool_size)]
                results = pool.starmap(process_page_range, tasks)

        # 合并所有进程的结果
        all_data = [item for sublist in results for item in sublist]

        # 保存到 Excel 文件
        save_to_excel(all_data, excel_path)
        return f"数据已成功提取并保存到 {excel_path}"

    except Exception as e:
        return f"处理 PDF 文件时发生错误：{e}"


def save_to_excel(data, excel_path):
    """
    将数据保存为 Excel 文件。

    参数:
        data (list): 合并后的数据。
        excel_path (str): 保存 Excel 文件的路径。
    """
    workbook = xlsxwriter.Workbook(excel_path)
    worksheet = workbook.add_worksheet()

    for row_idx, row_data in enumerate(data):
        for col_idx, value in enumerate(row_data):
            worksheet.write(row_idx, col_idx, value)

    workbook.close()

def get_lines(pdf_path):
    """
    获取当前 PDF 中默认识别到的水平线和垂直线位置

    参数：
        pdf_path: PDF文件路径
    """
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        table_finder = page.debug_tablefinder(table_settings={
            "vertical_strategy": "lines",
            "horizontal_strategy": "lines"
        })

        # 获取所有检测到的线条
        edges = table_finder.edges

        # 筛选出垂直线和水平线
        vertical_lines = [edge for edge in edges if edge["orientation"] == "v"]
        horizontal_lines = [edge for edge in edges if edge["orientation"] == "h"]

    return vertical_lines, horizontal_lines

if __name__ == "__main__":
    pdf_path = "test.pdf"
    excel_path = "temp.xlsx"
    pdf_to_excel(pdf_path, excel_path)