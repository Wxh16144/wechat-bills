import datetime
import os
import pathlib
import pdfplumber
import openpyxl
import argparse
import sys

def convert_pdf(files, output_file):
    try:
        total_rows = []
        for fi, file in enumerate(files):
            print(f"正在处理: {file}")
            with pdfplumber.open(file) as pdf:
                for pi, page in enumerate(pdf.pages):
                    tables = page.extract_tables()
                    for ti, table in enumerate(tables):
                        if fi != 0 and pi == 0 and ti == 0:
                            table = table[3:]  # 去除重复表头
                        for row in table:
                            if len(row) > 1:
                                row[1] = row[1].replace("\n", " ") if row[1] and "\n" in row[1] else row[1]
                                row = [col.replace("\n", "") if col else "" for col in row]
                                total_rows.append(row)
        if len(total_rows) < 4:
            print("警告: 文件中未找到表格")
            return

        header = total_rows[2]
        total_rows = total_rows[3:]
        total_rows.sort(key=lambda ele: ele[1])  # 按日期排序
        wb = openpyxl.Workbook()
        sheet = wb.active
        sheet.title = "账单"
        for ci, col in enumerate(header):  # 写表头
            sheet.cell(row=1, column=ci + 1, value=col)
        for ri, row in enumerate(total_rows):
            for ci, col in enumerate(row):
                if ci == 1:  # 日期转格式
                    try:
                        sheet.cell(row=ri + 2, column=ci + 1, value=datetime.datetime.fromisoformat(col))
                        sheet.cell(row=ri + 2, column=ci + 1).number_format = 'yyyy-mm-dd hh:mm:ss'
                    except:
                        pass
                elif ci == 5:  # 金额转小数
                    try:
                        sheet.cell(row=ri + 2, column=ci + 1, value=float(col))
                        sheet.cell(row=ri + 2, column=ci + 1).number_format = '0.00'  # 保留两位小数
                    except:
                        pass
                else:
                    sheet.cell(row=ri + 2, column=ci + 1, value=col)
        
        # 确保输出目录存在
        os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
        
        wb.save(output_file)
        print(f"转换成功\n共{len(total_rows)}条记录\n已保存至{output_file}")
    except Exception as e:
        print(f"转换失败:\n{e}")

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="微信账单PDF转Excel工具")
    parser.add_argument("input", help="输入文件或目录路径")
    parser.add_argument("output", nargs="?", help="输出Excel文件路径 (可选)")
    
    args = parser.parse_args()
    
    input_path = args.input
    output_path = args.output
    
    files = []
    if os.path.isfile(input_path):
        files.append(input_path)
    elif os.path.isdir(input_path):
        # 查找目录下所有PDF
        files = [os.path.join(input_path, f.name) for f in (pathlib.Path(input_path).glob('*.pdf'))]
    else:
        print(f"错误: 找不到输入路径 '{input_path}'")
        sys.exit(1)
        
    if not files:
        print("未找到PDF文件")
        sys.exit(1)
        
    # 确定输出路径
    if not output_path:
        # 如果未指定输出，在输入所在目录生成
        base_dir = os.path.dirname(files[0])
        output_path = os.path.join(base_dir, f"convert_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
    elif os.path.isdir(output_path):
        # 如果指定的是目录，生成文件名
        output_path = os.path.join(output_path, f"convert_{datetime.datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx")
        
    convert_pdf(files, output_path)
