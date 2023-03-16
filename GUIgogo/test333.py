import os
from openpyxl import load_workbook


def check_pdf_excel():
    excel_file_path = 'C:\\import\\HSCODE.xlsx'

    workbook = load_workbook(filename=excel_file_path)
    worksheet = workbook.active

    pdf_files = []
    for cell in worksheet['A'][1:]:
        count = 0
        filename = f"{os.path.splitext(cell.value)[0]} ({count+1}){os.path.splitext(cell.value)[1]}"
        while filename in pdf_files:
            count += 1
            filename = f"{os.path.splitext(cell.value)[0]} ({count+1}){os.path.splitext(cell.value)[1]}"
        pdf_files.append(filename)
    print(pdf_files)


check_pdf_excel()
