import os
from openpyxl import load_workbook


def check_pdf_excel():
    # # PDF 파일이 있는 폴더 경로
    # pdf_folder_path = 'C:\\Users\\swwoo\\Downloads'

    # 엑셀 파일 경로
    excel_file_path = 'C:\\import\\HSCODE.xlsx'

    # # PDF 파일 이름과 경로를 딕셔너리로 저장
    # pdf_dict = {}
    # for pdf in os.listdir(pdf_folder_path):
    #     if pdf.endswith(".pdf"):
    #         pdf_name = os.path.splitext(pdf)[0].split("(")[0].strip()  # 파일 이름에서 (1), (2) 등의 부분 제외
    #         pdf_path = os.path.join(pdf_folder_path, pdf)
    #         pdf_dict[pdf_name] = pdf_path

    # 엑셀 파일 열기
    workbook = load_workbook(filename=excel_file_path)
    worksheet = workbook.active

    pdf_files = []
    for cell in worksheet['A'][1:]:
        filename = cell.value
        count =+ 1
        filename = f"{os.path.splitext(cell.value)[0]} ({count}){os.path.splitext(cell.value)[1]}"
        while filename in pdf_files:
            count += 1
            filename = f"{os.path.splitext(cell.value)[0]} ({count}){os.path.splitext(cell.value)[1]}"
        pdf_files.append(filename)
    print(pdf_files)

check_pdf_excel()  