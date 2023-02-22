import os
import PyPDF2

# PDF 파일이 있는 폴더의 경로를 정의합니다.
folder_path = 'C:\\Users\\swwoo\\Downloads'

# 폴더 내의 모든 파일을 반복합니다.
for filename in os.listdir(folder_path):
    # PDF 파일만 처리합니다.
    if filename.endswith(".pdf"):
        # 파일의 전체 경로를 생성합니다.
        file_path = os.path.join(folder_path, filename)
        
        # PyPDF2를 사용하여 PDF 파일을 읽습니다.
        pdf_file = PyPDF2.PdfReader(file_path)
        
        # 수정된 버전을 저장할 새 PDF 파일을 생성합니다.
        output_pdf = PyPDF2.PdfWriter()
        
        # 새 PDF 파일에 1페이지를 추가합니다.
        output_pdf.add_page(pdf_file.pages[0])
       
        # 새 PDF 파일을 디스크에 쓰기
        with open(file_path, "wb") as output_file:
            output_pdf.write(output_file)