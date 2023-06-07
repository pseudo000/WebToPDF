import os
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfdevice import PDFDevice
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.layout import LAParams

def keep_first_page_only(file_path):
    output_path = "output.pdf"  # 남길 첫 번째 페이지를 저장할 파일 경로
    
    with open(file_path, 'rb') as file:
        parser = PDFParser(file)
        document = PDFDocument(parser)
        
        # 첫 번째 페이지만 남기고 나머지 페이지 삭제
        output = PDFDocument()
        output.set_strict(document.is_strict)
        output.set_version(document.version)
        output.add_page(next(document.get_pages()))
        
        # 결과 파일에 쓰기
        with open(output_path, 'wb') as output_file:
            output.write(output_file)
    
    return output_path

# PDF 파일이 있는 폴더의 경로를 정의합니다.
folder_path = 'C:\\Users\\swwoo\\Downloads'

# 폴더 내의 모든 파일을 반복합니다.
for filename in os.listdir(folder_path):
    # PDF 파일만 처리합니다.
    if filename.endswith(".pdf"):
        # 파일의 전체 경로를 생성합니다.
        file_path = os.path.join(folder_path, filename)
        
        # 첫 번째 페이지만 남기고 나머지 페이지 삭제
        output_file = keep_first_page_only(file_path)
        
        # 기존 파일을 삭제하고 남긴 첫 번째 페이지만 있는 파일로 대체
        os.remove(file_path)
        os.rename(output_file, file_path)
