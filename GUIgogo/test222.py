import os

def check_pdf_excel():
    pdf_folder_path = 'C:\\Users\\swwoo\\Downloads'

    pdf_dict = {}
    for pdf in os.listdir(pdf_folder_path):
        if pdf.endswith(".pdf"):
            pdf_name = os.path.splitext(pdf)[0] 
            pdf_path = os.path.join(pdf_folder_path, pdf)
            pdf_dict[pdf_name] = pdf_path
            print(pdf_dict)


check_pdf_excel()




