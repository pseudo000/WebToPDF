import pdfkit
# config = pdfkit.configuration(wkhtmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
options = {                                                 
    'page-size': 'A4',
    'margin-top': '0',
    'margin-right': '0',
    'margin-left': '0',
    'margin-bottom': '0',
    'zoom': '1.0',
    'encoding': "UTF-8",
    'enable-javascript': None,
    'javascript-delay': '3000',
    'images':None ,
    'enable-local-file-access':None, 
    'keep-relative-links':None           
    }  
import openpyxl 
import time


xl_file_path = 'C:\\import\\HSCODE.xlsx' 

wb =openpyxl.load_workbook(xl_file_path, data_only=True)  
ws = wb.worksheets[0]  


for row, item in enumerate(ws.rows):
    if(row > 0):   

        url_path = item[1].value  
        file_name = item[0].value+'.pdf' 
        save_file_path = 'C:\\import\\' 
        try:  
            pdfkit.from_url(url_path, file_name, options= options)  
            time.sleep(1)
        except:
            print(file_name + '실패')   