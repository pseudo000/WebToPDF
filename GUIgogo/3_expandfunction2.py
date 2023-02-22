import tkinter.ttk as ttk   
from tkinter import *  #__all__? filedialog는 서브 모듈이기 때문에 별도 임포트 해줘야 됨 
from tkinter import filedialog
import json
import time
import openpyxl
import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

root = Tk()
root.title("SP IMPORT")
root.resizable(False,False)
root.geometry("500x600")



#파일추가
def add_file():
    files = filedialog.askopenfilename(title="xlsx 파일을 선택하세요", \
        filetypes=(("xlsx 파일", "*.xlsx"),("All files", "*.*")), \
        initialdir="C:/")
    #사용자가 선택한 파일출력
    for file in files:
        txt_xls_path.insert(END, file)
    return files[0]

#MergedPDF저장경로
def savemerged_path():
    savemerged_selected = filedialog.askdirectory()
    if savemerged_selected == '': #사용자가 취소를 누를 때
        return
    txt_mergedpdf_path.delete(0, END)
    txt_mergedpdf_path.insert(0,savemerged_selected)


# print WebtoPDF
def PrintSetUp():
    chrome_options = Options()
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    app_state = {
        "recentDestinations": [
            {
                "id": "Save as PDF",
                "origin": "local",
                "account": ""
            }
        ],
        "selectedDestinationId": "Save as PDF",
        "version": 2,
        "isLandscapeEnabled": False,
        "pageSize": 'A4',
        "scalingType": 3,
        "scaling": "50",
        "isHeaderFooterEnabled": False,
        "isCssBackgroundEnabled": True,
        "isColorEnabled": False,
        "isCollateEnabled": True
    }
    prefs = {
        'printing.print_preview_sticky_settings.appState': json.dumps(app_state),
        "download.default_directory": "C:/Users/swwoo/Downloads" # 원하는 다운로드 경로로 변경
    }
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    return chrome_options


def WebToPDF(url):
    chrome_options = PrintSetUp()
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(12)
    driver.get(url)
    wait = WebDriverWait(driver, 17)
    wait.until(EC.presence_of_all_elements_located)
    time.sleep(2)
    driver.execute_script('return window.print()')
    time.sleep(3)
    driver.quit()


def start_prcess():
    xl_file_path = txt_xls_path.get()
    # xl_file_path = files    #'C:\\import\\HSCODE.xlsx'
    workbook = openpyxl.load_workbook(xl_file_path, data_only=True)
    worksheet = workbook.worksheets[0]

    for row, item in enumerate(worksheet.rows):
        if row > 0:
            url = item[1].value
            file_name = item[0].value

            try:
                WebToPDF(url)
        
                file_base = str(file_name) + ".pdf"
                file_count =+ 1
                filename = os.path.join('C:\\Users\\swwoo\\Downloads', f"{file_name} ({file_count}).pdf")
        
                while os.path.exists(filename):
                    file_count += 1
                    filename = os.path.join('C:\\Users\\swwoo\\Downloads', f"{file_name} ({file_count}).pdf")
            
                shutil.move(max(['C:\\Users\\swwoo\\Downloads' + "\\" + f for f in os.listdir('C:\\Users\\swwoo\\Downloads')], key=os.path.getctime), filename)

            except Exception as e:
                # print(f'Error processing URL {url}: {e}')
                print(url + '실패')





#Web to PDF 저장
savepdf_frame = LabelFrame(root, text="Web to PDF")
savepdf_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_xls_path = Entry(savepdf_frame)
txt_xls_path.pack(side="left", fill="x", expand="True", ipady=4) #

btn_start = Button(savepdf_frame, text="START",padx=5, pady=5, width=12, command=start_prcess)
btn_start.pack(side="right")

btn_savepdf_path = Button(savepdf_frame, text="...xlsx file",padx=5, pady=5, width=12, command=add_file)
btn_savepdf_path.pack(side="right")



#Slit PDF
splitpdf_frame = LabelFrame(root, text="Split PDF")
splitpdf_frame.pack(fill="x",padx=5, pady=5, ipady=5)

btn_splitpdf = Button(splitpdf_frame, text="START", padx=5, pady=5, width="20")
btn_splitpdf.pack()



# Merged PDF 저장경로
mergedpdf_frame = LabelFrame(root, text="Folder to merfed PDF")
mergedpdf_frame.pack(fill="x", padx=5, pady=5, ipady=5)

txt_mergedpdf_path = Entry(mergedpdf_frame)
txt_mergedpdf_path.pack(side="left", fill="x", expand="True", ipady=4) #

btn_mergestart = Button(mergedpdf_frame, text="START",padx=5, pady=5, width=12)
btn_mergestart.pack(side="right")

btn_mergedpdf_path = Button(mergedpdf_frame, text="...Folder to save",padx=5, pady=5, width=12, command=savemerged_path)
btn_mergedpdf_path.pack(side="right")





# 진행상황 profress bar
frame_progress = LabelFrame(root, text="진행상황")
frame_progress.pack(fill="x")



# 종료 프레임
close_frame = Frame(root)
close_frame.pack(side="right", padx=5, pady=5)

btn_close = Button(close_frame, padx=5, pady=5, text="CLOSE", width=12)
btn_close.pack()


p_var = DoubleVar()
progress_bar = ttk.Progressbar(frame_progress, maximum=100, variable=p_var)
progress_bar.pack(fill="x")





root.mainloop()

