from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import time
import openpyxl
import os
import shutil

 
def  PrintSetUp ():
    chopt=webdriver.ChromeOptions()
    chopt.add_experimental_option('excludeSwitches', ['enable-logging'])  #기능하지 않음 오류메시지 제거 위해 넣음
    appState = {
        "recentDestinations" : [
            {
                "id" : "Save as PDF" ,
                 "origin" : "local" ,
                 "account" : ""
            }
        ],
        "selectedDestinationId" : "Save as PDF" ,
         "version" : 2 ,
         "isLandscapeEnabled" : False , #인쇄 방향 지정 ture로 가로, false로 세로. 
        "pageSize" : 'A4' , #용지 종류(A3, A4, A5, Legal, Letter, Tabloid 등) paper size?
        #"mediaSize": {"height_microns": 355600, "width_microns": 215900}, #종이 크기 (10000 마이크로미터 = 1cm) 
        #"marginsType": 0, # 여백 유형 #0: 기본값 1: 여백 없음 2: 최소 
        "scalingType": 3 , #0: 기본값 1: 페이지에 맞추기 2: 용지에 맞추기 3: 사용자 정의 
        "scaling": "80" ,#배율 
        #"profile.managed_default_content_settings.images": 2,
        "isHeaderFooterEnabled": False,#헤더 및 바닥글 
        "isCssBackgroundEnabled" : True , #배경 그래픽 
        #"isDuplexEnabled": False, #양면인쇄 ture로 양면인쇄, false로 단면인쇄 
        "isColorEnabled": False, #컬러인쇄 true로 칼라, false로 흑백 
        "isCollateEnabled": True #부 단위로 인쇄
    }
    
    prefs = { 'printing.print_preview_sticky_settings.appState' :
             json.dumps(appState),
             "download.default_directory" : "~/Downloads" 
             } 
    chopt.add_experimental_option( 'prefs' , prefs) 
    chopt.add_argument( '--kiosk-printing' ) # 인쇄 대화 상자 열면 인쇄 버튼을 무조건 누름
    return chopt


def main_WebToPDF (BlogURL) :
    chopt = PrintSetUp()
    # driver_path = "./chromedriver"  #webdriver 경로
    driver = webdriver.Chrome (service=Service('path/to/chrome'), options=chopt)
    driver.implicitly_wait( 10 )
    driver.get(BlogURL)  
    WebDriverWait(driver, 15 ).until(EC.presence_of_all_elements_located)   # 페이지의 모든 요소가 로드될 때까지 대기(15초 ) 에서 타임 아웃 판정 
    driver.execute_script( 'return window.print()' ) 
    time.sleep( 10 ) 


xl_file_path = 'C:\\import\\HSCODE.xlsx' 

wb =openpyxl.load_workbook(xl_file_path, data_only=True)  
ws = wb.worksheets[0]  


for row, item in enumerate(ws.rows):
    if(row > 0):   

        url_path = item[1].value  
        file_name = item[0].value 
        # save_file_path = 'C:\\import\\' 
          
        if __name__ == '__main__' :
            BlogURLList=[url_path]
            try:
                for BlogURL in   BlogURLList:
                    main_WebToPDF (BlogURL)
            
                filename = max(['C:\\Users\\swwoo\\Downloads' + "\\" + f for f in os.listdir('C:\\Users\\swwoo\\Downloads')],key=os.path.getctime)
                shutil.move(filename,os.path.join('C:\\Users\\swwoo\\Downloads',str(file_name) +".pdf"))
                    
            except:
                print(BlogURL + '실패')

            # filename = max(['C:\\Users\\swwoo\\Downloads' + "\\" + f for f in os.listdir('C:\\Users\\swwoo\\Downloads')],key=os.path.getctime)
            # shutil.move(filename,os.path.join('C:\\Users\\swwoo\\Downloads',r"newfilename.ext"))