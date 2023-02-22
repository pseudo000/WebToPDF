#Step1 
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json
import time

#Step2 
def  PrintSetUp ():
     # 인쇄로 PDF 저장 설정
    chopt=webdriver.ChromeOptions()
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
         "isLandscapeEnabled" : True , #인쇄 방향 지정 ture로 가로, false로 세로. 
        "pageSize" : 'A4' , #용지 종류(A3, A4, A5, Legal, Letter, Tabloid 등) 
        #"mediaSize": {"height_microns": 355600, "width_microns": 215900}, #종이 크기 (10000 마이크로미터 = 1cm) 
        #"marginsType": 0, # 여백 유형 #0: 기본값 1: 여백 없음 2: 최소 
        #"scalingType": 3 , #0: 기본값 1: 페이지에 맞추기 2: 용지에 맞추기 3: 사용자 정의 
        #"scaling": "141" ,#배율 
        #"profile.managed_default_content_settings.images": 2,
        "isHeaderFooterEnabled": False,#헤더 및 바닥글 
        "isCssBackgroundEnabled" : True , #배경 그래픽 
        #"isDuplexEnabled": False, #양면인쇄 ture로 양면인쇄, false로 단면인쇄 
        #"isColorEnabled": True, #컬러인쇄 true로 칼라, false로 흑백 
        "isCollateEnabled": True #부 단위로 인쇄
    }
    
    prefs = { 'printing.print_preview_sticky_settings.appState' :
             json.dumps(appState),
             "download.default_directory" : "~/Downloads" 
             } #appState --> pref 
    chopt.add_experimental_option( 'prefs' , prefs) #prefs --> chopt 
    chopt.add_argument( '--kiosk-printing' ) # 인쇄 대화 상자 열면 인쇄 버튼을 무조건 누르십시오. 
    return chopt

#Step3 
def  main_WebToPDF (BlogURL) :
     #Web 페이지 또는 HTML 파일을 PDF로 Selenium을 사용하여 변환
    chopt = PrintSetUp()
    driver_path = "./chromedriver"  #webdriver 경로
    driver = webdriver.Chrome (executable_path=driver_path, options=chopt)
    driver.implicitly_wait( 10 ) # 초 암시적 대기 
    driver.get(BlogURL) #블로그 URL 로드 
    WebDriverWait(driver, 15 ).until(EC.presence_of_all_elements_located)   # 페이지의 모든 요소가 로드될 때까지 대기(15초 ) 에서 타임 아웃 판정) 
    driver.execute_script( 'return window.print()' ) #Print as PDF 
    time.sleep( 
    10 ) # 파일 다운로드를 위해 10초 대기
    
 #Step4 
if __name__ == '__main__' :
    BlogURLList=[ 'https://degitalization.hatenablog.jp/entry/2020/05/15/084033' , 
                  'https://note.com/makkynm/n/n1343f41c2fb7' ,
                 "file:///Users/makky /Documents/Python/Sample.html" ]
    for BlogURL in   BlogURLList:
        main_WebToPDF (BlogURL)