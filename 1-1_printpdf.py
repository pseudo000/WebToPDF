import json
import time
import openpyxl
import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

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
        "download.default_directory": "~/Downloads"
    }
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    return chrome_options

def WebToPDF(url):
    chrome_options = PrintSetUp()
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(12)
    driver.get(url)
    WebDriverWait(driver, 17).until(EC.presence_of_all_elements_located)
    driver.execute_script('return window.print()')
    time.sleep(12)
    driver.quit()

xl_file_path = 'C:\\import\\HSCODE.xlsx'    

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
            filename = os.path.join('C:\\Users\\pseudo\\Downloads', f"{file_name} ({file_count}).pdf")
    
            while os.path.exists(filename):
                file_count += 1
                filename = os.path.join('C:\\Users\\pseudo\\Downloads', f"{file_name} ({file_count}).pdf")
        
            shutil.move(max(['C:\\Users\\pseudo\\Downloads' + "\\" + f for f in os.listdir('C:\\Users\\pseudo\\Downloads')], key=os.path.getctime), filename)

        except Exception as e:
            # print(f'Error processing URL {url}: {e}')
            print(url + '실패')