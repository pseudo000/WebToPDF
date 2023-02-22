import argparse
import json
import time
import openpyxl
import os
import shutil
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


def PrintSetUp(download_dir):
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
        "download.default_directory": download_dir
    }
    chrome_options.add_experimental_option('prefs', prefs)
    chrome_options.add_argument('--kiosk-printing')
    return chrome_options


def WebToPDF(url, download_dir, wait_time):
    chrome_options = PrintSetUp(download_dir)
    driver = webdriver.Chrome(options=chrome_options)
    driver.implicitly_wait(12)
    driver.get(url)
    WebDriverWait(driver, 17).until(EC.presence_of_all_elements_located)
    driver.execute_script('return window.print()')
    # Wait for print preview to open
    time.sleep(2)
    while True:
        try:
            driver.execute_script('return window.frames["print-preview"].document.getElementById("preview-destination-section").textContent')
            break
        except:
            pass
        time.sleep(1)
    time.sleep(wait_time)
    driver.quit()


def main(xl_file_path, download_dir, wait_time):
    workbook = openpyxl.load_workbook(xl_file_path, data_only=True)
    worksheet = workbook.worksheets[0]

    for row, item in enumerate(worksheet.rows):
        if row > 0:
            url = item[1].value
            file_name = item[0].value
            file_count = 1

            try:
                WebToPDF(url, download_dir, wait_time)

                file_base = str(file_name) + ".pdf"
                file_count =+ 1
                filename = os.path.join(download_dir, f"{file_name} ({file_count}).pdf")

                while os.path.exists(filename):
                    file_count += 1
                    filename = os.path.join(download_dir, f"{file_name} ({file_count}).pdf")

                shutil.move(max([os.path.join(download_dir, f) for f in os.listdir(download_dir)], key=os.path.getctime), filename)

            except Exception as e:
                print(f'Error processing URL {url}: {e}')
                print(url + '실패')


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('-f', '--file', required=True, help='Path to the Excel file')
    parser.add_argument('-d', '--dir', required=True, help='Path to the download directory')
    parser.add_argument('-w', '--wait', type=int, default=12, help='Time in seconds to wait for print preview to load')
    args = parser.parse_args()

    main(args.file, args.dir, args.wait)



# python 파일이름.py -f 엑셀파일경로 -d 다운로드폴더경로 -w 대기시간(초)
# python 1-1printpdf2.py -f C:\\import\\HSCODE.xlsx -d C:\\Users\\swwoo\\Downloads -w 12
# 명령어 요소가 어려움

