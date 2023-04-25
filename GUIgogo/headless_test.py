### 헤들리스 테스트1 png파일까지 출력 완료

from selenium import webdriver
from PIL import Image
import io

DRIVER_PATH = 'chromedriver.exe'

# headless 모드로 Chrome 실행
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')

driver = webdriver.Chrome(DRIVER_PATH, options=options)
driver.get('http://m.helloneon.kr/product/에스더버니기대어-있는-에스더-83m-x-84cm/94/category/43/display/1/')

# 브라우저 창의 높이를 웹 페이지 높이에 맞게 조정
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
width = driver.execute_script("return Math.max(document.body.scrollWidth, document.body.offsetWidth, document.documentElement.clientWidth, document.documentElement.scrollWidth, document.documentElement.offsetWidth);")
height = driver.execute_script("return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")

driver.set_window_size(width, height)

# 스크린샷 이미지 저장
png = driver.get_screenshot_as_png()
im = Image.open(io.BytesIO(png)).convert('RGB')
im.save('screenshot.png')


driver.quit()