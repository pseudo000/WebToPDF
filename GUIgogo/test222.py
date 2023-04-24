# 셀레니움으로 png 출력, 배율지정 후 pdf 변환, 이후 png원본 삭제

from selenium import webdriver
from PIL import Image
import io
import os

DRIVER_PATH = 'chromedriver.exe'

# headless 모드로 Chrome 실행
options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--no-sandbox')

driver = webdriver.Chrome(DRIVER_PATH, options=options)
driver.get('https://smartstore.naver.com/pica_jp/products/6641532544')

# 브라우저 창의 높이를 웹 페이지 높이에 맞게 조정
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

# 스크린샷 이미지 크기 조정
width = driver.execute_script("return Math.max(document.body.scrollWidth, document.body.offsetWidth, document.documentElement.clientWidth, document.documentElement.scrollWidth, document.documentElement.offsetWidth);") 
height = driver.execute_script("return Math.max(document.body.scrollHeight, document.body.offsetHeight, document.documentElement.clientHeight, document.documentElement.scrollHeight, document.documentElement.offsetHeight);")
driver.set_window_size(width, height)

# 스크린샷 이미지 저장
png = driver.get_screenshot_as_png()
im = Image.open(io.BytesIO(png)).convert('RGB')
im.save('screenshot.png')

# 스크린샷 이미지를 A4 크기로 나누어 PDF 파일 생성
scale = 4.5
a4_width = int(297 * scale)  # A4용지 넓이 (픽셀)
a4_height = int(420 * scale)  # A4용지 높이 (픽셀)

im = Image.open('screenshot.png')
im_width, im_height = im.size

pages = []
for y in range(0, im_height, a4_height):
    for x in range(0, im_width, a4_width):
        box = (x, y, x+a4_width, y+a4_height)
        pages.append(im.crop(box))

pdf_path = 'output.pdf'
pages[0].save(pdf_path, save_all=True, append_images=pages[1:])
print(f"PDF 파일 {pdf_path}이 생성되었습니다.")

driver.quit()
# os.remove('screenshot.png')