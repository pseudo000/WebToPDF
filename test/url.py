# wget , openpyxl , time 을 import 
import wget # 다운로드
import openpyxl # 엑셀 읽고 쓰기
import time


xl_file_path = 'C:\\import\\HSCODE.xlsx' # 엑셀 경로

wb =openpyxl.load_workbook(xl_file_path, data_only=True)  #엑셀파일을 열고 읽음
ws = wb.worksheets[0]  # 몇 번째 시트에 해당 내용이 있는지 번호기재. 

# 읽은 엑셀을 루프를 돌면서 다운로드
for row, item in enumerate(ws.rows):
    if(row > 0):   # 첫행이 제목이라면 다음행부터

        url_path = item[1].value  # 1번쨰 열 다운로드 할 url 경로
        file_name = item[0].value+'.jpg' #0번쨰 '사진번호'를 파일이름으로 확장자까지 적어줌
        save_file_path = 'C:\\import\\' # 저장경로
        try:   #다운로드하다 사진이 없어 죽을 경우를 대비해 예외처리함
            wget.download(url_path, save_file_path+file_name)  # 이미지가 있는 주소와 파일 경로 + 저장할 이름

            time.sleep(0.5)  # 몇 만건이라면 디도스 공격이 될 테니 천천히 받도록 쉬는시간 넣음
        except:
            print(file_name + '실패')   # 저장 못한 경우 프린트 되도록