import os
import re
from collections import defaultdict
from shutil import copyfile
from PyPDF2 import PdfMerger

# 병합할 폴더 경로와 결과물 파일이 저장될 폴더 경로 설정
folder_path = 'C:\\Users\\swwoo\\Downloads'
result_folder_path = 'C:\\import\\Merged'

if not os.path.exists(result_folder_path):
    os.makedirs(result_folder_path)

# 파일 이름의 앞 11자리가 동일한 파일들을 그룹화
file_dict = defaultdict(list)
for filename in os.listdir(folder_path):
    if filename.endswith('.pdf'):
        key = filename[:11]  # 앞 11자리까지만 추출하여 key로 사용
        file_dict[key].append(filename)

# 그룹화된 파일들을 순회하며 각 그룹을 병합 또는 복사
for key, filenames in file_dict.items():
    if len(filenames) == 1:
        file_path = os.path.join(folder_path, filenames[0])
        result_name = re.sub(r'\(\d+\)', '', filenames[0])  # 괄호와 괄호 안에 있는 숫자 삭제
        result_path = os.path.join(result_folder_path, result_name)
        copyfile(file_path, result_path)
    else:
        merger = PdfMerger()
        for filename in sorted(filenames):
            file_path = os.path.join(folder_path, filename)
            merger.append(file_path)
        result_name = filenames[0]
        # 파일 이름에서 괄호와 괄호 안에 있는 숫자를 삭제하여 결과물 파일 이름 생성
        result_name = re.sub(r'\(\d+\)', '', result_name)
        i = 1
        while os.path.exists(os.path.join(result_folder_path, result_name)):
            result_name = re.sub(r'\(\d+\)', '', result_name) + f'({i})'
            i += 1
        result_path = os.path.join(result_folder_path, result_name)
        merger.write(result_path)
        merger.close()
