import os
import shutil
import pandas as pd

folder_path = input("폴더 경로를 입력하세요: ")
excel_path = input("엑셀 파일 경로를 입력하세요: ")

try:
    df = pd.read_excel(excel_path)
except Exception as e:
    print(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
    exit()

file_names = df['파일명'].tolist()

for root, dirs, files in os.walk(folder_path):
    for file in files:
        if file in file_names:
            row_index = file_names.index(file)
            target_folder = df['폴더경로'].iloc[row_index]

            if not os.path.exists(target_folder):
                try:
                    os.makedirs(target_folder)
                    print(f"{target_folder} 폴더가 생성되었습니다.")
                except Exception as e:
                    print(f"폴더 생성 중 오류 발생: {e}")
                    continue

            source_file_path = os.path.join(root, file)
            destination_file_path = os.path.join(target_folder, file)

            try:
                shutil.move(source_file_path, destination_file_path)
                print(f"{file} 파일이 {target_folder}로 이동되었습니다.")
            except Exception as e:
                print(f"파일 이동 중 오류 발생: {e}")
