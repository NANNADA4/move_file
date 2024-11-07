"""
폴더를 순회하며 '파일명', '폴더명'열로부터 이름이 같은 파일 혹은 폴더가 발견되면 같은 행 '폴더경로' 파일경로로 복사합니다
"""


import os
import sys
import shutil
import pandas as pd


def process_copy():
    """폴더를 순회하며 복사를 시도합니다"""
    folder_path = input("입력폴더 경로를 입력하세요.\n=> ")
    excel_path = input("엑셀 파일 경로를 입력하세요(확장자 포함).\n=> ")

    # 엑셀 파일 읽기
    try:
        df = pd.read_excel(excel_path)
    except Exception as e:  # pylint: disable=W0718
        print(f"엑셀 파일을 읽는 중 오류가 발생했습니다: {e}")
        sys.exit()

    try:
        file_names = df['파일명'].tolist()
    except KeyError:
        print("'파일명' 열이 엑셀 파일에 없습니다. 파일 복사를 건너뜁니다")
        file_names = []

    try:
        folder_names = df['폴더명'].tolist()
    except KeyError:
        print("'폴더명' 열이 엑셀 파일에 없습니다. 폴더 복사를 건너뜁니다.")
        folder_names = []

    file_names = df['파일명'].dropna().tolist()
    folder_names = df['폴더명'].dropna().tolist()

    if file_names:
        process_file(folder_path, file_names, df)

    if folder_names:
        process_folder(folder_path, folder_names, df)


def process_folder(folder_path, folder_names, df):
    """엑셀 파일에서 폴더명을 찾아 복사합니다"""
    for root, dirs, _ in os.walk(folder_path):
        for dir_name in dirs:
            if dir_name in folder_names:
                row_index = folder_names.index(dir_name)
                target_folder = df['폴더경로'].iloc[row_index]

                source_folder_path = os.path.join(root, dir_name)
                destination_folder_path = os.path.join(target_folder, dir_name)

                try:
                    ensure_directory_exists(target_folder)
                    if not os.path.exists(destination_folder_path):
                        shutil.copytree(source_folder_path,
                                        destination_folder_path)
                        print(f"{source_folder_path} 폴더가 {
                              destination_folder_path}로 복사되었습니다.")
                except Exception as e:  # pylint: disable=W0718
                    print(f"폴더 복사 중 오류 발생: {e}")


def process_file(folder_path, file_names, df):
    """엑셀 파일에서 파일명을 찾아 복사합니다"""
    for root, _, files in os.walk(folder_path):
        for file in files:
            if file in file_names:
                row_index = file_names.index(file)
                target_folder = df['폴더경로'].iloc[row_index]

                source_file_path = os.path.join(root, file)
                destination_file_path = os.path.join(target_folder, file)

                try:
                    ensure_directory_exists(target_folder)
                    shutil.copy(source_file_path, destination_file_path)
                    print(f"{file} 파일이 {target_folder}로 복사되었습니다.")
                except Exception as e:  # pylint: disable=W0718
                    print(f"파일 복사 중 오류 발생: {e}")


def ensure_directory_exists(target_folder):
    """대상 폴더가 존재하는지 확인하고 없으면 생성합니다"""
    if not os.path.exists(target_folder):
        os.makedirs(target_folder)
        print(f"{target_folder} 폴더가 생성되었습니다.")


if __name__ == "__main__":
    process_copy()
