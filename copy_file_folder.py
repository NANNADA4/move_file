"""
main 함수
"""

from module.process_folder import process_copy


def main():
    """main 함수"""
    print("="*60)
    print("\n>>>>>>파일명 기반 파일복사<<<<<<\n")
    print("***폴더 복사시 '폴더명'열, 파일 복사시 '파일명'열 필요***\n")
    print("="*60)

    print("[1] 파일 혹은 폴더 복사하기")
    print("[0] 프로그램 종료")
    input_num = input("=> ")

    match input_num:
        case '1':
            process_copy()
        case '0':
            return 0
        case _:
            print("\n====올바른 숫자를 입력해주세요====\n")

    return main()


if __name__ == "__main__":
    main()
