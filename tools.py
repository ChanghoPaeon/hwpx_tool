import os
import hashlib
import sys
import time
import shutil
from pathlib import Path

# 예시로 사용할 함수
# @benchmark
# def sample_function():
#     # 테스트용 함수 (예: 1부터 1000000까지 합산)
#     total = 0
#     for i in range(1, 1000001):
#         total += i
#     return total
#
#
# # 함수 호출 (실제 벤치마크가 실행됨)
# sample_function()





def benchmark(func):
    """
    주어진 함수의 실행 시간을 측정하고, 콘솔 및 파일에 로그를 남깁니다.
    """

    def wrapper(*args, **kwargs):
        start_time = time.time()  # 함수 실행 전 시간 기록
        result = func(*args, **kwargs)
        end_time = time.time()  # 함수 실행 후 시간 기록
        execution_time = end_time - start_time  # 실행 시간 계산

        # 콘솔에 로그 출력
        print(f"[BENCHMARK] {func.__name__} 실행 시간: {execution_time:.6f} 초")

        # 파일에 로그 기록
        # with open("benchmark_log.txt", "a") as log_file:
        #     log_file.write(f"[BENCHMARK] {func.__name__} 실행 시간: {execution_time:.6f} 초\n")

        return result

    return wrapper





# ✅ 실행 예제: 원드라이브 문서 폴더 내 중복 파일 삭제
# onedrive_path = os.path.expanduser("~/OneDrive")  # 기본 원드라이브 경로
# target_folder = os.path.join(onedrive_path, "문서")  # 검사할 폴더
#
# find_and_delete_duplicates(target_folder)

def get_file_hash(file_path):
    """파일의 SHA-256 해시 값을 계산합니다."""
    hasher = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            while chunk := f.read(4096):
                hasher.update(chunk)
    except Exception as e:
        print(f"해시 계산 중 오류 발생: {file_path}, 오류: {e}")
        return None
    return hasher.hexdigest()


def find_and_delete_duplicates(directory):
    """
    특정 경로 내에서 동일한 파일을 찾아 중복된 파일을 삭제합니다.

    :param directory: 탐색할 폴더 경로
    """
    if not os.path.exists(directory):
        print(f"경로가 존재하지 않습니다: {directory}")
        return False

    file_hashes = {}  # {해시값: 파일경로 리스트}

    for root, _, files in os.walk(directory):

        total_files = len(files)

        for i, file in enumerate(files, start=1):
            file_path = os.path.join(root, file)
            # print("before call hash: " + file_path)
            file_hash = get_file_hash(file_path)

            if file_hash:
                if file_hash in file_hashes:
                    file_hashes[file_hash].append(file_path)
                else:
                    file_hashes[file_hash] = [file_path]

            sys.stdout.write(f"\r🔄 hash 진행 중: {i}/{total_files} ({(i / total_files) * 100:.2f}%)")
            sys.stdout.flush()
        print(" ... end!")

    # 중복된 파일 삭제 (하나만 남기고 나머지 삭제)
    for hash_value, file_list in file_hashes.items():
        if len(file_list) > 1:
            print(f"\n🔹 중복 파일 발견 (해시: {hash_value[:10]}...):")
            for i, f in enumerate(file_list):
                print(f"  {i + 1}. {f}")

            # 첫 번째 파일을 남기고 나머지는 삭제
            for duplicate in file_list[1:]:
                try:
                    # os.remove(duplicate)
                    print(f"🗑️ 삭제 완료: {duplicate}")
                except Exception as e:
                    print(f"❌ 삭제 실패: {duplicate}, 오류: {e}")

        # sys.stdout.write(f"\r🔄 hash 진행 중: {i}/{total_files} ({(i / total_files) * 100:.2f}%)")
        # sys.stdout.flush()

    print("\n✅ 중복 파일 정리 완료!")
    return True


import os


def traverse_directory(path):
    """지정된 폴더 내의 모든 폴더 및 파일을 순회"""
    for root, dirs, files in os.walk(path):
        # root: 현재 디렉토리 경로
        # dirs: 현재 디렉토리 내의 하위 폴더 리스트
        # files: 현재 디렉토리 내의 파일 리스트

        print(f"현재 폴더: {root}")

        if dirs:
            print("하위 폴더들:")
            for dir_name in dirs:
                print(f"  {dir_name}")

        if files:
            print("파일들:")
            for file_name in files:
                print(f"  {file_name}")

        print("=" * 50)


# 예시 실행
# folder_path = "I:\test\NGD"  # 검색할 폴더 경로
# traverse_directory(folder_path)

import zipfile
# import os


# 예시 실행
# zip_file_path = "example.zip"  # ZIP 파일 경로
# output_folder = "extracted_files"  # 압축 풀 폴더
# extension = ".txt"  # 추출할 확장자 (예: .txt)

# 폴더가 없다면 생성
# if not os.path.exists(output_folder):
#     os.makedirs(output_folder)
#
# # 특정 확장자 파일만 추출
# extract_files_with_extension(zip_file_path, output_folder, extension)


def extract_files_with_extension(zip_path, extract_to_folder, ext):
    """
    ZIP 파일에서 특정 확장자를 가진 파일만 추출하는 함수
    :param zip_path: ZIP 파일 경로
    :param extract_to_folder: 압축을 풀 폴더 경로
    :param ext: 추출할 파일의 확장자 (예: '.txt', '.jpg')
    """
    # ZIP 파일 열기
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        # ZIP 파일 내의 모든 파일 목록 가져오기
        all_files = zip_ref.namelist()

        # 특정 확장자 파일만 필터링
        files_to_extract = [f for f in all_files if f.endswith(ext)]

        # 선택된 파일들만 추출
        for file in files_to_extract:
            zip_ref.extract(file, extract_to_folder)
            print(f"추출된 파일: {file}")


import os
import zipfile



def extract_hwp_from_zip(zip_path):
    # ZIP 파일 이름(확장자 제외)으로 폴더명 생성
    # print(os.path.basename(zip_path))
    folder_name = os.path.splitext(os.path.basename(zip_path))[0]
    # print(folder_name)
    extract_path = os.path.join(os.path.dirname(zip_path), folder_name)

    # 폴더가 없으면 생성
    os.makedirs(extract_path, exist_ok=True)

    # ZIP 파일 열기
    with zipfile.ZipFile(zip_path, 'r') as zip_ref:
        # .hwp 파일만 추출
        hwp_files = [file for file in zip_ref.namelist() if file.endswith('.hwp')]
        for hwp_file in hwp_files:
            zip_ref.extract(hwp_file, extract_path)

    print(f"추출 완료: {extract_path}")


import os


def find_zip_files(directory):
    zip_files = []

    # 폴더를 순회하며 ZIP 파일 찾기
    for root, _, files in os.walk(directory):
        for file in files:
            if file.lower().endswith('.zip'):  # ZIP 파일인지 확인
                zip_files.append(os.path.join(root, file))  # 전체 경로 저장

    return zip_files

zip_files = find_zip_files("C:\\Data\\workplace\\NGD")

@benchmark
def ex_work_test():
    for zip_file in zip_files:
        extract_hwp_from_zip(zip_file)

# 사용 예시 (ZIP 파일 경로 입력)
# zip_file_path = "I:\\test\\NGD\\[NGD]72차 기출 최종완성본 (총 336개 파일)_2020.5.29.zip"  # ZIP 파일 경로
# extract_hwp_from_zip(zip_file_path)

# ex_work_test()

def move_file(src_path, dest_dir):
    """
    파일을 지정된 디렉터리로 이동하는 함수

    :param src_path: 이동할 파일의 경로 (예: "C:/Users/User/Documents/file.txt")
    :param dest_dir: 이동할 대상 디렉터리 (예: "C:/Users/User/Desktop/")
    :return: 이동된 파일의 새 경로 또는 오류 메시지
    """
    try:
        print("src_path:" + src_path)
        print("dest_dir:" + dest_dir)
        # 대상 디렉터리가 없으면 생성
        os.makedirs(dest_dir, exist_ok=True)

        # 파일 이동
        dest_path = os.path.join(dest_dir, os.path.basename(src_path))
        shutil.move(src_path, dest_path)

        return f"파일이 이동되었습니다: {dest_path}"
    except Exception as e:
        return f"파일 이동 중 오류 발생: {e}"



file_name_list = ['1-1-a', '1-1-b', '1-2-a', '1-2-b',
                  '2-1-a', '2-1-b', '2-2-a', '2-2-b',
                  '3-1-a', '3-1-b', '3-2-a', '3-2-b'
                  ]

def exam_classifier_by_grade(input_path, output_root):
    for root, _, files in os.walk(input_path):
        for file in files:
            for i in file_name_list:
                if i in file:
                    if "[고]" in file:
                        move_file( os.path.join(root, file), os.path.join(output_root, os.path.join( "고", i))  )
                    elif  "[중]" in file:
                        move_file(os.path.join(root, file),  os.path.join(output_root, os.path.join("중", i))  )
                    else:
                        print("bug!!!: " + file)





def exam_classifier_by_lv_chap(input_path, output_root):
    # 상-이항분포-2022-3-1-b-청주외고-확통-18-Q
    for root, _, files in os.walk(input_path):
        for file in files:
            name_without_ext, ext = os.path.splitext(file)

            # 이름을 - 로 분리
            parts = name_without_ext.split('-')

            if len(parts) < 3:
                print("파일명이 너무 짧아 처리할 수 없습니다.")
                return

            first = parts[0]  # 첫 번째
            second = parts[1]  # 두 번째
            penultimate = parts[-2]  # 마지막에서 두 번째 = 마지막전

            # 대상 폴더 경로
            target_dir = os.path.join(output_root, first, penultimate, second)

            # 폴더가 없으면 생성
            os.makedirs(target_dir, exist_ok=True)

            # 새 파일 경로
            new_path = os.path.join(target_dir, file)

            # 파일 이동
            shutil.move(input_path, new_path)
            print(f"{input_path} → {new_path} 로 이동 완료.")


# input_path = "C:\\Data\\workplace\\NGD\\output250310_215903"
# output_root = "C:\\Data\DB\\NGD\\"
# input_path = "C:\\Data\\workplace\\output250407_223838"

input_path = "J:\\NGD"
output_root = "J:\\NGD\\done"


# combine_QnA(input_path, output_root)


# zip_files = find_zip_files(input_path)
#
#
# for zip_file in zip_files:
#     extract_hwp_from_zip(zip_file)




# exam_classifier_by_grade(input_path, output_root)

# zipfiles = [ file for file in os.listdir() if file.endswith('zip')]
# print(zipfiles)

# ret = find_zip_files("C:\\Data\\workplace\\NGD")
#
# for zip_path in ret:
#     # ZIP 파일 이름(확장자 제외)으로 폴더명 생성
#     # print(os.path.basename(zip_path))
#     folder_name = os.path.splitext(os.path.basename(zip_path))[0]
#     # print(folder_name)
#     extract_path = os.path.join(os.path.dirname(zip_path), folder_name)
#
#     # 폴더가 없으면 생성
#     os.makedirs(extract_path, exist_ok=True)
#
#     # ZIP 파일 열기
#     with zipfile.ZipFile(zip_path, 'r') as zip_ref:
#         # .hwp 파일만 추출
#         hwp_files = [file for file in zip_ref.namelist() if file.startswith('[고][2021][2-1-b]')]
#         if len(hwp_files) > 0 :
#             print("%s %s", zip_path , hwp_files)
#         # for hwp_file in hwp_files:
#         #     zip_ref.extract(hwp_file, extract_path)
#
#     # print(f"추출 완료: {extract_path}")
