import os
import sys
import csv
from collections import Counter
from glob import glob
import pyhwpx
hwp=pyhwpx.Hwp()

def count_keywords_in_file(dir, file_path, keywords):
    """HWP 파일에서 특정 키워드들의 등장 횟수를 계산"""
    try:
        # print(os.path.join(dir, file_path))
        doc = hwp.open(os.path.join(dir, file_path))
        # content = doc.text
        content = hwp.GetTextFile()

        search_result = {keyword: content.count(keyword) for keyword in keywords}

        # hwp.Clear(option=1)
        # print(search_result)

        return search_result
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return {keyword: 0 for keyword in keywords}


def search_keywords_in_folder(folder_path, keywords, output_csv):
    """폴더 내 HWP 파일에서 특정 단어를 검색하고 결과를 CSV로 저장"""
    # print("files: ", files)
    # print("folder_path: ", folder_path)
    # print("path: ", os.path.join(folder_path, "*.hwp"))

    files = os.listdir(folder_path)
    # print("files: ", files)

    results = []

    total_files = len(files)

    print("총 파일 수: ", total_files)

    # for i, file in enumerate(file_list, start=1):


    for i, file in enumerate(files, start=1):
        file_name = os.path.basename(file)
        word_counts = count_keywords_in_file(folder_path, file, keywords)
        # 중/고 + 연도/ 학기/학교명 + 범위
        # 0, 1, 2, 3, 4, 5, 6,
        split_file_name = file_name.split("][")

        pre_result = []


        if split_file_name[0].endswith("고"):
            pre_result = ["고",
                            split_file_name[1],
                            split_file_name[2],
                            split_file_name[3],
                            split_file_name[4],
                            split_file_name[5],
                            split_file_name[6]
                            ] + [word_counts[key] for key in keywords]
        elif split_file_name[0].endswith("중"):
            pre_result = ["중",
                           split_file_name[1],
                           split_file_name[2],
                           split_file_name[3],
                           split_file_name[4],
                           "-",
                           split_file_name[6]
                           ] + [word_counts[key] for key in keywords]


        results.append(pre_result)

        sys.stdout.write(f"\rProcessing {i}/{total_files} ({(i / total_files) * 100:.2f}%)")
        sys.stdout.flush()


    # CSV 저장
    with open(output_csv, 'w+', newline='', encoding='utf-8-sig') as f:
        writer = csv.writer(f)
        writer.writerow(["중고", "연도", "학기", "지역", "학교명", "과목", "범위"] + keywords)
        writer.writerows(results)

    print(f"검색 완료! 결과가 {output_csv}에 저장되었습니다.")


# 사용 예제
# folder_path = "I:\\test\\[NGD]179차 기출 최종완성본 (총 376파일)_2023.07.26 한글본\\"  # 검색할 폴더 경로
# keywords = ["[난이도] 특", "[난이도] 상", "[난이도] 중", "[난이도] 하", "검색어5"]  # 찾고 싶은 단어 목록
# output_csv = "I:\\test\\"+ "[NGD]179차"+".csv"  # 결과 CSV 파일명
#
# search_keywords_in_folder(folder_path, keywords, output_csv)