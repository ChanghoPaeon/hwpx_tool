import os
import sys
import csv

import tools

import datetime

from tools import benchmark


# ✅ 실행 예제
# folder_path = "I:\\test\\[NGD]179차 기출 최종완성본 (총 376파일)_2023.07.26 한글본\\"
# keywords = ["[난이도] 특", "[난이도] 상", "[난이도] 중", "[난이도] 하", "검색어5"]
# output_csv = "I:\\test\\" + "[NGD]179차" + ".csv"
#
# # 객체 생성 및 실행
# searcher = HwpKeywordSearcher(folder_path, keywords, output_csv)
# searcher.search()



class HwpKeywordSearcher:
    def __init__(self, hwp, folder_path, keywords, output_csv_path):
        """
        HWP 파일에서 특정 키워드를 검색하는 클래스
        :param folder_path: 검색할 폴더 경로
        :param keywords: 찾을 키워드 목록 (리스트)
        :param output_csv: 결과를 저장할 CSV 파일 경로
        """
        self.folder_path = folder_path
        self.keywords = keywords
        self.output_csv_path = output_csv_path
        self.working_path = folder_path

        self.hwp = hwp

    @benchmark
    def count_keywords_in_file(self, file_path):
        """
        HWP 파일에서 키워드 등장 횟수를 계산
        :param file_path: 대상 파일의 전체 경로
        :return: {키워드: 등장 횟수} 딕셔너리
        """
        try:
            doc = self.hwp.open(file_path)
            content = self.hwp.GetTextFile()
            return {keyword: content.count(keyword) for keyword in self.keywords}
        except Exception as e:
            print(f"\n❌ 파일 읽기 오류: {file_path}, 오류: {e}")
            return {keyword: 0 for keyword in self.keywords}

    def parse_filename(self, file_name):
        """
        파일명을 파싱하여 메타데이터를 추출
        :param file_name: HWP 파일 이름
        :return: ['중고', '연도', '학기', '지역', '학교명', '과목', '범위'] 리스트
        """
        split_name = file_name.split("][")

        if len(split_name) < 7:
            print(f"\n⚠️ 파일명 형식 오류: {file_name}")
            return None  # 잘못된 파일명은 None 반환

        if split_name[0].endswith("고"):
            return ["고"] + split_name[1:7]
        elif split_name[0].endswith("중"):
            return ["중"] + split_name[1:5] + ["-"] + [split_name[6]]
        else:
            print(f"\n⚠️ 학년 구분 실패: {file_name}")
            return None

    def save_to_csv(self, results):
        """
        결과를 CSV 파일로 저장
        :param results: 검색된 결과 리스트
        """

        suffix = datetime.datetime.now().strftime('%y%m%d_%H%M%S')
        output_csv_name = suffix + "total.csv"
        result_csv_full = self.output_csv_path + "\\" + output_csv_name
        with open(result_csv_full, 'a+', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(["중고", "연도", "학기", "지역", "학교명", "과목", "범위"] + self.keywords)
            writer.writerows(results)
        print(f"\n✅ 검색 완료! 결과가 {result_csv_full}에 저장되었습니다.")

    def search(self):

        """폴더 내 모든 HWP 파일을 검색하고 결과를 CSV로 저장"""

        print("\nstart at: " + self.working_path)

        files = [f for f in os.listdir(self.folder_path+"\\"+self.working_path) if f.endswith('.hwp')]
        total_files = len(files)

        if total_files == 0:
            print("📂 해당 폴더에 HWP 파일이 없습니다.")
            return

        print(f"🔍 총 {total_files}개의 파일을 검색합니다.")
        results = []

        for i, file in enumerate(files, start=1):
            file_path = os.path.join(self.folder_path, self.working_path, file)
            file_meta = self.parse_filename(file)

            if file_meta is None:
                continue  # 파일명 오류 시 스킵

            word_counts = self.count_keywords_in_file(file_path)
            results.append(file_meta + [word_counts[key] for key in self.keywords])

            sys.stdout.write(f"\r🔄 진행 중: {i}/{total_files} ({(i / total_files) * 100:.2f}%)")
            sys.stdout.flush()

        self.save_to_csv(results)

    def traverse_work(self):
        """지정된 폴더 내의 모든 폴더 및 파일을 순회"""
        for root, dirs, files in os.walk(self.folder_path):
            # root: 현재 디렉토리 경로
            # dirs: 현재 디렉토리 내의 하위 폴더 리스트
            # files: 현재 디렉토리 내의 파일 리스트

            print(f"현재 폴더: {root}")

            if dirs:
                print("하위 폴더들:")
                for dir_name in dirs:
                    print(f"  {dir_name}")
                    self.working_path = dir_name
                    self.search()

            if files:
                print("파일들:")
                for file_name in files:
                    print(f"  {file_name}")

            print("=" * 50)
