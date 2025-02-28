
import os
import sys
import csv

from hpw_tool import delete_NGD_by_api


class NGD_converter:
    def __init__(self, hwp, folder_path, keywords, output_path, work_start_time):
        """
        HWP 파일에서 특정 키워드를 검색하는 클래스
        :param folder_path: 검색할 폴더 경로
        :param keywords: 찾을 키워드 목록 (리스트)
        :param output_csv: 결과를 저장할 CSV 파일 경로
        """
        self.output_csv_name = work_start_time + "_total.csv"
        self.folder_path = folder_path
        self.output_path = output_path
        self.working_path = folder_path

        self.hwp = hwp

    def delete_ngd_ctrls(self, file_path):
        """
        HWP 파일에서 키워드 등장 횟수를 계산
        :param file_path: 대상 파일의 전체 경로
        :return: {키워드: 등장 횟수} 딕셔너리
        """
        try:
            doc = self.hwp.open(file_path)
            # content = self.hwp.HAction.GetDefault("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)
            # self.hwp.HParameterSet.HDeleteCtrls.CreateItemArray("DeleteCtrlType", 1)
            # self.hwp.HParameterSet.HDeleteCtrls.DeleteCtrlType.SetItem(0, 7)  # <--- Item을 SetItem으로 고쳤음.
            # self.hwp.HAction.Execute("DeleteCtrls", self.hwp.HParameterSet.HDeleteCtrls.HSet)


            #바탕쪽 삭제
            # 바탕쪽 삭제 후 다시 시도하는 경우 예외처리 필요
            if hasattr(self.hwp.HParameterSet.HMasterPage, "Hset"):
                self.hwp.HAction.GetDefault("MasterPage", self.hwp.HParameterSet.HMasterPage.Hset)
                self.hwp.HAction.Execute("MasterPage", self.hwp.HParameterSet.HMasterPage.Hset)

                self.hwp.HAction.Run("DeleteDocumentMasterPage")

            # 표 선택 및 삭제
            target = 1
            self.hwp.get_into_nth_table(target, select=True)
            # 해당 표안에 콘텐츠 관련 내용이 있으면, 삭제하는 코드 구현

            # A1셀 안에서 아래 코드를 실행하면 표 전체를 선택함
            self.hwp.SelectCtrlFront()



            self.hwp.Delete()
            print("pdf: "+file_path)
            print("target path:" + os.path.join(self.output_path, os.path.split(file_path)[-1]+"_converted.pdf"))
            # result = self.hwp.save_as(file_path+"_converted.pdf.pdf", format="pdf")

            pset = self.hwp.HParameterSet.HFileOpenSave
            self.hwp.HAction.GetDefault("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)
            self.hwp.HParameterSet.HFileOpenSave.filename = os.path.join(self.output_path, os.path.split(file_path)[-1]+"_converted.pdf")
            self.hwp.HParameterSet.HFileOpenSave.Format = "PDF"
            self.hwp.HParameterSet.HFileOpenSave.Attributes = 16384
            result = self.hwp.HAction.Execute("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)

            print(result)

            return
        except Exception as e:
            print(f"\n❌ 파일 읽기 오류: {file_path}, 오류: {e}")
            # return {keyword: 0 for keyword in self.keywords}
            return

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
        output_csv_name = "total.csv"
        result_csv_full = self.output_path + "\\" + self.output_csv_name
        with open(result_csv_full, 'a+', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f)
            writer.writerow(["중고", "연도", "학기", "지역", "학교명", "과목", "범위"] + self.keywords)
            writer.writerows(results)
        print(f"\n✅ 검색 완료! 결과가 {result_csv_full}에 저장되었습니다.")

    def converter(self):

        """폴더 내 모든 HWP 파일을 검색하고 결과를 CSV로 저장"""

        print("\nstart at: " + self.working_path)

        target_path = ""
        if self.working_path == self.folder_path:
            target_path = self.working_path
        else:
            target_path = self.folder_path+"\\"+self.working_path
        files = [f for f in os.listdir(target_path) if f.endswith('.hwp')]
        total_files = len(files)

        if total_files == 0:
            print("📂 해당 폴더에 HWP 파일이 없습니다.")
            return

        print(f"🔍 총 {total_files}개의 파일의 변환을 시작합니다.")
        results = []

        for i, file in enumerate(files, start=1):
            file_path = os.path.join(target_path, file)
            file_meta = self.parse_filename(file)

            if file_meta is None:
                continue  # 파일명 오류 시 스킵

            # word_counts = self.count_keywords_in_file(file_path)
            # results.append(file_meta + [word_counts[key] for key in self.keywords])

            self.delete_ngd_ctrls(file_path)

            sys.stdout.write(f"\r🔄 진행 중: {i}/{total_files} ({(i / total_files) * 100:.2f}%)")
            sys.stdout.flush()

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
                    self.converter()

            if files:
                print("파일들:")
                for file_name in files:
                    print(f"  {file_name}")
                    self.converter()

            print("=" * 50)
