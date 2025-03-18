
import os
import sys
import re
import csv

import logger
from tools import benchmark

from hwp_tool import delete_NGD_by_api
import re
import pyperclip

def createDirectory(directory):
    try:
        if not os.path.exists(directory):
            os.makedirs(directory)
    except OSError:
        print("Error: Failed to create the directory.")



# hwp.XHwpWindows.Active_XHwpWindow.Visible = False
# 07030
pattern = r"To\s*\n\d{5}"


# hwp.XHwpWindows.Active_XHwpWindow.Visible = False
# 07030
pattern = r"To\s*\n\d{5}"

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
        self.output_path = output_path + work_start_time
        self.working_path = folder_path

        self.hwp = hwp
        self.work_cnt = 0

        createDirectory(self.output_path)

        self.hwp.XHwpWindows.Active_XHwpWindow.Visible = True

    @benchmark
    def delete_keywords_in_file(self, file_path):
        """
        HWP 파일에서 키워드 등장 횟수를 계산
        :param file_path: 대상 파일의 전체 경로
        :return: {키워드: 등장 횟수} 딕셔너리
        """
        try:
            self.work_cnt = self.work_cnt + 1
            logger.logger.debug(str(self.work_cnt) + " 번째 파일: " + file_path)
            src = re.compile(r'^\[중단원\] .+$')
            dst = re.compile(r".")
            # self.hwp.find_replace_all(src, dst, regex=True)

            whole_text = self.hwp.get_text_file()

            src = re.compile(r'^\[중단원\] .+$', re.MULTILINE)
            dst = r'\n'

            src_list = [i.group() for i in re.finditer(src, whole_text)]
            src = r'^\[중단원\] .+$'
            dst_list = [re.sub(src, dst, i) for i in src_list]
            for i, j in zip(src_list, dst_list):
                self.hwp.find_replace_all(i, j, False, 1, 0, 1, 1,
                         0, 1, 0, 0, 1,
                         1, 0, "", "", 1)

            src = re.compile(r'^\[난이도\] .+$', re.MULTILINE)
            dst = r'\n'

            src_list = [i.group() for i in re.finditer(src, whole_text)]
            src = r'^\[난이도\] .+$'
            dst_list = [re.sub(src, dst, i) for i in src_list]
            for i, j in zip(src_list, dst_list):
                self.hwp.find_replace_all(i, j, False, 1, 0, 1, 1,
                                          0, 1, 0, 0, 1,
                                          1, 0, "", "", 1)



        except Exception as e:
            print(f"\n❌ 파일 읽기 오류: {file_path}, 오류: {e}")

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


            return
        except Exception as e:
            print(f"\n❌ 파일 읽기 오류: {file_path}, 오류: {e}")
            # return {keyword: 0 for keyword in self.keywords}
            return

    def math_eq_refiner(self):

        # content = hwp.GetTextFile()
        # 자릿수 맞춰 줘야 함?
        # 메시지 박스 타입에 맞게 설정해줘야 함
        self.hwp.SetMessageBoxMode(0x00000010)

        self.hwp.MoveDocBegin()
        # 메시지 박스 자동
        for ctrl in self.hwp.ctrl_list:
            # print(ctrl)
            # 미주 미주

            # print(ctrl.UserDesc)
            if ctrl.UserDesc == "수식":

                ret = self.hwp.select_ctrl(ctrl)
                self.hwp.move_to_ctrl(ctrl)
                # hwp.MoveRight()
                # hwp.BreakPara()
                # hwp.set_font(TextColor="Red")
                # hwp.insert_text(ctrl.Properties.Item("String"))
                matheq = ctrl.Properties.Item("String")
                matches = re.findall(pattern, matheq)
                # print(matheq)

                if matches:
                    print("=================== 바꾸기 전")
                    print(matheq)
                    result = re.sub(pattern, "", matheq)
                    print("=================== 바꾸고 난 후")
                    result = re.sub(r"\s+$", " ", result)
                    # # stripped = result.rstrip() + " " if result.strip() != result.rstrip() else result
                    print(result)
                    print("=================== ")

                    ctrl.Properties.SetItem("String", result)
                    ctrl.Properties.SetItem("VisualString", result)
                    # hwp.HAction.GetDefault("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
                    # # hwp.HParameterSet.SetItem("String", result)
                    # print(hwp.HParameterSet.HEqEdit.string)
                    # print(hwp.HParameterSet.HEqEdit)
                    # hwp.HParameterSet.HEqEdit.VisualString = result
                    # print(hwp.HParameterSet.HEqEdit.VisualString)
                    # print(hwp.HParameterSet.HEqEdit.Version)
                    # hwp.HAction.Execute("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
                    # hwp.HAction.Run("SelectAll")
                    # 수식 하나 바꾼 후 죽음 이게 원인인가
                    # hwp.HAction.Run("Delete")
                    # 바뀌긴 하나 파일 수식 열어서 편집창 등왔다리 갔다리 해야 변경 적용됨 이유는 뭘까?
                    # pset = hwp.HParameterSet.HEqEdit
                    # hwp.HAction.GetDefault("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
                    # hwp.HParameterSet.HEqEdit.string = result
                    # hwp.HParameterSet.HEqEdit.VisualString = result
                    # # hwp.HParameterSet.HEqEdit.Version = "Equation Version 60"
                    # hwp.HAction.Execute("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
                    # hwp.Run("Cancel")  # 폰트 예뻐짐
                    # hwp.Run("MoveRight")  # 다음 수식 삽입 준비
                    # hwp.MoveRight()

                    # https://employeecoding.tistory.com/335
                    # act = hwp.CreateAction("EquationPropertyDialog")
                    # pset = act.CreateSet()
                    # act.GetDefault(pset)
                    #
                    # pset.SetItem("String", result)
                    # act.Execute(pset)

                    self.hwp.HAction.GetDefault("EquationCreate", self.hwp.HParameterSet.HEqEdit.HSet)
                    self.hwp.HParameterSet.HEqEdit.EqFontName = "HancomEQN"
                    self.hwp.HParameterSet.HEqEdit.EqFontName = "HYhwpEQ"
                    self.hwp.HParameterSet.HEqEdit.string = result
                    self.hwp.HParameterSet.HEqEdit.TreatAsChar = True
                    self.hwp.HParameterSet.HEqEdit.BaseUnit = self.hwp.PointToHwpUnit(11.0)
                    self.hwp.HParameterSet.HEqEdit.Version = "Equation Version 60"
                    self.hwp.HAction.Execute("EquationCreate", self.hwp.HParameterSet.HEqEdit.HSet)

                    self.hwp.HAction.Run("Delete")

                    # hwp.HAction.GetDefault("EquationPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
                    # hwp.HParameterSet.HShapeObject.string = result
                    # hwp.HAction.Execute("EquationPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)

                    #
                    self.hwp.MoveRight()
                    # hwp.BreakPara()
                    # hwp.set_font(TextColor="Red")
                    # hwp.insert_text(ctrl.Properties.Item("String"))

                # if matheq.find("N.G.D") !=-1 or matheq.find("NGD") != -1:
                #     print(ctrl.Properties.Item("String"))

    def save_as_pdf(self, file_path):
        print("pdf: " + file_path)
        print("target path:" + os.path.join(self.output_path, os.path.split(file_path)[-1] + "_converted.pdf"))

        pset = self.hwp.HParameterSet.HFileOpenSave
        self.hwp.HAction.GetDefault("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)
        self.hwp.HParameterSet.HFileOpenSave.filename = os.path.join(self.output_path,
                                                                     os.path.split(file_path)[-1] + "_converted.pdf")
        self.hwp.HParameterSet.HFileOpenSave.Format = "PDF"
        self.hwp.HParameterSet.HFileOpenSave.Attributes = 16384
        result = self.hwp.HAction.Execute("FileSaveAsPdf", self.hwp.HParameterSet.HFileOpenSave.HSet)

        print(result)

    def save_as_hwpx(self, file_path):
        print("pdf: " + file_path)
        print("target path:" + os.path.join(self.output_path, os.path.split(file_path)[-1] + "_converted.hwpx"))

        self.hwp.SaveAs(os.path.join(self.output_path, os.path.split(file_path)[-1] + "_converted.hwpx"),
                        "HWPX")  # hwpx로 저장


    def math_eq_refiner(self):

        self.hwp.MoveDocBegin()
        # 메시지 박스 자동
        for ctrl in self.hwp.ctrl_list:
            # print(ctrl)
            # 미주 미주

            # print(ctrl.UserDesc)
            if ctrl.UserDesc == "수식":

                ret = self.hwp.select_ctrl(ctrl)
                self.hwp.move_to_ctrl(ctrl)
                # hwp.MoveRight()
                # hwp.BreakPara()
                # hwp.set_font(TextColor="Red")
                # hwp.insert_text(ctrl.Properties.Item("String"))
                matheq = ctrl.Properties.Item("String")
                matches = re.findall(pattern, matheq)
                # print(matheq)

                if matches:
                    print("=================== 바꾸기 전")
                    print(matheq)
                    result = self.re.sub(pattern, "", matheq)
                    print("=================== 바꾸고 난 후")
                    result = self.re.sub(r"\s+$", " ", result)
                    # # stripped = result.rstrip() + " " if result.strip() != result.rstrip() else result
                    print(result)
                    print("=================== ")

                    ctrl.Properties.SetItem("String", result)
                    ctrl.Properties.SetItem("VisualString", result)
                    # hwp.HAction.GetDefault("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
                    # # hwp.HParameterSet.SetItem("String", result)
                    # print(hwp.HParameterSet.HEqEdit.string)
                    # print(hwp.HParameterSet.HEqEdit)
                    # hwp.HParameterSet.HEqEdit.VisualString = result
                    # print(hwp.HParameterSet.HEqEdit.VisualString)
                    # print(hwp.HParameterSet.HEqEdit.Version)
                    # hwp.HAction.Execute("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
                    # hwp.HAction.Run("SelectAll")
                    # 수식 하나 바꾼 후 죽음 이게 원인인가
                    # hwp.HAction.Run("Delete")
                    # 바뀌긴 하나 파일 수식 열어서 편집창 등왔다리 갔다리 해야 변경 적용됨 이유는 뭘까?
                    # pset = hwp.HParameterSet.HEqEdit
                    # hwp.HAction.GetDefault("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
                    # hwp.HParameterSet.HEqEdit.string = result
                    # hwp.HParameterSet.HEqEdit.VisualString = result
                    # # hwp.HParameterSet.HEqEdit.Version = "Equation Version 60"
                    # hwp.HAction.Execute("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
                    # hwp.Run("Cancel")  # 폰트 예뻐짐
                    # hwp.Run("MoveRight")  # 다음 수식 삽입 준비
                    # hwp.MoveRight()

                    # https://employeecoding.tistory.com/335
                    # act = hwp.CreateAction("EquationPropertyDialog")
                    # pset = act.CreateSet()
                    # act.GetDefault(pset)
                    #
                    # pset.SetItem("String", result)
                    # act.Execute(pset)

                    self.hwp.HAction.GetDefault("EquationCreate", self.hwp.HParameterSet.HEqEdit.HSet)
                    self.hwp.HParameterSet.HEqEdit.EqFontName = "HancomEQN"
                    self.hwp.HParameterSet.HEqEdit.EqFontName = "HYhwpEQ"
                    self.hwp.HParameterSet.HEqEdit.string = result
                    self.hwp.HParameterSet.HEqEdit.TreatAsChar = True
                    self.hwp.HParameterSet.HEqEdit.BaseUnit = self.hwp.PointToHwpUnit(11.0)
                    self.hwp.HParameterSet.HEqEdit.Version = "Equation Version 60"
                    self.hwp.HAction.Execute("EquationCreate", self.hwp.HParameterSet.HEqEdit.HSet)

                    self.hwp.HAction.Run("Delete")

                    # hwp.HAction.GetDefault("EquationPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
                    # hwp.HParameterSet.HShapeObject.string = result
                    # hwp.HAction.Execute("EquationPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)

                    #
                    self.hwp.MoveRight()
                    # hwp.BreakPara()
                    # hwp.set_font(TextColor="Red")
                    # hwp.insert_text(ctrl.Properties.Item("String"))

                # if matheq.find("N.G.D") !=-1 or matheq.find("NGD") != -1:
                #     print(ctrl.Properties.Item("String"))

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

    @benchmark
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
            self.delete_keywords_in_file(file_path)
            # self.math_eq_refiner()
            # self.save_as_hwpx(file_path)
            self.save_as_pdf(file_path)

            sys.stdout.write(f"\r🔄 진행 중: {i}/{total_files} ({(i / total_files) * 100:.2f}%)")
            sys.stdout.flush()




    def traverse_work(self):
        """지정된 폴더 내의 모든 폴더 및 파일을 순회"""
        print(self.folder_path)
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
