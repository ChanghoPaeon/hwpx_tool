import pyhwpx
import os
import hwp_tool
import datetime
import re
import logger
from pyhwpx import Hwp


class ngd_prob_worker:
    def __init__(self, folder_path="./", db_path = "./db"):
        """
        """
        print("init ngd_prob_worker")
        self.folder_path = folder_path
        self.db_path = os.path.join(folder_path, db_path)
        self.working_path = folder_path

        self.hwp = pyhwpx.Hwp()
        self.hwp.XHwpWindows.Item(0).Visible = False  # 기본값 = 백그라운드 해제
        self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        # 보안팝업 자동클릭
        # hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

    def __del__(self):
        self.hwp.quit()
        print("del ngd_prob_worker")


    def traverse_work(self):
        """지정된 폴더 내의 모든 폴더 및 파일을 순회"""
        print(self.folder_path)
        idx = 1
        for root, dirs, files in os.walk(self.folder_path):
            # root: 현재 디렉토리 경로
            # dirs: 현재 디렉토리 내의 하위 폴더 리스트
            # files: 현재 디렉토리 내의 파일 리스트

            logger.logger.debug(f"현재 폴더: {root}")

            if dirs:
                print("하위 폴더들:")
                for dir_name in dirs:
                    print(f"  {dir_name}")
                    self.working_path = root
                    # self.ebs_prob_to_db(file_name, self.db_path)

            if files:
                print("파일들:")
                for file_name in files:
                    print(idx)
                    idx = idx + 1
                    logger.logger.debug(f"  {file_name}")
                    if file_name.endswith('.hwp') or file_name.endswith('.hwpx'):
                        self.ngd_prob_to_db(root, file_name, self.db_path)
                        if os.path.exists(os.path.join(self.folder_path, file_name)):
                            logger.logger.debug("완료!!: " + os.path.join(self.folder_path, file_name))
                            # os.remove(os.path.exists(os.path.join(self.folder_path, file_name)))


            print("=" * 50)

    # def save_footnote(self, filepath):
    #     pset = hwp.HParameterSet.HSaveFootnote
    #     hwp.HAction.GetDefault("SaveFootnote", pset.HSet)
    #     pset.filename = filepath
    #     return hwp.HAction.Execute("SaveFootnote", pset.HSet)


    def ngd_prob_to_db(self, root, file_name, out_path ="E:\\workspace\\NGD170\\"):

        pattern =     r"^\[서술형[^\]]*\]"
        result = re.findall(pattern, file_name)
        rerined_filename = file_name
        if result:
            rerined_filename = re.sub(pattern, "", file_name)

        if rerined_filename.startswith("[중]"):
            print("중학교 skip %s", file_name)
            return
        try:
            # file_name = "[고][2021][1-1-b][울산][강남고][수상][복소수-도형의이동][02035][00002][99001][02010][그림1-5-0-0].hwp"
            ret = True

            print("file_name: " + file_name)

            print("out_path: " + out_path)

            full_path = os.path.join(root, file_name)

            print("join: " + full_path )
            print(os.getcwd())

            self.hwp.Open(full_path)
            self.hwp.MoveDocBegin()

            # hwp.switch_to(0)
            # hwp.MoveDocBegin()

            pos_list_prob = []

            prob_chap_list = []
            prob_lv_list = []
            # 파일명으로부터 정보추출
            # 년도, 학기,
            pattern = r"\[(\d{4})\]"
            result = re.findall(pattern, rerined_filename)
            year = result[0]

            # 학기
            pattern = r"(?<=\[)(\d-\d-[ab])(?=\])"
            result = re.findall(pattern, rerined_filename)
            semester = result[0]

            # 학교 이름
            pattern = r"\[([^\[\]]{1,}고)\]"
            result = re.findall(pattern, rerined_filename)
            school_name = result[0]

            # 과목 이름
            items = re.findall(r"\[([^\[\]]+)\]", rerined_filename)
            subj = items[5]

            save_name = "-".join([year, semester, school_name, subj])

            # 중단원 및 난이도 추출
            while ret:
                # print("start 중단원")
                ret = self.hwp.find("[중단원]")

                pos_problem = self.hwp.get_pos()
                # print(*pos_problem)
                lst = list(pos_problem)
                lst[2] = 0
                pos_modification = tuple(lst)
                # pos_list_prob.append(pos_problem)
                pos_list_prob.append(pos_modification)

                # print(ret)
                self.hwp.MoveRight()
                self.hwp.Select()
                self.hwp.MoveLineEnd()

                chapter_name = self.hwp.get_selected_text()
                prob_chap_list.append(chapter_name.replace(" ", ""))

                self.hwp.Cancel()
                # print("start 난이도")
                ret = self.hwp.find("[난이도]")
                # print(ret)
                self.hwp.MoveRight()
                self.hwp.Select()
                self.hwp.MoveLineEnd()

                chapter_name = self.hwp.get_selected_text()
                prob_lv_list.append(chapter_name)
                self.hwp.Cancel()

            print("prob_chap_list: " + str(len(prob_chap_list)))
            print("prob_lv_list: " + str(len(prob_lv_list)))

            pos_prob_num = []
            # 문제 추출 및 저장.
            idx = 1
            for ctrl in self.hwp.ctrl_list:
                # 미주이면
                if ctrl.UserDesc == "미주":
                    # 해당미주로 이동
                    self.hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    # 해당 미주로 진입.
                    self.hwp.Run("MoveNextPosEx")
                    # 미주 전체 선택
                    self.hwp.HAction.Run("SelectAll")  # 미주노트 전체선택

                    save_name_base = "-".join([prob_lv_list[idx-1], prob_chap_list[idx-1],  save_name + "-" + '{0:02d}'.format(idx)])


                    self.hwp.save_block_as(os.path.join(out_path, save_name_base + "-A.hwpx"), "HWPX")
                    self.hwp.Cancel()

                    self.hwp.Run("MoveNextPosEx")
                    # print("미주 아웃!")
                    # 문제 시작 위치
                    # 선택하고
                    # 중단원 위치 전까지 가서 선택
                    self.hwp.Select()
                    self.hwp.set_pos(*pos_list_prob[idx - 1])
                    self.hwp.save_block_as(os.path.join(out_path, save_name_base + "-Q.hwpx"), "HWPX")
                    self.hwp.Cancel()
                    self.hwp.MoveLineDown()
                    self.hwp.MoveLineDown()
                    idx = idx + 1
            # logger.logger.debug("완")
            print("완")
            self.hwp.Close()
            self.hwp.Clear()
            # hwp.Quit()
        except Exception as e:
            print("예외 발생:", e)
            logger.logger.debug("%s %s", file_name, e)
            self.hwp.Close()
            self.hwp.Clear()
            # hwp.Quit()

        return




# file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS 2026학년도 수능특강 수학영역  수학Ⅱ_(185).hwp"
# file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS test.hwp"
# file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS 2026학년도 수능특강 수학영역 수학Ⅰ_(182).hwp"
# file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS 2026학년도 수능특강 수학영역 확률과 통계_(186).hwp"
# file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS 2026학년도 수능특강 수학영역 미적분_(172).hwp"

suffix = datetime.datetime.now().strftime('%y%m%d_%H%M%S')

# file_path = "E:\\학원-과외\\학원-부산\\문제집-일반\\EBS\\4주특강\\EBS 2024학년도 수능연계완성 4주 특강 고난도·신유형 수학영역 수학Ⅰ·수학Ⅱ·확률과 통계.hwp"
# ebs_worker = ebs_prob_worker()
# converted_path = ebs_worker.extract_from_ebs(file_path, "E:\\workspace\\ebs\\")
# #
# hwp = hwp_tool.hwp_init(converted_path)

# converted_path = "E:\\workspace\\ebs\\EBS 2024학년도 수능연계완성 4주 특강 고난도·신유형 수학영역 수학Ⅰ·수학Ⅱ·확률과 통계.hwp-new-.hwpx"
# hwp = hwp_tool.hwp_init(converted_path)
# hwp_tool.adjust_image_width(hwp)
# hwp_tool.adjust_table_width(hwp)
# 완료 1, 2,  6 
# 1 수특
# 2 수완
# 3 모고
# 4 ?
# 5 4특
# 6 올림포스 고난도
# C:\Data\workplace\NGD170\[NGD]231차 기출 최종완성본 (총 412파일)_2025.01.09 한글본
target = [
# "[NGD]106차 기출 최종완성본 (총 384파일)_2021.05.29(1)",
# "[NGD]106차 기출 최종완성본 (총 384파일)_2021.05.29(2)",
#     "[NGD]109차 기출 최종완성본 (총 388파일)_2021.06.29(1)", O
#     "[NGD]109차 기출 최종완성본 (총 388파일)_2021.06.29(2)", O
    # "[NGD]110차 기출 최종완성본 (총 376파일)_2021.07.09(1)",
    # "[NGD]110차 기출 최종완성본 (총 376파일)_2021.07.09(2)",
    # "[NGD]111차 기출 최종완성본 (총 376파일)_2021.07.19(1)",
    # "[NGD]111차 기출 최종완성본 (총 376파일)_2021.07.19(2)",
    # "[NGD]113차 기출 최종완성본 (총 366파일)_2021.08.09(1)",
    # "[NGD]113차 기출 최종완성본 (총 366파일)_2021.08.09(2)",
    # '[NGD]148차 기출 최종완성본 (총 420파일)_2022.08.29 한글본', 대전용산고 이후 X
    # '[NGD]149차 기출 최종완성본 (총 408파일)_2022.09.19 한글본',  대전용산고
    # '[NGD]150차 기출 최종완성본 (총 414파일)_2022.09.29 한글본', '대전 제일고
    # "[NGD]151차 기출 최종완성본 (총 424파일)_2022.10.09 한글본",
    # "[NGD]153차 기출 최종완성본 (총 402파일)_2022.10.29 한글본",
    # "[NGD]156차 기출 최종완성본 (총 442파일)_2022.11.29 한글본",
    # "[NGD]157차 기출 최종완성본 (총 420파일)_2022.12.09 한글본",
    # "[NGD]158차 기출 최종완성본  (총 438파일)_2022.12.19 한글본",
    # "[NGD]159차 기출 최종완성본 (총 425파일)_2022.12.29 한글본",
    # "[NGD]160차 기출 최종완성본 (총 420파일)_2023.01.09 한글본",
    "[NGD]161차 기출 최종완성본 (총 404파일)_2023.01.19 한글본",
    # "[NGD]162차 기출 최종완성본 (총 436파일)_2023.02.09 한글본",
    # "[NGD]163차 기출 최종완성본 (총 402파일)_2023.02.19 한글본",
    # "[NGD]164차 기출 최종완성본 (총 394파일)_2023.02.29 한글본",
    # "[NGD]165차 기출 최종완성본 (총 406파일)_2023.03.09 한글본",
    # "[NGD]166차 기출 최종완성본 (총 422파일)_2023.03.19 한글본",
    # "[NGD]167차 기출 최종완성본 (총 422파일)_2023.03.29 한글본",
    # "[NGD]168차 기출 최종완성본 (총 398파일)_2023.04.09 한글본",
    # "[NGD]169차 기출 최종완성본 (총 404파일)_2023.04.19 한글본",

]


for i in target:
    ngd_worker = ngd_prob_worker( "I:\\test\\NGD\\" + i, "C:\\Data\\DB\\NGD\\")
    ngd_worker.traverse_work()
    del ngd_worker

# ebs_worker.traverse_work()
