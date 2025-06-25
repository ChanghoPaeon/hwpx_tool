import pyhwpx
import os
import hwp_tool
import datetime
import re
import logger
from pyhwpx import Hwp
hwp = pyhwpx.Hwp()

# 보안팝업 자동클릭
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

# hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 프로그램 실행
# hwp.XHwpWindows.Item(0).Visible = visible  # 기본값 = 백그라운드 해제
# hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈

class ngd_prob_worker:
    def __init__(self, folder_path="./", db_path = "./db"):
        """
        """
        self.folder_path = folder_path
        self.db_path = os.path.join(folder_path, db_path)
        self.working_path = folder_path

    def traverse_work(self):
        """지정된 폴더 내의 모든 폴더 및 파일을 순회"""
        print(self.folder_path)
        idx = 1
        for root, dirs, files in os.walk(self.folder_path):
            # root: 현재 디렉토리 경로
            # dirs: 현재 디렉토리 내의 하위 폴더 리스트
            # files: 현재 디렉토리 내의 파일 리스트

            print(f"현재 폴더: {root}")

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
                    print(f"  {file_name}")
                    if file_name.endswith('.hwp') or file_name.endswith('.hwpx'):
                        self.ngd_prob_to_db(root, file_name, self.db_path)
                        if os.path.exists(os.path.join(self.folder_path, file_name)):
                            print("완료!!: " + os.path.join(self.folder_path, file_name))
                            # os.remove(os.path.exists(os.path.join(self.folder_path, file_name)))


            print("=" * 50)

    # def save_footnote(self, filepath):
    #     pset = hwp.HParameterSet.HSaveFootnote
    #     hwp.HAction.GetDefault("SaveFootnote", pset.HSet)
    #     pset.filename = filepath
    #     return hwp.HAction.Execute("SaveFootnote", pset.HSet)


    def ngd_prob_to_db(self, root, file_name, out_path ="E:\\workspace\\NGD170\\"):

        if file_name.startswith("[중]"):
            print("중학교 skip %s", file_name)
            return

        try:

            # file_name = "[고][2021][1-1-b][울산][강남고][수상][복소수-도형의이동][02035][00002][99001][02010][그림1-5-0-0].hwp"
            ret = True

            print("file_name: " + file_name)

            print("out_path: " + out_path)

            full_path = os.path.join(root, file_name)

            print("join: " + full_path )

            hwp.XHwpWindows.Item(0).Visible = True  # 기본값 = 백그라운드 해제
            # hwp.Open(os.path.join(self.folder_path, file_name))

            print(os.getcwd())

            hwp.Open(full_path)
            hwp.MoveDocBegin()

            # hwp.switch_to(0)
            # hwp.MoveDocBegin()

            pos_list_prob = []
            pos_list_ans = []
            pos_list_ans_end = []

            prob_chap_list = []
            prob_lv_list = []
            # 파일명으로부터 정보추출
            # 년도, 학기,
            pattern = r"\[(\d{4})\]"
            result = re.findall(pattern, file_name)
            year = result[0]

            # 학기
            pattern = r"(?<=\[)(\d-\d-[ab])(?=\])"
            result = re.findall(pattern, file_name)
            semester = result[0]

            # 학교 이름
            pattern = r"\[([^\[\]]{1,}고)\]"
            result = re.findall(pattern, file_name)
            school_name = result[0]

            # 과목 이름
            items = re.findall(r"\[([^\[\]]+)\]", file_name)
            subj = items[5]

            save_name = "-".join([year, semester, school_name, subj])

            # 중단원 및 난이도 추출
            while ret:
                # print("start 중단원")
                ret = hwp.find("[중단원]")

                pos_problem = hwp.get_pos()
                # print(*pos_problem)
                lst = list(pos_problem)
                lst[2] = 0
                pos_modification = tuple(lst)
                # pos_list_prob.append(pos_problem)
                pos_list_prob.append(pos_modification)

                # print(ret)
                hwp.MoveRight()
                hwp.Select()
                hwp.MoveLineEnd()

                chapter_name = hwp.get_selected_text()
                prob_chap_list.append(chapter_name.replace(" ", ""))

                hwp.Cancel()
                # print("start 난이도")
                ret = hwp.find("[난이도]")
                # print(ret)
                hwp.MoveRight()
                hwp.Select()
                hwp.MoveLineEnd()

                chapter_name = hwp.get_selected_text()
                prob_lv_list.append(chapter_name)
                hwp.Cancel()

            print("prob_chap_list: " + str(len(prob_chap_list)))
            print("prob_lv_list: " + str(len(prob_lv_list)))

            pos_prob_num = []
            # 문제 추출 및 저장.
            idx = 1
            for ctrl in hwp.ctrl_list:
                # 미주이면
                if ctrl.UserDesc == "미주":
                    # 해당미주로 이동
                    hwp.SetPosBySet(ctrl.GetAnchorPos(0))
                    # 해당 미주로 진입.
                    hwp.Run("MoveNextPosEx")
                    # 미주 전체 선택
                    hwp.HAction.Run("SelectAll")  # 미주노트 전체선택

                    save_name_base = "-".join([prob_lv_list[idx-1], prob_chap_list[idx-1],  save_name + "-" + '{0:02d}'.format(idx)])


                    hwp.save_block_as(os.path.join(out_path, save_name_base + "-A.hwpx"), "HWPX")
                    hwp.Cancel()

                    hwp.Run("MoveNextPosEx")
                    # print("미주 아웃!")
                    # 문제 시작 위치
                    # 선택하고
                    # 중단원 위치 전까지 가서 선택
                    hwp.Select()
                    hwp.set_pos(*pos_list_prob[idx - 1])
                    hwp.save_block_as(os.path.join(out_path, save_name_base + "-Q.hwpx"), "HWPX")
                    hwp.Cancel()
                    hwp.MoveLineDown()
                    hwp.MoveLineDown()
                    idx = idx + 1
            print("완")
            hwp.Close()
        except Exception as e:
            print("예외 발생:", e)
            logger.logger.debug("%s %s", file_name, e)
            hwp.Close()

        return



    def create_hwp_doc(self):
        hwp.add_doc()
        # hwp.Open(file_path + ".hwp")

        # 페이지 여백 설정
        act = hwp.CreateAction("PageSetup")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("ApplyTo", 3)

        item_set = pset.CreateItemSet("PageDef", "PageDef")
        margin = hwp.MiliToHwpUnit(10)
        item_set.SetItem("TopMargin", margin)
        item_set.SetItem("BottomMargin", margin)
        item_set.SetItem("LeftMargin", margin)
        item_set.SetItem("RightMargin", margin)
        item_set.SetItem("HeaderLen", margin)
        item_set.SetItem("FooterLen", margin)
        item_set.SetItem("GutterLen", margin)

        act.Execute(pset)

        return hwp

    def set_multi_col(self):

        #  단 설정
        # 파라미터셋._prop_map_get_.keys()
        # https://www.inflearn.com/community/questions/1077679/%ED%95%9C%EA%B8%80-%ED%8C%8C%EC%9D%B4%EC%8D%AC-%EB%B0%94%ED%83%95%EC%AA%BD-%EB%8B%A4%EB%8B%A8?srsltid=AfmBOorgKdk__dZPqReicG_VmRNLMrrEm6LP6nHglMw_9Ud5QRInHfSL
        pset = hwp.HParameterSet.HColDef
        hwp.HAction.GetDefault("MultiColumn", pset.HSet)
        pset.Count = 2
        pset.SameSize = 1
        pset.LineType = 1
        pset.SameGap = hwp.MiliToHwpUnit(8.0)  # <--
        pset.HSet.SetItem("ApplyClass", 832)
        pset.HSet.SetItem("ApplyTo", 6)
        hwp.HAction.Execute("MultiColumn", pset.HSet)
        # act.Execute(pset)

        # hwp.XHwpWindows.Active_XHwpWindow.Visible = False


    def extract_from_ebs(self, file_path, out_path="E:\\workspace\\ebs\\"):

        # hwp.XHwpWindows.Active_XHwpWindow.Visible = False

        ret = True

        print("file_path: " + file_path)

        print("out_path: " + out_path)


        hwp.Open(file_path)

        hwp.MoveDocBegin()

        self.create_hwp_doc()

        self.set_multi_col()



        hwp.switch_to(0)
        hwp.MoveDocBegin()

        pos_list_prob = []
        pos_list_ans = []
        pos_list_ans_end = []

        pos_prob_num = []

        while ret:

            ret = hwp.find("#문항코드")
            print(ret)
            hwp.MoveRight()
            hwp.Select()
            hwp.MoveWordEnd()
            prob_num = hwp.get_selected_text()
            print(prob_num)
            pos_prob_num.append(prob_num)
            ret = hwp.MoveLineUp()
            print(ret)
            hwp.Cancel()

            # 문제 찾아서 위치 저장
            ret = hwp.find("[문제]")
            print("문제")
            print(ret)
            # hwp.MoveLineUp()
            pos_problem = hwp.get_pos()
            print(*pos_problem)

            pos_list_prob.append(pos_problem)

            # 정답 및 해설 찾아서 위치 저장
            ret = hwp.find("[정답/모범답안]")
            # print(ret)
            pos_ans = hwp.get_pos()
            print("[정답/모범답안]")
            print(*pos_ans)
            pos_list_ans.append(pos_ans)

            ret = hwp.find("#강")
            # print(ret)

            if not ret:
                print("문서 끝 처리 필요")
                hwp.MoveDocEnd()
            hwp.MoveLineUp()
            pos_ans_end = hwp.get_pos()
            print("pos_ans_end" )
            print(*pos_ans_end)

            pos_list_ans_end.append(pos_ans_end)

            # 문제로 이동
            # select
            # move to ans
            # copy

            # 정답 선택해서 후 복사붙여넣기
            hwp.set_pos(*pos_ans)
            hwp.Select()

            hwp.set_pos(*pos_ans_end)
            hwp.Copy()

            # 새 문서로 전환
            hwp.switch_to(1)

            hwp.Run("BreakColumn")
            # 미주 넣기
            hwp.Run("InsertEndnote")
            # hwp.MoveNextPosEx()
            hwp.Paste()

            # 미주 빠져나오기
            hwp.MoveNextPosEx()
            hwp.MoveNextPosEx()


            # 원본 문서 이동
            hwp.switch_to(0)

            # 문제  선택해서 후 복사붙여넣기
            hwp.set_pos(*pos_problem)
            hwp.Select()
            lst = list(pos_ans)
            lst[2] = 0
            pos_modification = tuple(lst)
            hwp.set_pos(*pos_modification)
            hwp.Copy()
            hwp.switch_to(1)
            # hwp.insert_text(" ["+prob_num+"]")
            hwp.insert_text(" "+prob_num)
            hwp.Paste()
            hwp.switch_to(0)




        hwp.switch_to(1)


        # ctrl + enter
        hwp.Run("BreakPage")

        # 맨 앞으로 가서 del 두번
        hwp.MoveDocBegin()
        hwp.Delete()
        hwp.Delete()

        # i = 0
        # while hwp.get_into_nth_table(i):
        #     hwp.set_table_width()
        #     i += 1


        hwp.save_as(out_path+os.path.basename(file_path)+"-new-.hwpx", "HWPX")

        hwp.XHwpWindows.Active_XHwpWindow.Visible = True
        hwp.Close()
        hwp.switch_to(0)
        hwp.XHwpWindows.Active_XHwpWindow.Visible = True

        hwp.Close()

        print(pos_prob_num)
        print(pos_list_prob)
        print(pos_list_ans)
        print(pos_list_ans_end)

        print("end")

        return out_path+os.path.basename(file_path)+"-new-.hwpx"


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
ebs_worker = ngd_prob_worker("C:\\Data\\workplace\\NGD170\\[NGD]219차 기출 최종완성본 (총 414파일)_2024.09.02 한글본\\", "C:\\Data\\DB\\NGD\\")
ebs_worker.traverse_work()
