import pyhwpx
import os
import hwp_tool
import datetime
from pyhwpx import Hwp
from pathlib import Path

import logger

# hwp = pyhwpx.Hwp()

# 보안팝업 자동클릭
# hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

# 한글 문서 만들기

# 페이지 셋업

import pyhwpx
import os
import hwp_tool
import datetime
from pyhwpx import Hwp

import pandas as pd


# 보안팝업 자동클릭
# hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

# hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 프로그램 실행
# hwp.XHwpWindows.Item(0).Visible = visible  # 기본값 = 백그라운드 해제
# hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈

class workbook_publisher:
    def __init__(self, folder_path="./", output_path = "./db", db_path="./"):
        """
        """
        self.folder_path = folder_path
        self.db_path = db_path
        self.working_path = folder_path
        self.hwp = pyhwpx.Hwp()
        self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
        # self.hwp.XHwpWindows.Item(0).Visible = False  # 기본값 = 백그라운드 해제
        self.prob_type_writing_list = [[],[],[],[],[],[],[],[],[],[],[],[],[]]
        self.prob_type_wrong_anslist = [[],[],[],[],[],[],[],[],[],[],[],[],[]]
        self.prob_type_often_list = [[],[],[],[],[],[],[],[],[],[],[],[],[]]
        self.output_path = output_path

        self.hwp_to_pdf_list =[]


        self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

        # 보안팝업 자동클릭
        # hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")


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
                    # self.ebs_prob_to_db(file_name, self.db_path)

            if files:
                print("파일들:")
                for file_name in files:
                    # print(f"  {file_name}")
                    if file_name.startswith("[최다오"):
                        chapter = file_name.split(".")
                    else:
                        chapter = file_name.split("-")
                    print(chapter[0][-1])
                    d = int(chapter[0][-1]) - 1


                    if file_name.startswith("[서술형") :
                        self.prob_type_writing_list[d].append( os.path.join(self.working_path, file_name))
                    elif file_name.startswith("[최다빈출"):
                        self.prob_type_often_list[d].append(os.path.join(self.working_path, file_name))
                    elif file_name.startswith("[최다오답"):
                        self.prob_type_wrong_anslist[d].append(os.path.join(self.working_path, file_name))
                    self.hwp_to_pdf_list.append(os.path.join(self.working_path, file_name))

            print("=" * 50)
        print( self.prob_type_writing_list)
        print(self.prob_type_often_list)
        print(self.prob_type_wrong_anslist)


    def create_hwp_doc(self):
        self.hwp.add_doc()
        # hwp.Open(file_path + ".hwp")

        # 페이지 여백 설정
        act = self.hwp.CreateAction("PageSetup")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("ApplyTo", 3)

        item_set = pset.CreateItemSet("PageDef", "PageDef")
        margin = self.hwp.MiliToHwpUnit(5)
        item_set.SetItem("TopMargin", margin)
        item_set.SetItem("BottomMargin", margin)
        item_set.SetItem("LeftMargin", margin)
        item_set.SetItem("RightMargin", margin)
        item_set.SetItem("HeaderLen", margin)
        item_set.SetItem("FooterLen", margin)
        item_set.SetItem("GutterLen", margin)

        act.Execute(pset)

        return self.hwp

    def set_multi_col(self):

        #  단 설정
        # 파라미터셋._prop_map_get_.keys()
        # https://www.inflearn.com/community/questions/1077679/%ED%95%9C%EA%B8%80-%ED%8C%8C%EC%9D%B4%EC%8D%AC-%EB%B0%94%ED%83%95%EC%AA%BD-%EB%8B%A4%EB%8B%A8?srsltid=AfmBOorgKdk__dZPqReicG_VmRNLMrrEm6LP6nHglMw_9Ud5QRInHfSL
        pset = self.hwp.HParameterSet.HColDef
        self.hwp.HAction.GetDefault("MultiColumn", pset.HSet)
        pset.Count = 2
        pset.SameSize = 1
        pset.LineType = 1
        pset.SameGap = self.hwp.MiliToHwpUnit(8.0)  # <--
        pset.HSet.SetItem("ApplyClass", 832)
        pset.HSet.SetItem("ApplyTo", 6)
        self.hwp.HAction.Execute("MultiColumn", pset.HSet)
        # act.Execute(pset)

        # hwp.XHwpWindows.Active_XHwpWindow.Visible = False

    def publish_workbook(self):

        n_of_chapter = len(self.prob_type_often_list)
        file_path = ""

        for i in range(n_of_chapter):
            if len(self.prob_type_often_list[i]) == 0 :
                continue
            if self.hwp is None:
                self.hwp = pyhwpx.Hwp()
                # self.hwp.XHwpWindows.Item(0).Visible = False  # 기본값 = 백그라운드 해제
                # self.hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
                self.create_hwp_doc()
                # self.hwp.XHwpWindows.Item(1).Visible = False  # 기본값 = 백그라운드 해제
            else:
                self.create_hwp_doc()
                # self.hwp.Open("I:\\workspace\\중등\\[중등] 수학 template.hwpx")
                # self.hwp.XHwpWindows.Item(1).Visible = False  # 기본값 = 백그라운드 해제

            self.hwp.switch_to(0)
            self.set_multi_col()


            # 초다빈출
            # lst = sorted(self.prob_type_often_list[i] , reverse=True)
            self.prob_type_often_list[i].sort()
            for j in self.prob_type_often_list[i]:
                # file_path = os.path.join(self.working_path , self.prob_type_often_list[j])
                print(j)
                file_path=j

                self.hwp.switch_to(1)
                self.hwp.Open(file_path)

                is_del = False
                for ctrl in self.hwp.ctrl_list:
                    # if ctrl.UserDesc in ["그리기", "머릿말(양쪽)", "꼬릿말(양쪽)", "감추기", "선", "누름틀" ]:
                    #     print(ctrl.UserDesc)
                    #     self.hwp.Delete()
                    if ctrl.CtrlID == "head" or ctrl.CtrlID == "foot":
                        self.hwp.DeleteCtrl(ctrl)

                    if ctrl.UserDesc == "감추기":
                        self.hwp.DeleteCtrl(ctrl)

                    if ctrl.UserDesc == "그리기" and is_del is False:
                        self.hwp.DeleteCtrl(ctrl)
                        is_del = True


                self.hwp.SelectAll()
                self.hwp.MoveSelLeft()
                self.hwp.Copy()
                self.hwp.switch_to(0)
                self.hwp.Paste()
                self.hwp.switch_to(1)
                # self.hwp.XHwpDocuments.Item(1).Close(isDirty=False)
            # 최다 오답

            self.prob_type_wrong_anslist[i].sort()
            for j in self.prob_type_wrong_anslist[i]:
                print(j)
                file_path = j
                self.hwp.switch_to(1)
                self.hwp.Open(file_path)

                is_del = False
                for ctrl in self.hwp.ctrl_list:
                    # if ctrl.UserDesc in ["그리기", "머릿말(양쪽)", "꼬릿말(양쪽)", "감추기", "선", "누름틀" ]:
                    #     print(ctrl.UserDesc)
                    #     self.hwp.Delete()
                    if ctrl.CtrlID == "head" or ctrl.CtrlID == "foot":
                        self.hwp.DeleteCtrl(ctrl)

                    if ctrl.UserDesc == "감추기" :
                        self.hwp.DeleteCtrl(ctrl)

                    if ctrl.UserDesc == "그리기" and is_del is False:
                        self.hwp.DeleteCtrl(ctrl)
                        is_del = True

                self.hwp.SelectAll()
                self.hwp.MoveSelLeft()
                self.hwp.Copy()
                self.hwp.switch_to(0)
                self.hwp.Paste()
                self.hwp.switch_to(1)
                # self.hwp.XHwpDocuments.Item(1).Close(isDirty=False)
                # self.hwp.Clear(1)
            # 서술

            self.prob_type_writing_list[i].sort()
            for j in self.prob_type_writing_list[i]:
                print(j)
                file_path = j
                self.hwp.switch_to(1)
                self.hwp.Open(file_path)

                is_del = False
                for ctrl in self.hwp.ctrl_list:
                    # if ctrl.UserDesc in ["그리기", "머릿말(양쪽)", "꼬릿말(양쪽)", "감추기", "선", "누름틀" ]:
                    #     print(ctrl.UserDesc)
                    #     self.hwp.Delete()
                    if ctrl.CtrlID == "head" or ctrl.CtrlID == "foot":
                        self.hwp.DeleteCtrl(ctrl)

                    if ctrl.UserDesc == "감추기" :
                        self.hwp.DeleteCtrl(ctrl)

                    if ctrl.UserDesc == "그리기" and is_del is False:
                        self.hwp.DeleteCtrl(ctrl)
                        is_del = True

                self.hwp.SelectAll()
                self.hwp.MoveSelLeft()
                self.hwp.MoveSelLeft()
                self.hwp.MoveSelLeft()
                self.hwp.MoveSelLeft()

                self.hwp.Copy()
                self.hwp.switch_to(0)
                self.hwp.Paste()
                self.hwp.switch_to(1)
                # self.hwp.XHwpDocuments.Item(1).Close(isDirty=False)

            self.hwp.switch_to(0)
            self.set_multi_col()

            src_name = os.path.split(file_path)
            # print("src_name  " , src_name)
            # save_name = os.path.join(src_name[:-2], src_name[-2]+"-"+src_name[-1])

            # 경로를 폴더별로 분할
            parts = file_path.split(os.sep)

            # 학년, 단원 추출 (예: '3학년', '3단원')
            grade = ""
            unit = ""
            for part in parts:
                if "학년" in part:
                    grade = part
                elif "단원" in part:
                    unit = part

            # 저장할 새 경로
            new_dir = r"I:\workspace\중등"
            new_filename = f"{grade}-{unit}-tmp"
            save_path = os.path.join(new_dir, new_filename)

            print("!!!save_name + " + save_path)
            self.hwp.save_as(save_path + ".hwp", "HWP")
            self.hwp.save_as(save_path +".hwpx", "HWPX")
            # self.hwp.Run("FileClose")
            # self.hwp.XHwpDocuments.Item(0).Close(isDirty=False)
            self.hwp.quit()
            self.hwp = None



        return

    def print(self):
        for i in self.hwp_to_pdf_list:
            if self.hwp is None:
                self.hwp = pyhwpx.Hwp()
            self.hwp.open(i)
            self.hwp.save_as(i + ".pdf", "HWPX")







cnt = 3


for i in [
    # "F:\\download\\족보닷컴-수학\\족보닷컴-수학\\1학년\\1단원",
    # "F:\\download\\족보닷컴-수학\\족보닷컴-수학\\1학년\\2단원",
    # "F:\\download\\족보닷컴-수학\\족보닷컴-수학\\1학년\\3단원",
    # "F:\\download\\족보닷컴-수학\\족보닷컴-수학\\1학년\\4단원",
    # "F:\\download\\족보닷컴-수학\\족보닷컴-수학\\1학년\\5단원",
    # "F:\\download\\족보닷컴-수학\\족보닷컴-수학\\1학년\\6단원",
    "F:\\download\\족보닷컴-수학\\족보닷컴-수학\\3학년\\2단원"
          ]:
    maker = workbook_publisher(i, "F:\\download\\족보닷컴-수학\\족보닷컴-수학\\", "C:\\Data\\DB")
    maker.traverse_work()
    maker.publish_workbook()
    maker.print()
    del maker
