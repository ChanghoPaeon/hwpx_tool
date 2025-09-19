import pyhwpx
import os
import hwp_tool
import datetime
from pyhwpx import Hwp
from pathlib import Path

hwp = pyhwpx.Hwp()

# 보안팝업 자동클릭
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")

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
                    print(f"  {file_name}")


            print("=" * 50)


    def create_hwp_doc(self):
        hwp.add_doc()
        # hwp.Open(file_path + ".hwp")

        # 페이지 여백 설정
        act = hwp.CreateAction("PageSetup")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("ApplyTo", 3)

        item_set = pset.CreateItemSet("PageDef", "PageDef")
        margin = hwp.MiliToHwpUnit(5)
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

    def publish_workbook(self):
        print("Reading DB Start")
        df = pd.read_csv("./db/db.csv", encoding='euc-kr')
        q_result = df.query('문제타입 == "lv1" or 문제타입 == "lv2"')

        prob_lv = "lv2"
        prob_subj = "확통"
        qeury = "문제타입 == \"" + prob_lv + "\" and " + "과목 ==" +"\""+ prob_subj+"\""
        print(qeury)
        q_result = df.query(qeury)
        prob_type = "문제타입 == " + prob_lv

        # type_order = {"ex": 0, "lv1": 1, "lv2": 2, "lv3": 3}
        # df_sorted = q_result.sort_values(by=["문제집", "문제타입"], ascending=[False, True], key=lambda x: x.map(type_order))

        type_order = {"ex": 0, "lv1": 1, "lv2": 2, "lv3": 3}
        df_sorted = q_result.sort_values(by=["문제집", "문제번호"], ascending=[False, True])
        self.create_hwp_doc()
        hwp.switch_to(0)
        self.set_multi_col()

        hwp.switch_to(1)


        for _, row in df_sorted.iterrows():
            # print(f"문제집: {row['문제집']}, 문제타입: {row['문제타입']}")
            # print(self.db_path)
            current_prob = str(row['문제집']) + "-" +  str(row['문제번호']).zfill(4)+"-A"
            print("!!! >>  " + current_prob )
            print(os.path.join(self.db_path), current_prob)


            folder = Path(self.db_path)
            file_path = folder / f"{row['문제집']}-{str(row['문제번호']).zfill(4)}-A"

            hwp.Open(str(file_path))
            hwp.SelectAll()
            hwp.Copy()


            hwp.switch_to(0)

            hwp.Run("BreakColumn")
            # 미주 넣기
            hwp.Run("InsertEndnote")
            hwp.Paste()


            # 미주 빠져나오기
            hwp.MoveNextPosEx()
            hwp.MoveNextPosEx()


            hwp.switch_to(1)
            hwp.Clear(1)

            # hwp.switch_to(0)

            file_path = folder / f"{row['문제집']}-{str(row['문제번호']).zfill(4)}-Q"

            hwp.Open(str(file_path))
            hwp.SelectAll()
            hwp.Copy()


            hwp.switch_to(0)
            hwp.insert_text(" " + str(row['문제집']) + "-" +  str(row['문제번호']).zfill(4))
            hwp.BreakPara()
            hwp.insert_text( " " + str(row['문제타입']) )
            hwp.Paste()

            hwp.switch_to(1)
            hwp.Clear(1)

            # hwp.switch_to(0)

        hwp.switch_to(0)

        hwp.Run("BreakPage")

        # 맨 앞으로 가서 del 두번
        hwp.MoveDocBegin()
        hwp.Delete()
        hwp.Delete()


        hwp_tool.adjust_image_width(hwp)
        hwp_tool.adjust_table_width(hwp)

        save_name = "E:\\workspace\\ebs\\" + "수특-"+prob_subj+"-"+prob_lv
        hwp.save_as(save_name + ".hwp", "HWP")
        hwp.save_as(save_name +".hwpx", "HWPX")

        return












maker = workbook_publisher("C:\\Data\\", "C:\\Data\\", "C:\\Data\\DB")
maker.publish_workbook()