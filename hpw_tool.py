import os
import tkinter as tk
import zipfile
from tkinter.filedialog import askopenfilename

import win32com.client as win32

# hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행
# hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안팝업 자동클릭

# hwp.XHwpWindows.Item(0).Visible = True  # 백그라운드 작업
# FILE_PATH = r"C:\Users\smj02\OneDrive\바탕 화면\빈 문서1.hwp"


def 파일선택():  # 파일선택 GUI 추가
    win = tk.Tk()
    win.withdraw()
    file_path = askopenfilename(title="hwp->hwpx..",
                                # initialdir=os.getcwd(),
                                initialdir="I:\\workspace\\새 폴더",
                                filetypes=[("한/글 파일", "*.hwp *.hwpx")])
    return file_path


def hwpx로_저장(hwp, path):
    hwp.Open(path)  # 문서 불러오기
    hpwx_path = hwp.Path + "x"
    hwp.SaveAs(hpwx_path, "HWPX")  # hwpx로 저장
    # hwp.Quit()
    return hpwx_path



def 압축해제(path):
    print("start 압축해제: " + path)
    # hwpx 인지 체크하기
    os.chdir(os.path.dirname(path))  # 파일경로 중 경로만 추출해서 그 경로로 이동
    target_path = os.path.join(os.getcwd(), path.split("\\")[-1]+"unpacked")  # 압축 풀 하위폴더는 "./hwpx"
    with zipfile.ZipFile(path, 'r') as zf:  # 압축파일 인스턴스를 생성하고
        zf.extractall(path=target_path)  # target_path 안에 압축해제
    # os.remove(path + "x")  # hwpx 파일은 단호하게 삭제함
    print("end 압축해제: " + target_path)
    return target_path


def delete_NGD_by_raw(hwpx_path):
    target_path = 압축해제(hwpx_path)
    target_path += "\\Contents\\section0.xml"

def delete_NGD_by_api(hwp, hwpx_path):
    print("start delete_NGD_by_api: " + hwpx_path)
    hwp.Open(hwpx_path)  # 문서 불러오기
    ctrl = hwp.HeadCtrl
    table_dict = {}
    index = 1

    while ctrl:
        if ctrl.UserDesc == "표":
            table_dict[index] = ctrl.GetAnchorPos(0)
        # print(f"Table {index}: Position {ctrl.GetAnchorPos(0)}")
            index += 1
        ctrl = ctrl.Next

    selected_pos = table_dict.get(1)
    if selected_pos:
        hwp.SetPosBySet(selected_pos)
    hwp.FindCtrl()
    hwp.HAction.Run("Delete")
    hwp.SaveAs(hwpx_path+".del_by_api.hwpx", "HWPX")  # hwpx로 저장
    hwp.Quit()
    print("end delete_NGD_by_api: " + hwpx_path)

def hwp_init(filename):  # 한/글 여는 코드가 길어서 미리 만들어둠
    hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")  # 한/글 객체 생성
    hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안모듈 실행
    hwp.Open(filename)  # GUI에서 선택한 파일 열기
    hwp.XHwpWindows.Item(0).Visible = True  # 한/글 창 숨김해제(초기에는 백그라운드상태)
    # hwp.HAction.Run("FrameFullScreen")  # 전체화면
    return hwp  # hwp객체 리턴


def hwp_unit_to_mm(hwp_unit):  # MiliToHwpUnit메서드의 반대. 계산 귀찮아서 만들어둠.
    return hwp_unit / 7200 * 25.4  # 1 inch = 7200 HWPUNIT, 1mm = 283.465 HWPUNIT


def move_to_ctrl(hwp, ctrl):  # 그리기객체나 표 등 컨트롤 오브젝트로 이동하는 함수
    pos_set = ctrl.GetAnchorPos(0)  # 현재 선택된 컨트롤의 위치파라미터 추출
    pos = (pos_set.Item("List"), pos_set.Item("Para"), pos_set.Item("Pos"))  # 튜플로 변환
    hwp.SetPos(*pos)  # 해당 위치로 커서(캐럿) 이동



def adjust_width(hwp):
    page_action = hwp.CreateAction("PageSetup")  # 페이지셋업 액션 실행준비
    page_set = page_action.CreateSet()  # 페이지 설정을 위한 파라미터 배열(비어있음) 생성
    page_action.GetDefault(page_set)  # 파라미터에 현재문서의 값을 채워넣음
    종이너비 = hwp_unit_to_mm(page_set.Item("PageDef").Item("PaperWidth"))  # 채워넣은 값 중 종이너비
    좌측여백 = hwp_unit_to_mm(page_set.Item("PageDef").Item("LeftMargin"))  # 채워넣은 값 중 좌측여백
    우측여백 = hwp_unit_to_mm(page_set.Item("PageDef").Item("RightMargin"))  # 채워넣은 값 중 우측여백
    제본여백 = hwp_unit_to_mm(page_set.Item("PageDef").Item("GutterLen"))  # 채워넣은 값 중 제본여백
    제본타입 = hwp_unit_to_mm(page_set.Item("PageDef").Item("GutterType"))  # 0:한쪽, 1:맞쪽, 2: 위쪽

    print("종이너비: " + str(종이너비))
    print("좌측여백: " + str(좌측여백))
    print("우측여백: " + str(우측여백))
    print("제본여백: " + str(제본여백))
    print("제본타입: " + str(제본타입))


    ctrl = hwp.HeadCtrl  # 문서 중 첫번째 컨트롤 선택
    while ctrl != None:  # 마지막 컨트롤까지 순회할 것.
        if ctrl.CtrlID == "gso":  # 컨트롤아이디가 그리기객체(gso)이면
            move_to_ctrl(hwp, ctrl)  # 위에서 정의한 이동함수
            hwp.FindCtrl()  # 해당 객체 선택
            이미지액션 = hwp.CreateAction("ShapeObjDialog")  # 이미지수정액선 실행준비
            이미지세트 = 이미지액션.CreateSet()  # 이미지수정을 위한 파라미터 배열(비어있음) 생성
            이미지액션.GetDefault(이미지세트)  # 빈 파라미터 배열에 현재문서의 값을 채워넣음
            기존너비 = hwp_unit_to_mm(이미지세트.Item("Width"))  # 선택 이미지의 현재너비 저장
            기존높이 = hwp_unit_to_mm(이미지세트.Item("Height"))  # 선택 이미지의 현재높이 저장
            변경쪽너비 = 종이너비-좌측여백-우측여백-(0 if 제본타입==2 else 제본여백)  # 변경할 쪽너비 계산
            변경쪽너비 = 변경쪽너비/2.2

            if 기존너비 > 변경쪽너비:
                hwp.HParameterSet.HShapeObject.HSet.SetItem("Width", hwp.MiliToHwpUnit(변경쪽너비))  # 이미지너비 변경값 입력
                hwp.HParameterSet.HShapeObject.HSet.SetItem("Height", hwp.MiliToHwpUnit(기존높이*변경쪽너비/기존너비))
                hwp.HAction.Execute("ShapeObjDialog", hwp.HParameterSet.HShapeObject.HSet)

        else:  # 컨트롤아이디가 그리기객체가 아니면
            pass  # 그냥 넘어가기
        ctrl = ctrl.Next  # 다음 컨트롤로 이동


file_path = "I:\\workspace\\ebs\\"
file_name = "EBS 2026학년도 수능 기출의 미래 수학Ⅱ_(314).hwp-new-"+".hwpx"

hwp = hwp_init(file_path+file_name)

adjust_width(hwp)

hwp.SaveAs(file_path+file_name, "HWPX")

