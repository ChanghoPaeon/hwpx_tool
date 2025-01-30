import os
import tkinter as tk
import zipfile
from tkinter.filedialog import askopenfilename

import win32com.client as win32

hwp = win32.gencache.EnsureDispatch("hwpframe.hwpobject")  # 한/글 실행
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")  # 보안팝업 자동클릭

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


def hwpx로_저장(path):
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

def count_string(path, string):
    return


def delete_table():
    return

def delete_NGD_by_raw(hwpx_path):
    target_path = 압축해제(hwpx_path)
    target_path += "\\Contents\\section0.xml"

def delete_NGD_by_api(hwpx_path):
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


