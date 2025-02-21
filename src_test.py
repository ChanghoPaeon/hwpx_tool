import pyhwpx
import re
import pyperclip

hwp = pyhwpx.Hwp()

hwp.XHwpWindows.Active_XHwpWindow.Visible = False
# 07030
pattern = r"To\s*\n\d{5}"
# pattern = r"To\d{5}"

file_path = "F:\\download\\EBS 2023 올림포스 전국연합학력평가 기출문제집 미적분 (1).hwp"

# file_path="F:\\git_repo\\hwpx_tool\\[고][2021][1-1-b][울산][강남고][수상][복소수-도형의이동][02035][00002][99001][02010][그림1-5-0-0].hwp"

file_path ="F:\git_repo\hwpx_tool\[고][2022][1-2-b][경기과천시][과천중앙고][수하][신사고][함수-조합][04017][10002][04001][그림7-0-7-0].hwp"



# file_path = "F:\git_repo\hwpx_tool\\test.hwp"
doc = hwp.open(file_path)
# content = hwp.GetTextFile()

hwp.MoveDocBegin()

for ctrl in hwp.ctrl_list:
    # print(ctrl)
    # 미주 미주

    # print(ctrl.UserDesc)
    if ctrl.UserDesc == "수식":
        ret = hwp.select_ctrl(ctrl)
        hwp.move_to_ctrl(ctrl)
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
            result = re.sub(pattern,"",matheq)
            print("=================== 바꾸고 난 후")
            result = re.sub(r"\s+$", " ", result)
            # # stripped = result.rstrip() + " " if result.strip() != result.rstrip() else result
            print(result)
            print("=================== ")

            ctrl.Properties.SetItem("String",result)
            ctrl.Properties.SetItem("VisualString", result)
            hwp.HAction.GetDefault("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
            # hwp.HParameterSet.SetItem("String", result)
            print(hwp.HParameterSet.HEqEdit.string)
            print(hwp.HParameterSet.HEqEdit)
            hwp.HParameterSet.HEqEdit.VisualString = result
            print(hwp.HParameterSet.HEqEdit.VisualString)
            print(hwp.HParameterSet.HEqEdit.Version)
            hwp.HAction.Execute("EquationModify", hwp.HParameterSet.HEqEdit.HSet)
            # hwp.HAction.Run("SelectAll")
            # 수식 하나 바꾼 후 죽음 이게 원인인가
            hwp.HAction.Run("Delete")

            # pset = hwp.HParameterSet.HEqEdit
            # hwp.HAction.GetDefault("EquationModify", pset.HSet)
            # pset.string = pset.VisualString
            # pset.Version = "Equation Version 60"
            # hwp.HAction.Execute("EquationModify", pset.HSet)


            hwp.HAction.GetDefault("EquationCreate", hwp.HParameterSet.HEqEdit.HSet)
            hwp.HParameterSet.HEqEdit.EqFontName = "HancomEQN"
            hwp.HParameterSet.HEqEdit.string = result
            hwp.HParameterSet.HEqEdit.BaseUnit = hwp.PointToHwpUnit(11.0)
            hwp.HAction.Execute("EquationCreate", hwp.HParameterSet.HEqEdit.HSet)


            # hwp.HAction.GetDefault("EquationPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)
            # hwp.HParameterSet.HShapeObject.string = result
            # hwp.HAction.Execute("EquationPropertyDialog", hwp.HParameterSet.HShapeObject.HSet)

            #
            # hwp.MoveRight()
            # hwp.BreakPara()
            # hwp.set_font(TextColor="Red")
            # hwp.insert_text(ctrl.Properties.Item("String"))


        # if matheq.find("N.G.D") !=-1 or matheq.find("NGD") != -1:
        #     print(ctrl.Properties.Item("String"))

hwp.save_as(file_path+".hwp")

hwp.XHwpWindows.Active_XHwpWindow.Visible = True

print("end")