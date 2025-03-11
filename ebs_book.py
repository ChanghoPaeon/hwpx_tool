import pyhwpx
import os
hwp = pyhwpx.Hwp()

# 보안팝업 자동클릭
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")


def extract_from_ebs(file_path, out_path="I:\\workspace\\ebs\\"):

    # hwp.XHwpWindows.Active_XHwpWindow.Visible = False

    ret = True



    hwp.Open(file_path)

    hwp.MoveDocBegin()

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


file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS 2026학년도 수능특강 수학영역  수학Ⅱ_(185).hwp"
# file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS test.hwp"
file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS 2026학년도 수능특강 수학영역 수학Ⅰ_(182).hwp"
file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS 2026학년도 수능특강 수학영역 확률과 통계_(186).hwp"
# file_path = "F:\\git_repo\\hwpx_tool\\sample\\EBS 2026학년도 수능특강 수학영역 미적분_(172).hwp"

file_path = "I:\\학원-과외\\학원-부산\\문제집-일반\\EBS\\기출의 미래\\EBS 2026학년도 수능 기출의 미래 수학Ⅱ_(314).hwp"

extract_from_ebs(file_path)