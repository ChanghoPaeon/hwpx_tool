
import hpw_tool
import tools
import hwp_format_converter



# hwp_tool 사용예제
# path = hpw_tool.파일선택()
# hpwx_path = hpw_tool.hwpx로_저장(path)
# path = hpwx_path
# hpw_tool.압축해제(path)
# hpw_tool.delete_NGD_by_api(path)
# hpw_tool.hwp.Quit()  # 한/글 종료

import pyhwpx
hwp = pyhwpx.Hwp()

import statistics
import configparser
config = configparser.ConfigParser() # 유니코드 문자열로 처리
with open('config.ini', 'r',encoding='utf-8') as f:
    config.read_file(f)


import win32com.client as win32



# 사용 예제
# folder_path = "I:\\test\\[NGD]179차 기출 최종완성본 (총 376파일)_2023.07.26 한글본\\"  # 검색할 폴더 경로
# keywords = ["[난이도] 특", "[난이도] 상", "[난이도] 중", "[난이도] 하", "검색어5"]  # 찾고 싶은 단어 목록
# output_csv = "I:\\test\\"+ "[NGD]179차"+".csv"  # 결과 CSV 파일명
#
# search_keywords_in_folder(folder_path, keywords, output_csv)

hwp.XHwpWindows.Active_XHwpWindow.Visible = False

if __name__ == '__main__':
    keywords = ["[난이도] 킬러","[난이도] 킬","[난이도] 특", "[난이도] 상"]  # 찾고 싶은 단어 목록
    # keywords = ["[난이도] 킬러", "[난이도] 킬", "[난이도] 특", "[난이도] 상", "[난이도] 중", "[난이도] 하"]  # 찾고 싶은 단어 목록

    output_csv_path = config.get("path", "output_csv_path")  # 결과 CSV 파일명

    folder_path = config.get("path", "folder_path")# 검색할 폴더 경로

    import datetime

    suffix = datetime.datetime.now().strftime('%y%m%d_%H%M%S')


    # 난이도 search
    searcher = statistics.HwpKeywordSearcher(hwp, folder_path, keywords, output_csv_path, suffix)
    searcher.traverse_work()
    hwp.XHwpWindows.Active_XHwpWindow.Visible = True

    # converter = hwp_format_converter.NGD_converter(hwp, folder_path, keywords, output_csv_path, suffix)
    # converter.traverse_work()

    # tools.find_and_delete_duplicates(folder_path)