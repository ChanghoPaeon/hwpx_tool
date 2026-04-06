from pathlib import Path
from datetime import datetime

import pyhwpx
from pyhwpx import Hwp


hwp = pyhwpx.Hwp()


def convert_hwp_to_pdf(root_dir):
    root = Path(root_dir)

    hwp = Hwp(visible=False)


    for file in root.rglob("*"):
        if file.suffix.lower() in [".hwp", ".hwpx"]:

            try:
                print("convert:", file)

                # 날짜시간 문자열
                now = datetime.now().strftime("%Y%m%d-%H%M%S")

                # 새 파일명
                pdf_path = file.with_name(f"{file.stem}-{now}.pdf")

                # 파일 열기
                hwp.Open(str(file))

                # PDF 저장
                hwp.SaveAs(str(pdf_path), "PDF")

                # 문서 닫기
                hwp.Clear()

            except Exception as e:
                print("error:", file, e)

    hwp.Quit()



if __name__ == "__main__":
    convert_hwp_to_pdf(r"I:\\학원-과외\\20 과고기출\\00 울산과학고 울산과고")