import os
import sys

from pyhwpx import Hwp


def hwp_to_pdf(hwp_path):
    if not os.path.exists(hwp_path):
        print("❌ 파일이 존재하지 않습니다.")
        return

    try:
        hwp = Hwp()
        hwp.open(hwp_path)

        output_path = os.path.splitext(hwp_path)[0] + ".pdf"
        hwp.save_as(output_path)
        print(f"✅ PDF 저장 완료: {output_path}")
        hwp.quit()

    except Exception as e:
        print("⚠️ 변환 중 오류 발생:", e)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        hwp_file = sys.argv[1]
        hwp_to_pdf(hwp_file)
    else:
        print("ℹ️ 변환할 .hwp 또는 .hwpx 파일을 드래그해서 실행해주세요.")