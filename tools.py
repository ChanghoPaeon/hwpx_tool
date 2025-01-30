import os
import hashlib


def get_file_hash(file_path):
    """파일의 SHA-256 해시 값을 계산합니다."""
    hasher = hashlib.sha256()
    try:
        with open(file_path, "rb") as f:
            while chunk := f.read(4096):
                hasher.update(chunk)
    except Exception as e:
        print(f"해시 계산 중 오류 발생: {file_path}, 오류: {e}")
        return None
    return hasher.hexdigest()


def find_and_delete_duplicates(directory):
    """
    특정 경로 내에서 동일한 파일을 찾아 중복된 파일을 삭제합니다.

    :param directory: 탐색할 폴더 경로
    """
    if not os.path.exists(directory):
        print(f"경로가 존재하지 않습니다: {directory}")
        return False

    file_hashes = {}  # {해시값: 파일경로 리스트}

    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            file_hash = get_file_hash(file_path)

            if file_hash:
                if file_hash in file_hashes:
                    file_hashes[file_hash].append(file_path)
                else:
                    file_hashes[file_hash] = [file_path]

    # 중복된 파일 삭제 (하나만 남기고 나머지 삭제)
    for hash_value, file_list in file_hashes.items():
        if len(file_list) > 1:
            print(f"\n🔹 중복 파일 발견 (해시: {hash_value[:10]}...):")
            for i, f in enumerate(file_list):
                print(f"  {i + 1}. {f}")

            # 첫 번째 파일을 남기고 나머지는 삭제
            for duplicate in file_list[1:]:
                try:
                    os.remove(duplicate)
                    print(f"🗑️ 삭제 완료: {duplicate}")
                except Exception as e:
                    print(f"❌ 삭제 실패: {duplicate}, 오류: {e}")

    print("\n✅ 중복 파일 정리 완료!")
    return True


# ✅ 실행 예제: 원드라이브 문서 폴더 내 중복 파일 삭제
onedrive_path = os.path.expanduser("~/OneDrive")  # 기본 원드라이브 경로
target_folder = os.path.join(onedrive_path, "문서")  # 검사할 폴더

find_and_delete_duplicates(target_folder)