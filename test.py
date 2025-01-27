import sys, os

def resource_path(relative_path):
    """
    PyInstaller로 빌드된 exe 내부/외부 어느 경우든,
    상대 경로로부터 절대 경로를 얻어 반환하는 헬퍼 함수입니다.
    """
    try:
        # PyInstaller 임시 폴더(실행 파일 포함) 경로
        base_path = sys._MEIPASS
    except Exception:
        # PyInstaller로 빌드되지 않은 경우(일반 Python 실행)
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)


print(resource_path("word_files"))