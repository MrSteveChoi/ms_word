from PyInstaller.utils.hooks import collect_submodules, collect_data_files

# Pandas의 모든 서브모듈을 강제 포함
hiddenimports = ["pandas", 'pandas._libs.tslibs.timedeltas']

# Pandas 데이터 파일 포함 (예: CSV, JSON 등)
datas = collect_data_files("pandas")

# Pandas에서 필요한 모든 DLL 포함
binaries = collect_data_files("pandas._libs", include_py_files=True)
