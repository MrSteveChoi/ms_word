import sys
import os
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QPushButton,
    QDialog, QGridLayout, QLabel, QCheckBox,
    QFileDialog, QVBoxLayout, QGroupBox, QWidget,
    QScrollArea, QDialogButtonBox, QHBoxLayout
)
from PyQt5.QtCore import Qt

from PyQt5.QtWidgets import QSpacerItem, QSizePolicy
from PyQt5.QtWidgets import QMessageBox, QLineEdit

from combine_wordfiles_with_spaces import group_docs_by_page, combine_all_docx, combine_all_docx_one_by_one, combine_all_answer_docx, combine_all_docx_seamless
from file_category_csv import add_categories_to_csv
from divide_questions import split_docx

# exe 파일이 위치한 절대경로 가져오기
exe_dir = os.path.dirname(sys.executable)
os.chdir(exe_dir)  # 프로그램 시작 전에 변경

exe_dir = os.getcwd()
split_dir = os.path.join(exe_dir, "split")
category_csv_dir = os.path.join(exe_dir, "csv_file/categories.csv")

# === 사용자 환경 설정 ===
WORK_DIR = os.getcwd()                                     # 현재 작업 디렉터리
# WORD_FILES_DIR = os.path.join(WORK_DIR, "word_files")      # 분할할 .docx 파일(문제/정답)이 있는 폴더
WORD_FILES_DIR = os.path.join(WORK_DIR, "word_files/test_word")      # 분할할 .docx 파일(문제/정답)이 있는 폴더
OUTPUT_DIR     = os.path.join(WORK_DIR, "split")         # 분할된 파일을 저장할 폴더

# pyinstaller --onefile --clean --add-data "C:\users\chy\.conda\envs\msword\lib\site-packages\docxcompose;docxcompose" --add-binary "C:\Users\CHY\.conda\envs\msword\Lib\site-packages\pandas.libs\msvcp140-0f2ea95580b32bcfc81c235d5751ce78.dll;." main.py

# pyinstaller --clean main_my.spec

# -------------------------------------------------------------
# 1) 공통 필터 함수
# -------------------------------------------------------------
def filter_files(dataframe, checkboxes_dict):
    """
    주어진 checkboxes_dict를 이용해 (AND 조건)으로 데이터프레임을 필터링한 뒤,
    file_name 컬럼 리스트를 반환합니다.
    
    - checkboxes_dict는 다음과 같은 구조:
      {
        'year': { 2017: QCheckBox객체, 2018: QCheckBox객체, ... },
        'difficulty': { 'hard': QCheckBox객체, 'easy': QCheckBox객체, ... },
        ...
      }
    - 각 컬럼(column)에 대해 체크된 값이 있다면 OR 조건으로 필터링,
      다른 컬럼과는 AND 조건으로 적용합니다.
    """
    if dataframe is None or dataframe.empty:
        return []
    
    filtered_df = dataframe.copy()
    
    for col_name, checkboxes in checkboxes_dict.items():
        # 이 컬럼에서 체크된 값들만 추출
        selected_values = [
            val for val, cb in checkboxes.items() if cb.isChecked()
        ]
        # 체크된 값이 하나도 없으면 필터를 적용하지 않음
        if selected_values:
            filtered_df = filtered_df[filtered_df[col_name].isin(selected_values)]
    
    return filtered_df["file_name"].tolist()


# -------------------------------------------------------------
# 2) 카테고리 필터 다이얼로그
# -------------------------------------------------------------
class CategoryDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("카테고리 선택 (csv_file 폴더의 categories.csv파일이 존재해야 합니다.)")
        self.setFixedSize(600, 500)

        self.initUI()
        
        # CSV 파일로부터 로드한 DataFrame
        self.df = pd.DataFrame()
        
        # 동적 체크박스를 담을 딕셔너리: {col_name: {value: QCheckBox}}
        self.checkboxes_dict = {}

        ### 임시 경로
        # csv파일 자동으로 읽기
        file_path = category_csv_dir # exe경로
        # file_path =r"D:/YearDreamSchool-D/python_projects/msword_pjt/csv_file/categories.csv" # .py경로

        if file_path:
            # self.csv_path_label.setText(f"Selected File: {file_path}")
            try:
                self.df = pd.read_csv(file_path)
                
                # 체크박스 UI 재구성
                self.create_dynamic_checkboxes()
                
            except Exception as e:
                QMessageBox.warning(self, "경고", f"csv파일이 존재하지 않습니다.\n올바른 경로에 csv파일을 위치해 주세요.")
                # self.csv_path_label.setText(f"Error: {e}")
        else:
            QMessageBox.warning(self, "경고", "csv파일이 존재하지 않습니다.")
            # QMessageBox.information(self, "알림", "병합이 완료되었습니다!")

        # 필터링 결과를 담을 변수
        self.filtered_result = []
    
    def initUI(self):
        """
        QDialog 내부의 레이아웃 구성
        """
        # 전체 레이아웃(GridLayout)
        self.layout = QGridLayout()
        self.setLayout(self.layout)
        
        # 1) CSV 파일 선택 버튼
        # self.csv_button = QPushButton("Select CSV File")
        # self.csv_button.clicked.connect(self.load_csv)
        # self.layout.addWidget(self.csv_button, 0, 0, 1, 2)
        
        # 2) CSV 경로 표시 라벨
        self.csv_path_label = QLabel("원하는 카테고리를 선택해 주세요.")
        self.layout.addWidget(self.csv_path_label, 0, 0, 1, 2)
        
        # 3) 체크박스들을 스크롤 영역에 담기 위한 준비
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        
        # 스크롤 영역 안에 들어갈 컨테이너 위젯
        self.checkboxes_container = QWidget()
        self.checkboxes_layout = QVBoxLayout()
        self.checkboxes_container.setLayout(self.checkboxes_layout)
        
        self.scroll_area.setWidget(self.checkboxes_container)
        # 스크롤 영역을 전체 레이아웃에 추가 (2행 아래)
        self.layout.addWidget(self.scroll_area, 1, 0, 1, 2)
        
        # 4) 필터 버튼
        self.filter_button = QPushButton("Filter")
        self.filter_button.clicked.connect(self.apply_filter)
        self.layout.addWidget(self.filter_button, 2, 0, 1, 2)
        
        # 5) 결과 라벨
        self.result_label = QLabel("Filtered Files: ")
        self.layout.addWidget(self.result_label, 3, 0, 1, 2)
        
        # 6) 다이얼로그 버튼(확인/취소 등)
        self.dialog_buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        self.dialog_buttons.accepted.connect(self.accept)   # "확인" 누르면 QDialog.Accepted
        self.dialog_buttons.rejected.connect(self.reject)   # "취소" 누르면 QDialog.Rejected
        self.layout.addWidget(self.dialog_buttons, 4, 0, 1, 2, alignment=Qt.AlignCenter)
        
    # def load_csv(self):
    #     """
    #     CSV 파일을 선택하고 DataFrame으로 로드한 뒤,
    #     'file_name'을 제외한 나머지 컬럼들에 대한 체크박스를 동적으로 생성한다.
    #     """
    #     file_path, _ = QFileDialog.getOpenFileName(self, "Select a CSV file", "", "CSV Files (*.csv)")
        
    #     if file_path:
    #         self.csv_path_label.setText(f"Selected File: {file_path}")
    #         try:
    #             self.df = pd.read_csv(file_path)
                
    #             # 체크박스 UI 재구성
    #             self.create_dynamic_checkboxes()
                
    #         except Exception as e:
    #             self.csv_path_label.setText(f"Error: {e}")
    #     else:
    #         self.csv_path_label.setText("No file selected")
    
    def create_dynamic_checkboxes(self):
        """
        self.df의 컬럼(헤더) 정보를 바탕으로
        file_name을 제외한 각 컬럼에 대해 유니크한 값의 체크박스를 동적으로 생성한다.
        """
        # 기존에 만들어진 체크박스 UI가 있다면 모두 제거
        for i in reversed(range(self.checkboxes_layout.count())):
            widget = self.checkboxes_layout.itemAt(i).widget()
            if widget:
                widget.setParent(None)
        
        # 딕셔너리도 초기화
        self.checkboxes_dict.clear()
        
        # (1) file_name 컬럼이 실제로 존재하는지 체크
        if 'file_name' not in self.df.columns:
            self.result_label.setText("Error: 'file_name' 컬럼이 없습니다.")
            return
        
        # (2) file_name 컬럼 제외한 나머지 컬럼만 처리
        for col_name in self.df.columns:
            if col_name == "file_name":
                continue

            # 문제 번호 column은 제외
            if col_name == "q_number":
                continue
            
            # 이 컬럼에 대한 unique한 값들
            unique_values = self.df[col_name].unique()
            
            # GroupBox(각 컬럼별로 묶음)
            group_box = QGroupBox(col_name)
            vbox = QVBoxLayout()
            group_box.setLayout(vbox)
            
            # 딕셔너리 구조: self.checkboxes_dict[col_name] = { val: checkBox, ... }
            self.checkboxes_dict[col_name] = {}
            
            for val in unique_values:
                cb = QCheckBox(str(val))
                vbox.addWidget(cb)
                self.checkboxes_dict[col_name][val] = cb
            
            # 완성된 GroupBox를 메인 체크박스 레이아웃에 추가
            self.checkboxes_layout.addWidget(group_box)
    
    def apply_filter(self):
        """
        생성된 동적 체크박스들을 확인해 필터 적용 후 결과를 표시한다.
        """
        self.filtered_result = filter_files(self.df, self.checkboxes_dict)
        
        if self.filtered_result:
            self.result_label.setText(f"Filtered Files: {', '.join(self.filtered_result)}")
        else:
            self.result_label.setText("Filtered Files: (No matched files)")
    
    def get_filtered_result(self):
        """
        필터링된 결과 리스트를 반환한다.
        """
        return self.filtered_result


# -------------------------------------------------------------
# 3) 문제 선택 다이얼로그
# -------------------------------------------------------------
class ProblemSelectDialog(QDialog):
    def __init__(self, file_list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("문제 선택")
        self.setFixedSize(400, 500)
        
        self.file_list = file_list  # 필터링된 파일 리스트
        self.checkboxes = {}
        self.selected_files = []
        
        # (추가) 병합 작업에 필요한 함수(가령 get_document_info, group_docs_by_page 등)를
        #        import 또는 동일 파일 내에 정의했다고 가정
        
        self.initUI()
    
    def initUI(self):
        layout = QGridLayout()
        self.setLayout(layout)
        
        # 안내 라벨
        info_label = QLabel("문제파일을 선택 후 병합")
        layout.addWidget(info_label, 0, 0, 1, 2)

        # 스크롤 영역
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setMinimumWidth(380)
        
        container = QWidget()
        self.vbox = QVBoxLayout()
        container.setLayout(self.vbox)
        
        self.scroll_area.setWidget(container)
        layout.addWidget(self.scroll_area, 1, 0, 1, 2)
        
        # word_files 폴더 내 존재 여부 확인 후 체크박스 생성
        self.create_file_checkboxes()

        ### ✅ "모두 선택" & "모두 선택 해제" 버튼 추가
        select_button_layout = QHBoxLayout()
        
        self.select_all_button = QPushButton("모두 선택")
        self.select_all_button.clicked.connect(self.select_all_files)

        self.deselect_all_button = QPushButton("모두 선택 해제")
        self.deselect_all_button.clicked.connect(self.deselect_all_files)

        select_button_layout.addWidget(self.select_all_button)
        select_button_layout.addWidget(self.deselect_all_button)

        layout.addLayout(select_button_layout, 2, 0, 1, 2)  # ✅ 한 줄에 배치
        ###

        # ✅ "결과물 이름" 입력 필드 추가
        output_name_layout = QHBoxLayout()

        result_name_label = QLabel("결과물 이름:")
        self.result_name_input = QLineEdit()
        self.result_name_input.setText("output")  # 기본값 설정 ✅

        output_name_layout.addWidget(result_name_label)  # 라벨 추가
        output_name_layout.addWidget(self.result_name_input)  # 입력 필드 추가

        layout.addLayout(output_name_layout, 3, 0, 1, 2)  # ✅ 한 줄에 배치


        ### ✅ 병합 버튼 + 체크박스를 하나의 레이아웃으로 묶기
        merge_layout = QHBoxLayout()

        # 병합 버튼
        # self.merge_button = QPushButton("병합(채워서)")
        # self.merge_button.clicked.connect(self.on_merge_clicked)
        # merge_layout.addWidget(self.merge_button)  

        # self.merge_button_one_by_one = QPushButton("병합(따로)")
        # self.merge_button_one_by_one.clicked.connect(self.on_merge_clicked_one_by_one)
        # merge_layout.addWidget(self.merge_button_one_by_one)

        self.merge_button = QPushButton("병합")
        self.merge_button.clicked.connect(self.on_merge_clicked_seamless)
        merge_layout.addWidget(self.merge_button) 



        # ✅ 병합 버튼과 체크박스 사이에 SpacerItem 추가 (체크박스 밀려나는 문제 해결)
        # merge_layout.addStretch(1)

        # ✅ 체크박스를 병합 버튼과 같은 레이아웃에 추가 (고정된 위치 유지)
        self.with_answer_checkbox = QCheckBox("정답도 같이")
        self.with_answer_checkbox.setChecked(True) # default : True 
        merge_layout.addWidget(self.with_answer_checkbox)  
        
        self.is_reset_num = QCheckBox("번호 리셋")
        self.is_reset_num.setChecked(True) # default : True 
        merge_layout.addWidget(self.is_reset_num) 

        layout.addLayout(merge_layout, 4, 0, 1, 3)  # ✅ 병합 버튼 + 체크박스를 같은 행(row=3)에 배치


        # 다이얼로그 버튼(확인/취소)
        # self.dialog_buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        # self.dialog_buttons.accepted.connect(self.on_accept)
        # self.dialog_buttons.rejected.connect(self.reject)
        # layout.addWidget(self.dialog_buttons, 3, 2, 1, 1, alignment=Qt.AlignRight)
    
    def create_file_checkboxes(self):
        """
        word_files 폴더 안에 존재하면 활성 체크박스,
        없으면 비활성(회색) 체크박스로 표시
        """
        # ------------------### 임시로 사용 ###------------------
        word_files_path = os.path.join(exe_dir, "split") # .exe 경로
        # word_files_path = r"D:\YearDreamSchool-D\python_projects\msword_pjt\split" # .py 경로
        # ------------------### 임시로 사용 ###------------------
        
        for file_name in self.file_list:
            full_path = os.path.join(word_files_path, file_name)
            cb = QCheckBox(file_name)
            
            if os.path.exists(full_path):
                # 존재하는 경우 - 활성
                cb.setEnabled(True)
            else:
                # 존재하지 않는 경우 - 비활성(회색)
                cb.setEnabled(False)
                cb.setStyleSheet("color: gray;")
            
            self.vbox.addWidget(cb)
            self.checkboxes[file_name] = cb

    ###
    def select_all_files(self):
        """ 모든 활성화된 체크박스를 선택 """
        for cb in self.checkboxes.values():
            if cb.isEnabled():
                cb.setChecked(True)

    def deselect_all_files(self):
        """ 모든 활성화된 체크박스를 선택 해제 """
        for cb in self.checkboxes.values():
            if cb.isEnabled():
                cb.setChecked(False)
    ###

    def on_accept(self):
        """
        "확인" 버튼을 누르면, 체크된 파일들만 self.selected_files에 담고 닫기
        """
        self.selected_files = [
            fname for fname, cb in self.checkboxes.items() if cb.isChecked() and cb.isEnabled()
        ]
        self.accept()  # QDialog.Accepted
    
    def on_merge_clicked(self):
        """
        "병합" 버튼을 누르면,
        1) 현재 선택된 파일 확인
        2) group_docs_by_page(base_path, file_list) 함수 호출
        """
        # 직접 적은 파일명 변수
        output_filename = self.result_name_input.text() 

        self.selected_files = [
            fname for fname, cb in self.checkboxes.items() if cb.isChecked() and cb.isEnabled()
        ]
        
        if not self.selected_files:
            # 선택된 파일이 없다면 경고 메시지 등을 띄울 수 있음
            QMessageBox.warning(self, "경고", "병합할 파일이 없습니다.")
            return
        
        # ------------------### 임시로 사용 ###------------------
        base_path = os.path.join(exe_dir, "split") # .exe 경로
        # base_path = r"D:\YearDreamSchool-D\python_projects\msword_pjt\split" # .py 경로
        # ------------------### 임시로 사용 ###------------------
        
        ### 문제 파일 병합
        # 문제의 길이를 계산해서 페이지를 나눈 list 반환
        files_list = group_docs_by_page(base_path, self.selected_files)
        # group화 된 list를 하나의 file로 병합
        combine_all_docx(base_path, files_list, output_filename, self.is_reset_num.isChecked())

        ### 정답 파일 병합
        # 정답도 같이 포함이면 선택된 파일들의 list를 기반으로 정답파일의 이름을 list화
        include_answers = self.with_answer_checkbox.isChecked()
        if include_answers:
            self.selected_files_answer = [os.path.splitext(i)[0]+"_MS.docx" for i in self.selected_files]
            # list를 하나의 file로 병합
            combine_all_answer_docx(base_path, self.selected_files_answer, output_filename+"_MS", self.is_reset_num.isChecked())
        
        
        QMessageBox.information(self, "알림", "병합 완료!")

        # 필요하다면 메시지 박스 등으로 알려줄 수 있음
        # QMessageBox.information(self, "알림", "병합이 완료되었습니다.")

    def on_merge_clicked_one_by_one(self):
        """
        "병합" 버튼을 누르면,
        1) 현재 선택된 파일 확인
        2) group_docs_by_page(base_path, file_list) 함수 호출
        """
        # 직접 적은 파일명 변수
        output_filename = self.result_name_input.text() 

        # 선택된 파일명 list
        self.selected_files = [
            fname for fname, cb in self.checkboxes.items() if cb.isChecked() and cb.isEnabled()
        ]
        
        if not self.selected_files:
            # 선택된 파일이 없다면 경고 메시지 등을 띄울 수 있음
            QMessageBox.warning(self, "경고", "병합할 파일이 없습니다.")
            return
        
        # ------------------### 임시로 사용 ###------------------
        base_path = os.path.join(exe_dir, "split") # .exe 경로
        # base_path = r"D:\YearDreamSchool-D\python_projects\msword_pjt\split" # .py 경로
        # ------------------### 임시로 사용 ###------------------
        
        ### 문제 파일 병합
        # 그룹화 할 필요가 없기때문에 file list를 바로 병합
        files_list = self.selected_files
        combine_all_docx_one_by_one(base_path, files_list, output_filename, self.is_reset_num.isChecked())

        ### 정답 파일 병합
        # 정답도 같이 포함이면 선택된 파일들의 list를 기반으로 정답파일의 이름을 list화
        include_answers = self.with_answer_checkbox.isChecked()
        if include_answers:
            self.selected_files_answer = [os.path.splitext(i)[0]+"_MS.docx" for i in self.selected_files]
            # group화 된 list를 하나의 file로 병합
            combine_all_answer_docx(base_path, self.selected_files_answer, output_filename+"_MS", self.is_reset_num.isChecked())
        
        QMessageBox.information(self, "알림", "병합이 완료되었습니다!")

    def on_merge_clicked_seamless(self):
        """
        "병합" 버튼을 누르면,
        1) 현재 선택된 파일 확인
        2) group_docs_by_page(base_path, file_list) 함수 호출
        """
        # 직접 적은 파일명 변수
        output_filename = self.result_name_input.text() 

        # 선택된 파일명 list
        self.selected_files = [
            fname for fname, cb in self.checkboxes.items() if cb.isChecked() and cb.isEnabled()
        ]
        
        if not self.selected_files:
            # 선택된 파일이 없다면 경고 메시지 등을 띄울 수 있음
            QMessageBox.warning(self, "경고", "병합할 파일이 없습니다.")
            return
        
        # ------------------### 임시로 사용 ###------------------
        base_path = os.path.join(exe_dir, "split") # .exe 경로
        # base_path = r"D:\YearDreamSchool-D\python_projects\msword_pjt\split" # .py 경로
        # ------------------### 임시로 사용 ###------------------
        
        ### 문제 파일 병합
        # 그룹화 할 필요가 없기때문에 file list를 바로 병합
        files_list = self.selected_files
        combine_all_docx_seamless(base_path, files_list, output_filename, self.is_reset_num.isChecked())

        ### 정답 파일 병합
        # 정답도 같이 포함이면 선택된 파일들의 list를 기반으로 정답파일의 이름을 list화
        include_answers = self.with_answer_checkbox.isChecked()
        if include_answers:
            self.selected_files_answer = [os.path.splitext(i)[0]+"_MS.docx" for i in self.selected_files]
            # group화 된 list를 하나의 file로 병합
            combine_all_answer_docx(base_path, self.selected_files_answer, output_filename+"_MS", self.is_reset_num.isChecked())
        
        QMessageBox.information(self, "알림", "병합이 완료되었습니다!")


# -------------------------------------------------------------
# 4) 메인 윈도우
# -------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("메인 윈도우")
        self.setFixedSize(300, 300)

        print(f"cwd:{os.getcwd()}")

        # '카테고리 열기'로 필터링한 결과를 저장
        self.filtered_files = []
        self.docx_files = []

        # (A) 카테고리 열기 버튼
        self.split_word_button = QPushButton("문제 분리", self)
        self.split_word_button.setGeometry(80, 40, 140, 40)
        self.split_word_button.clicked.connect(self.split_selected_docxs)

        # (A) 폴더 선택 버튼
        self.select_folder_button = QPushButton("DB 추가", self)
        self.select_folder_button.setGeometry(80, 100, 140, 40) # (윈도우x, y, 버튼x, y)
        self.select_folder_button.clicked.connect(self.add_db) # 함수 연결
        
        # (A) 카테고리 열기 버튼
        self.category_button = QPushButton("카테고리 CSV", self)
        self.category_button.setGeometry(80, 160, 140, 40)
        self.category_button.clicked.connect(self.open_category_dialog)
        
        # (B) 문제 선택 버튼
        self.select_problem_button = QPushButton("문제 선택 병합", self)
        self.select_problem_button.setGeometry(80, 220, 140, 40)
        ### ----------임시로 버튼확인용 주석처리----------
        # 초기에 비활성화해두고, 필터링 결과가 있을 때만 활성화
        self.select_problem_button.setEnabled(False)
        # self.select_problem_button.setEnabled(True)
        ### ----------임시로 버튼확인용 주석처리----------
        self.select_problem_button.clicked.connect(self.open_problem_select_dialog)

    def split_selected_docxs(self):
        import win32com.client as win32
        from pathlib import Path
        """
        폴더를 선택하고, 해당 폴더 내에 존재하는 모든 .docx 파일명을 가져온다.
        """
        folder_path = QFileDialog.getExistingDirectory(self, "Select a folder")

        word = win32.gencache.EnsureDispatch('Word.Application')
        word.Visible = False

        add_category_file_list = []

        if folder_path:  # 사용자가 폴더를 정상적으로 선택했다면
            # 폴더 내 파일 중 .docx 확장자를 가진 파일만 필터링
            self.docx_files = [f for f in os.listdir(folder_path) if f.endswith(".docx")]
            
            # 리스트가 비어있지 않다면 해당 파일 목록을 표시
            if self.docx_files:

                for file in self.docx_files:
                    # q랑 a 구분
                    if os.path.splitext(file)[0].split("_")[-1] == "MS":
                        q_a = "a"
                    else:
                        q_a = "q"

                    ### 경로가 안맞음. 수정중
                    # input_path = f"'{str(Path(folder_path) / file)}'"
                    # input_path = f"'{os.path.normpath(os.path.join(folder_path, file))}'"

                    input_path = Path(folder_path) / file
                    input_path = str(input_path)
                    # print(f"넘겨주기 전 경로:{input_path}")

                    # ----------------### 임시 경로 ###----------------
                    output_path = os.path.join(os.getcwd(), "split")
                    # output_path = r"D:\YearDreamSchool-D\python_projects\msword_pjt\split"
                    # ----------------### 임시 경로 ###----------------

                    result_file_names = split_docx(
                        input_path=input_path, 
                        output_dir=output_path, 
                        word_app=word, 
                        q_a=q_a
                        )
                    
                    ### categories.csv파일에 분류해서 넣을 파일들의 목록을 저장
                    add_category_file_list.extend(result_file_names)

                
                add_categories_to_csv(category_csv_dir, add_category_file_list, exe_dir)

                QMessageBox.information(None, "안내", "분할/카테고리 추가\n완료")
                return
        
            else:
                QMessageBox.warning(self, "경고", "선택한 폴더에 .docx 파일이 존재하지 않습니다.")
                return
            
        word.Quit()


    def add_db(self):
        """
        폴더를 선택하고, 해당 폴더 내에 존재하는 모든 .docx 파일명을 가져온다.
        """
        folder_path = QFileDialog.getExistingDirectory(self, "Select a folder")

        if folder_path:  # 사용자가 폴더를 정상적으로 선택했다면
            # 폴더 내 파일 중 .docx 확장자를 가진 파일만 필터링
            self.docx_files = [f for f in os.listdir(folder_path) if f.endswith(".docx")]
            
            # 리스트가 비어있지 않다면 해당 파일 목록을 표시
            if self.docx_files:

                add_categories_to_csv(category_csv_dir, self.docx_files, exe_dir)
                QMessageBox.information(None, "안내", "카테고리 추가 완료")
                
            else:
                self.result_label.setText(f"선택된 폴더: {folder_path}\n\nDOCX 파일이 없습니다.")

        else:
            # self.result_label.setText("폴더 선택이 취소되었습니다.")
            QMessageBox.information(self, "알림", "폴더 선택 취소")

    def open_category_dialog(self):
        """
        카테고리 필터 창을 모달로 열고, 닫힌 후 필터링 결과를 가져온다.
        """
        dialog = CategoryDialog(self)
        # exec_()로 모달 띄움
        result = dialog.exec_()
        
        if result == QDialog.Accepted:
            # 사용자가 확인(OK) 버튼으로 닫았다면, 필터 결과를 가져옴
            self.filtered_files = dialog.get_filtered_result()
            # 필터링 결과가 있다면 '문제 선택' 버튼 활성화
            if self.filtered_files:
                self.select_problem_button.setEnabled(True)
            else:
                self.select_problem_button.setEnabled(False)
        else:
            # 사용자가 취소(Cancel)로 닫았거나 그냥 닫았으면
            pass

    def open_problem_select_dialog(self):
        """
        문제 선택 창을 열어, 체크박스로 특정 문제를 선택할 수 있게 한다.
        """
        ### ----------임시로 버튼확인용 주석처리----------
        if not self.filtered_files:
            return  # 필터링된 목록이 없으면 그냥 리턴
        ### ----------임시로 버튼확인용 주석처리----------
        
        dialog = ProblemSelectDialog(self.filtered_files, self)
        result = dialog.exec_()
        
        if result == QDialog.Accepted:
            ## 최종적으로 선택된 merge할 파일 목록
            selected_files = dialog.get_selected_files()
            if selected_files:
                print("사용자가 선택한 파일들:", selected_files)
            else:
                print("선택된 파일이 없습니다.")
        else:
            print("문제 선택을 취소했습니다.")
            

# -------------------------------------------------------------
# 5) 실행 부분
# -------------------------------------------------------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
