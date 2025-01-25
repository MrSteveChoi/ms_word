import sys
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QDialog,
    QPushButton, QVBoxLayout, QHBoxLayout, QMessageBox,
    QLabel, QCheckBox, QDialogButtonBox
)


class CategoryDialog(QDialog):
    """카테고리 선택 창"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("카테고리 선택")
        self.resize(300, 200)

        layout = QVBoxLayout()

        # 예시 체크박스
        self.case1 = QCheckBox("Case 1")
        self.case2 = QCheckBox("Case 2")
        self.case3 = QCheckBox("Case 3")

        layout.addWidget(self.case1)
        layout.addWidget(self.case2)
        layout.addWidget(self.case3)

        # 확인 / 취소 버튼
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.accept_clicked)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self.setLayout(layout)

    def accept_clicked(self):
        # 예시 로직: 카테고리 선택 결과 확인
        checked_list = []
        if self.case1.isChecked():
            checked_list.append("Case 1")
        if self.case2.isChecked():
            checked_list.append("Case 2")
        if self.case3.isChecked():
            checked_list.append("Case 3")

        QMessageBox.information(self, "카테고리 확정", f"선택된 카테고리: {checked_list}")
        self.accept()


class MergeDialog(QDialog):
    """Merge 옵션 창"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Merge 옵션")
        self.resize(300, 220)

        layout = QVBoxLayout()

        # 예시 체크박스
        self.chk_dup = QCheckBox("중복 제거")
        self.chk_file = QCheckBox("파일 합치기")
        self.chk_outchk = QCheckBox("출력폴더 체크")
        self.chk_outpath = QCheckBox("Output 경로")

        layout.addWidget(self.chk_dup)
        layout.addWidget(self.chk_file)
        layout.addWidget(self.chk_outchk)
        layout.addWidget(self.chk_outpath)

        # 확인 버튼
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.accepted.connect(self.on_accept)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self.setLayout(layout)

    def on_accept(self):
        # 예시 로직: 체크박스 상태 확인
        selections = []
        if self.chk_dup.isChecked():
            selections.append("중복 제거")
        if self.chk_file.isChecked():
            selections.append("파일 합치기")
        if self.chk_outchk.isChecked():
            selections.append("출력폴더 체크")
        if self.chk_outpath.isChecked():
            selections.append("Output 경로")

        QMessageBox.information(self, "Merge 옵션", f"설정된 옵션: {selections}")
        self.accept()


class QuestionDialog(QDialog):
    """문제 분리 창"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("문제 분리")
        self.resize(300, 180)

        layout = QVBoxLayout()
        label = QLabel("문제 분리 옵션 예시")
        layout.addWidget(label)

        btn = QPushButton("분리 실행")
        btn.clicked.connect(self.execute_split)
        layout.addWidget(btn)

        self.setLayout(layout)

    def execute_split(self):
        # 예시 동작
        QMessageBox.information(self, "문제 분리", "문제 분리 실행됨")


class AnswerDialog(QDialog):
    """정답 분리 창"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("정답 분리")
        self.resize(300, 180)

        layout = QVBoxLayout()
        label = QLabel("정답 분리 옵션 예시")
        layout.addWidget(label)

        btn = QPushButton("분리 실행")
        btn.clicked.connect(self.execute_split)
        layout.addWidget(btn)

        self.setLayout(layout)

    def execute_split(self):
        # 예시 동작
        QMessageBox.information(self, "정답 분리", "정답 분리 실행됨")


class CSVDialog(QDialog):
    """CSV 파일로 창"""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("CSV 파일로")
        self.resize(300, 180)

        layout = QVBoxLayout()
        label = QLabel("CSV 저장 옵션 예시")
        layout.addWidget(label)

        btn = QPushButton("저장 실행")
        btn.clicked.connect(self.save_csv)
        layout.addWidget(btn)

        self.setLayout(layout)

    def save_csv(self):
        # 예시 동작
        QMessageBox.information(self, "CSV 저장", "CSV로 저장 실행됨")


class MainWindow(QMainWindow):
    """메인 창"""
    def __init__(self):
        super().__init__()
        self.setWindowTitle("메인 GUI")
        self.resize(500, 300)

        # 중앙 위젯 (전체 레이아웃을 담을 컨테이너)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)

        # 메인 레이아웃(수직)
        main_layout = QVBoxLayout(central_widget)

        # 상단: 왼쪽과 오른쪽에 버튼 2개씩 배치 (QHBoxLayout 안에 QVBoxLayout 2개)
        top_layout = QHBoxLayout()

        # 왼쪽 버튼 2개
        left_layout = QVBoxLayout()
        self.btn_category = QPushButton("카테고리 선택")
        self.btn_merge = QPushButton("Merge 옵션")
        left_layout.addWidget(self.btn_category)
        left_layout.addWidget(self.btn_merge)

        # 오른쪽 버튼 2개
        right_layout = QVBoxLayout()
        self.btn_q_split = QPushButton("문제 분리")
        self.btn_a_split = QPushButton("정답 분리")
        right_layout.addWidget(self.btn_q_split)
        right_layout.addWidget(self.btn_a_split)

        # top_layout에 왼쪽/오른쪽 레이아웃 추가
        top_layout.addLayout(left_layout)
        top_layout.addSpacing(50)  # 왼쪽, 오른쪽 간격 조절(필요시)
        top_layout.addLayout(right_layout)

        # 하단: CSV 버튼 가운데 배치
        bottom_layout = QHBoxLayout()
        self.btn_csv = QPushButton("CSV 파일로")
        bottom_layout.addStretch()
        bottom_layout.addWidget(self.btn_csv)
        bottom_layout.addStretch()

        # 메인 레이아웃에 top_layout, bottom_layout 순서대로 추가
        main_layout.addLayout(top_layout)
        main_layout.addStretch()    # 위/아래 공간 확장
        main_layout.addLayout(bottom_layout)

        # 버튼 클릭 시 다이얼로그 열기
        self.btn_category.clicked.connect(self.open_category_dialog)
        self.btn_merge.clicked.connect(self.open_merge_dialog)
        self.btn_q_split.clicked.connect(self.open_question_dialog)
        self.btn_a_split.clicked.connect(self.open_answer_dialog)
        self.btn_csv.clicked.connect(self.open_csv_dialog)

    def open_category_dialog(self):
        dialog = CategoryDialog(self)
        dialog.exec_()

    def open_merge_dialog(self):
        dialog = MergeDialog(self)
        dialog.exec_()

    def open_question_dialog(self):
        dialog = QuestionDialog(self)
        dialog.exec_()

    def open_answer_dialog(self):
        dialog = AnswerDialog(self)
        dialog.exec_()

    def open_csv_dialog(self):
        dialog = CSVDialog(self)
        dialog.exec_()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
