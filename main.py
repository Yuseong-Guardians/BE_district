import sys
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, 
                             QFileDialog, QLabel, QListWidget, QListWidgetItem)
from PyQt5.QtCore import Qt
import pandas as pd
import os

class ExcelMerger(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        layout = QVBoxLayout()

        # 드래그 앤 드롭 안내 메시지
        self.label = QLabel('파일을 드래그하거나, 아래 버튼으로 파일을 추가하세요.')
        layout.addWidget(self.label)

        # 파일 리스트를 표시할 QListWidget
        self.file_list = QListWidget()
        layout.addWidget(self.file_list)

        # 파일 선택 버튼
        self.btn_select = QPushButton('엑셀 파일 선택', self)
        self.btn_select.clicked.connect(self.select_files)
        layout.addWidget(self.btn_select)

        # 병합 버튼
        self.btn_merge = QPushButton('병합하기', self)
        self.btn_merge.clicked.connect(self.merge_files)
        layout.addWidget(self.btn_merge)

        # 드래그 앤 드롭을 위한 설정
        self.setAcceptDrops(True)

        self.setLayout(layout)
        self.setWindowTitle('엑셀 병합기')
        self.setGeometry(300, 300, 400, 300)

        self.files = []

    # 드래그된 파일이 들어왔을 때 호출되는 메서드
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    # 드롭된 파일을 처리하는 메서드
    def dropEvent(self, event):
        urls = event.mimeData().urls()
        for url in urls:
            file_path = url.toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):
                self.add_file(file_path)

    # 파일을 리스트에 추가하는 메서드
    def add_file(self, file_path):
        if file_path not in self.files:
            self.files.append(file_path)
            file_name = os.path.basename(file_path)
            item = QListWidgetItem(file_name)
            item.setToolTip(file_path)
            self.file_list.addItem(item)
        else:
            print(f"{file_path}는 이미 추가되었습니다.")

    # 파일 선택을 통해 엑셀 파일을 추가
    def select_files(self):
        selected_files, _ = QFileDialog.getOpenFileNames(self, '엑셀 파일 선택', '', '엑셀 파일 (*.xlsx *.xls)')
        for file in selected_files:
            self.add_file(file)

    # 엑셀 파일 병합
    def merge_files(self):
        if not self.files:
            print("파일을 먼저 선택하세요.")
            return

        # 병합할 시트 이름 목록
        sheet_names = ['현황', '신규자', '중지자']
        
        # 각각의 시트 데이터를 저장할 딕셔너리 초기화
        merged_sheets = {sheet_name: [] for sheet_name in sheet_names}

        # 각 파일의 시트를 읽어 병합
        for file in self.files:
            excel_data = pd.read_excel(file, sheet_name=None)  # 파일에서 모든 시트 읽기
            for sheet_name in sheet_names:
                if sheet_name in excel_data:
                    merged_sheets[sheet_name].append(excel_data[sheet_name])
                else:
                    print(f"파일 {file}에 {sheet_name} 시트가 없습니다.")

        # 시트별로 데이터를 병합
        for sheet_name in sheet_names:
            if merged_sheets[sheet_name]:
                merged_sheets[sheet_name] = pd.concat(merged_sheets[sheet_name], ignore_index=True)
            else:
                print(f"{sheet_name} 시트에 병합할 데이터가 없습니다.")

        # 병합된 데이터를 새로운 엑셀 파일로 저장
        save_path, _ = QFileDialog.getSaveFileName(self, '저장할 파일 이름', '', '엑셀 파일 (*.xlsx)')
        if save_path:
            with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
                for sheet_name, df in merged_sheets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
            print(f"파일이 저장되었습니다: {save_path}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelMerger()
    ex.show()
    sys.exit(app.exec_())
