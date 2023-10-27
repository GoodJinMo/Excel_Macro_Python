import sys
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtGui import QColor

class Show_ex(QDialog):
    def __init__(self,file_path):
        super().__init__()
        self.setWindowTitle("Excel_show")
        self.resize(800, 500)

        self.comboBox = QComboBox(self)
        self.comboBox.currentIndexChanged.connect(self.load_selected_sheet)
        self.comboBox.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum) 
        
        hlayout_=QHBoxLayout()
        self.radio_button1 = QRadioButton('match', self)
        self.radio_button1.setChecked(True)

        self.radio_button2 = QRadioButton('contains', self)
        
        hlayout_.addWidget( self.comboBox)
        hlayout_.addWidget(self.radio_button1)
        hlayout_.addWidget(self.radio_button2)
        
        self.table_widget = QTableWidget()
        self.table_widget.verticalHeader().setVisible(True)  # 수직 헤더 표시
        self.table_widget.horizontalHeader().setVisible(True)  # 수평 헤더 표시
        
        ok_button = QPushButton('OK', self)
        ok_button.clicked.connect(self.accept)

        cancel_button = QPushButton('Cancel', self)
        cancel_button.clicked.connect(self.reject)
        
        hlayout = QHBoxLayout()
        hlayout.addWidget(ok_button)
        hlayout.addWidget(cancel_button)
        
        layout = QVBoxLayout()
        layout.addLayout(hlayout_)
        layout.addWidget(self.table_widget)
        layout.addLayout(hlayout)
        self.setLayout(layout)
      
        self.excel_data = None
        self.load_excel_file(file_path)
        
        self.current_item = None  # 현재 선택한 셀의 아이템을 저장하기 위한 변수

    def load_excel_file(self, file_path):
        self.excel_data = pd.ExcelFile(file_path)
        self.sheets = self.excel_data.sheet_names
        self.comboBox.addItems(self.sheets)

    def load_selected_sheet(self):
        selected_sheet = self.comboBox.currentText()
        df = pd.read_excel(self.excel_data, sheet_name=selected_sheet, header=None)  # header=None로 설정하여 헤더 없이 가져옴
        self.sheet = selected_sheet
        self.table_widget.setRowCount(df.shape[0])
        self.table_widget.setColumnCount(df.shape[1])
        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                cell_value = str(df.iloc[i, j])
                item = QTableWidgetItem(cell_value)
                self.table_widget.setItem(i, j, item)

    def contextMenuEvent(self, event):
        context_menu = QMenu(self)
        merge_action = context_menu.addAction("select")
        action = context_menu.exec_(self.mapToGlobal(event.pos()))

        if action == merge_action:
            self.select()

    def select(self):
        self.row = self.table_widget.currentIndex().row()
        self.col = self.table_widget.currentIndex().column()
        if self.current_item is not None:
            self.current_item.setBackground(QColor(255, 255, 255))  # 기본 색상 (흰색)
        
        item = self.table_widget.item(self.row, self.col)
        item.setBackground(QColor(255, 0, 0))  # 배경색을 빨간색으로 변경
        self.current_item = item  # 현재 선택한 셀을 저장
        self.value = item.text()
  