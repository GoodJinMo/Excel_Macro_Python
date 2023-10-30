#mywindow
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from Excel_split_Modify import exm_main
from Nozzle_Max import NM_main
from Graph_Rename import GR_main
import style

version="1.01.00 v"
class MyWindow(QMainWindow):
    
    def __init__(self):
       super().__init__()
       self.setWindowTitle("Excel_Macro")
       self.setFixedSize(540, 960)
       self.second_window = None 
       
       font = QFont("Arial", 12)
       font.setBold(True) 
       app.setFont(font)
       
       app.setStyleSheet(style.theme)
       
       self.Home_UI()
        
    def Home_UI(self):
        
        
        layout = QVBoxLayout()
        
        btn_1 =  QPushButton("Split")
        btn_1.setFixedSize(270,50)
        btn_1.clicked.connect(self.Open_exm)
        #btn_1.setStyleSheet(style.btn_style())
        
        
        btn_2 =  QPushButton("Nozzle_Max")
        btn_2.setFixedSize(270,50)
        btn_2.clicked.connect(self.Open_NM)
        
        btn_3 =  QPushButton("Graph_Rename")
        btn_3.setFixedSize(270,50)
        btn_3.clicked.connect(self.Open_GR)
        
        layout.addWidget(btn_1 )
        layout.addWidget(btn_2 )
        layout.addWidget(btn_3 )
        layout.setAlignment(Qt.AlignCenter|Qt.AlignCenter)
        layout.setSpacing(50)
        
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)
        jinmun=QLabel("JiMoo "+version,self)
        jinmun.resize(230,50)
        jinmun.move(400,920)
    def Open_exm(self):
         if self.second_window is None:  # 두 번째 창이 열려 있지 않으면
             self.second_window = exm_main()
             self.second_window.show()
             self.second_window.closed.connect(self.second_window_closed)  # 두 번째 창 닫히는 이벤트 연결
    def Open_NM(self):
         if self.second_window is None:  # 두 번째 창이 열려 있지 않으면
             self.second_window = NM_main()
             self.second_window.show()
             self.second_window.closed.connect(self.second_window_closed)  # 두 번째 창 닫히는 이벤트 연결
    def Open_GR(self):
         if self.second_window is None:  # 두 번째 창이 열려 있지 않으면
             self.second_window = GR_main()
             self.second_window.show()
             self.second_window.closed.connect(self.second_window_closed)  # 두 번째 창 닫히는 이벤트 연결

    def second_window_closed(self):
        self.second_window = None  # 두 번째 창이 닫힐 때 self.second_window를 None으로 설정


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    sys.exit(app.exec_())
