import sys
import pandas as pd
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from openpyxl import load_workbook
from openpyxl.styles import Font,Alignment
from openpyxl.styles import Border, Side
from Show_ex import Show_ex

class  NM_main(QMainWindow):
    closed = pyqtSignal()
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Nozzle_Max")
        self.setAcceptDrops(True)
        self.resize(800, 700)
        
        self.path_lb = QLabel("File_path")
        self.path_input = QLineEdit()

        self.sheet=""
        self.dprow = ""
        self.dp = ""
        
        self.state_name=QLabel("None")
        self.table_widget = QTableWidget()
        
       
        self.save_button = QPushButton("Save and Export")
        self.save_button.clicked.connect(self.save_and_export)
        
        
        hlayout=QHBoxLayout()
        
        self.Max=QPushButton("Max")
        self.Max.clicked.connect(self.btn_max)
        self.Lookup_Max=QPushButton("Lookup")
        self.Lookup_Max.clicked.connect(self.btn_lookup)
        self.Range=QPushButton("Range")
        self.Range.clicked.connect(self.btn_range)
        
        hlayout.addWidget(self.Max)
        hlayout.addWidget(self.Lookup_Max)
        hlayout.addWidget(self.Range)
        
        

        layout = QVBoxLayout()
        layout.addWidget(self.path_lb)
        layout.addWidget(self.path_input)
        layout.addWidget(self.state_name)
        layout.addWidget(self.table_widget)
        layout.addLayout(hlayout)
       
        layout.addWidget(self.save_button)
       

        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)
        
        
        self.MC=None
        self.state=0
        self.start=None
        self.current=[-1,-1,-1,-1]
        self.cn=[None,None,None,None]
    def btn_max(self):
        self.state=1
        self.state_name.setText("Max: "+chr(self.current[0]+65))
    def btn_lookup(self):
        self.state=2
        self.state_name.setText("Lookup: "+chr(self.current[1]+65))
    def btn_range(self):
        self.state=3
        self.state_name.setText("range_s: "+chr(self.current[2]+65))
    
        
    def contextMenuEvent(self, event):
        if self.state !=0 :
            context_menu = QMenu(self)
            merge_action = context_menu.addAction("select")
            action = context_menu.exec_(self.mapToGlobal(event.pos()))
    
            if action == merge_action :
                self.select()
    def select(self):
         self.row = self.table_widget.currentIndex().row()
         self.c = self.table_widget.currentIndex().column()
         self.col = self.table_widget.currentColumn()
         
         col_header = self.table_widget.horizontalHeaderItem(self.col).text()
         state=self.state-1
         print(state,self.current[state])
         for row in range(self.table_widget.rowCount()):
             if self.current[state] != -1:
                 item = self.table_widget.item( row,self.current[state])
                 item.setBackground(QColor(255, 255, 255))
         if self.state == 3:
            if self.current[3] != -1:
                for row in range(self.table_widget.rowCount()):
                    for ra in range(self.current[2],self.current[3]+1):
                     item = self.table_widget.item(row,ra)
                     item.setBackground(QColor(255, 255, 255))  
         self.current[state] = self.c
         self.cn[state]=col_header
        
         if self.state ==1:
             for row in range(self.table_widget.rowCount()):
                 item = self.table_widget.item(row,self.c)
                 item.setBackground(QColor(255, 0, 0,100)) 
             self.state_name.setText("Max: "+chr(self.current[state]+65))
             
         elif self.state ==2:
            for row in range(self.table_widget.rowCount()):
                item = self.table_widget.item(row,self.c)
                item.setBackground(QColor(0, 255, 0,100)) 
            self.state_name.setText("Lookup: "+chr(self.current[state]+65))
            
         elif self.state ==3:
             
             
            self.start=self.c
            for row in range(self.table_widget.rowCount()):
                item = self.table_widget.item(row,self.c)
                item.setBackground(QColor(0, 0, 255,100)) 
            self.state=4
            self.state_name.setText("range_s:e "+chr(self.current[state-1]+65)+" : "+chr(self.current[state]+65))
            
         elif self.state ==4:
           for row in range(self.table_widget.rowCount()):
               for ra in range(self.start+1,self.c+1):
                item = self.table_widget.item(row,ra)
                item.setBackground(QColor(0, 0, 150,100)) 
           self.state =3
           self.state_name.setText("range_s:e "+chr(self.current[state-1]+65)+" : "+chr(self.current[state]+65))
    def openDialog(self):
         dialog = Show_ex(self.path_input.text())
         result = dialog.exec_()

         if result == QDialog.Accepted:
             self.sheet = dialog.sheet
             self.dprow = dialog.col
             self.MC=dialog.radio_button1.isChecked()
             if self.MC:
                 self.dp = dialog.value
             else :
                dp, ok = QInputDialog.getText(self, "Excel file name", "name:",text=dialog.value)
                if ok:
                    self.dp=dp
             self.open_new_window()
         else:
             print('Dialog canceled.')
              
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()
            
    def dropEvent(self, event):
        files = [u.toLocalFile() for u in event.mimeData().urls()]
        for f in files:
            self.path_input.setText(f)
        self.openDialog()
    def open_new_window(self):
        st = 'part'
        nt = 0
        pa = []
        pas = ['part0']
       
        df = pd.read_excel(self.path_input.text(), sheet_name=self.sheet )
       
        if self.MC:
            for i in df[df.columns[int(self.dprow)]]:
                if self.dp == str(i):
                    nt += 1
                    pas.append(st + str(nt))
                pa.append(st + str(nt))
        else:
            for i in df[df.columns[int(self.dprow)]]:
                if self.dp in str(i):
                    nt += 1
                    pas.append(st + str(nt))
                pa.append(st + str(nt)) 
        df['part'] = pa

        dfs = []
        for name in pas:
            fil = df['part'] == name
            i = df.loc[fil].reset_index()
            i = i.drop(['part', 'index'], axis=1)
            dfs.append(i)

        self.dataframes = dfs[1:]
        self.update_table_widget()
    
    def update_table_widget(self):
        dataframe = self.dataframes[0]
        self.table_widget.setRowCount(dataframe.shape[0])
        self.table_widget.setColumnCount(dataframe.shape[1])
        
        cols=[chr(65+i) for i in range(len(dataframe.columns))]
        
        self.table_widget.setHorizontalHeaderLabels(cols)
        
        for i in range(dataframe.shape[0]):
            for j in range(dataframe.shape[1]):
                cell_value = str(dataframe.iloc[i, j])
                self.table_widget.setItem(i, j, QTableWidgetItem(cell_value))

   


    

    def save_and_export(self):
        dialog = Select_window()
        result = dialog.exec_()
        if result == QDialog.Accepted:
            file_name=self.chooseFolder()+"/"+dialog.file_name_input.text()+".xlsx"
            cn=[ord(n)-65 for n in self.cn]
            n_df=pd.DataFrame()
         
            
            for i in range(len(self.dataframes)):
                self.dataframes[i].iloc[:,cn[0]]=self.dataframes[i].iloc[:,cn[0]].fillna(0) ## null값 대체
                numbers_only = [x for x in self.dataframes[i].iloc[:,cn[0]] if isinstance(x, (int, float))] ## 숫자값만 가져오기
                max_value = max(numbers_only) #최댓값 가져오기
                step=self.dataframes[i][self.dataframes[i].iloc[:,cn[0]]==max_value].iloc[:,cn[1]].tolist()[0] #맥스값 행의 열값 가져오기
                n_df = pd.concat([n_df, self.dataframes[i][self.dataframes[i].iloc[:, cn[2]] == step].iloc[:, cn[2] + 1:cn[3]+1]], ignore_index=True)
            
            try :
                n_df.to_excel(file_name, index=False,header=False)
                QMessageBox.about(self,"suc",f'The file has been successfully saved to "{self.folder_path}".')
            except Exception as e:
                QMessageBox.about(self,"fail",f'Error Message:"{e}".')
    def chooseFolder(self):
         options = QFileDialog.Options()
         self.folder_path = QFileDialog.getExistingDirectory(self, "Choose Folder", options=options)
         
         if self.folder_path:
             return self.folder_path
    def closeEvent(self, event):
        self.closed.emit()
class Select_window(QDialog):
    def __init__(self):
        super().__init__()

        self.initUI()

    def initUI(self):
        self.resize(400, 200)
        self.setWindowTitle('Save')
        
        hlayout=QHBoxLayout()
        self.file_name=QLabel("File_Name : ")
        self.file_name_input=QLineEdit()
        
        hlayout.addWidget(self.file_name)
        hlayout.addWidget(self.file_name_input)
        
        
        ok_button = QPushButton('OK', self)
        ok_button.clicked.connect(self.accept)

        cancel_button = QPushButton('Cancel', self)
        cancel_button.clicked.connect(self.reject)
        
        hlayout3 = QHBoxLayout()
        hlayout3.addWidget(ok_button)
        hlayout3.addWidget(cancel_button)
        
        layout=QVBoxLayout()
        layout.addLayout(hlayout)
        layout.addLayout(hlayout3)
        self.setLayout(layout)
    
   

