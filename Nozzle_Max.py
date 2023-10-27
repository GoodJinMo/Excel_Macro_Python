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
        self.Save_row_input = QLineEdit()
        
        

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
         
         if self.state ==1:
             for row in range(self.table_widget.rowCount()):
                 if self.current[0] !=-1:
                     item = self.table_widget.item( row,self.current[0])
                     item.setBackground(QColor(255, 255, 255))
             for row in range(self.table_widget.rowCount()):
                 item = self.table_widget.item(row,self.c)
                 item.setBackground(QColor(255, 0, 0)) 
             self.current[0] = self.c
             self.cn[0]=col_header
             self.state_name.setText("Max: "+chr(self.current[0]+65))
         if self.state ==2:
            for row in range(self.table_widget.rowCount()):
                if self.current[1] !=-1:
                    item = self.table_widget.item( row,self.current[1])
                    item.setBackground(QColor(255, 255, 255))
            for row in range(self.table_widget.rowCount()):
                item = self.table_widget.item(row,self.c)
                item.setBackground(QColor(0, 255, 0)) 
            self.current[1] = self.c 
            self.cn[1]=col_header
            self.state_name.setText("Max: "+chr(self.current[1]+65))
            
         if self.state ==3:
            for row in range(self.table_widget.rowCount()):
                if self.current[2] !=-1:
                    item = self.table_widget.item( row,self.current[2])
                    item.setBackground(QColor(255, 255, 255))
            for row in range(self.table_widget.rowCount()):
                item = self.table_widget.item(row,self.c)
                item.setBackground(QColor(0, 0, 255)) 
            self.current[2] = self.c 
            self.state=4
            self.state_name.setText("range_e")
            self.cn[2]=col_header
            self.state_name.setText("Max: "+chr(self.current[2]+65))
            
         if self.state ==4:
           for row in range(self.table_widget.rowCount()):
               if self.current[3] !=-1:
                   item = self.table_widget.item( row,self.current[3])
                   item.setBackground(QColor(255, 255, 255))
           for row in range(self.table_widget.rowCount()):
                item = self.table_widget.item(row,self.c)
                item.setBackground(QColor(0, 0, 255)) 
           self.current[3] = self.c
           self.cn[3]=col_header
           self.state_name.setText("range_s")
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
             print(self.sheet ,self.dprow ,  self.dp )
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
        file_name, ok = QInputDialog.getText(self, "Excel file name", "name:")
        
        if ok and file_name:       
            cn=self.Save_row_input.text().split(",")
            for i,n in enumerate(cn):
                cn[i]=ord(n)-65
            index_num=int(self.dprow)
            modified_dataframe = self.dataframes[0].copy()
            n_df=pd.DataFrame()
         
            
            for i in range(len(self.dataframes)):
                self.dataframes[i][cn[0]]=self.dataframes[i][cn[0]].fillna(0) ## null값 대체
                numbers_only = [x for x in self.dataframes[i][cn[0]] if isinstance(x, (int, float))] ## 숫자값만 가져오기
                max_value = max(numbers_only) #최댓값 가져오기
                step=self.dataframes[i][self.dataframes[i][cn[0]]==max_value][cn[1]].tolist()[0] #맥스값 행의 열값 가져오기
                n_df=n_df.append(self.dataframes[i][self.dataframes[i][cn[2]]==step].loc[:, cn[2]+1:cn[3]])
            
            try :
                n_df.to_excel(f"{file_name}.xlsx", index=False,header=False)
                QMessageBox.about(self,f'The file has been successfully saved to "{self.folder_path}".')
            except Exception as e:
                QMessageBox.about(self,f'Error Message:"{e}".')
        
        
    def closeEvent(self, event):
        self.closed.emit()
   

