from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
    QLabel, QComboBox, QProgressBar, QSizePolicy, QSpacerItem, QFileDialog, QMessageBox
)
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt, QTimer

import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from src.frontend.caselist import get_case_list_widget
from src.frontend.searchbar import SearchBar
from src.frontend.utils import *
from src.backend.read_case import *

class ImportCaseDataWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.exist_unmatched = False
        self.data_source = "中国工商案例库"
        self.copyright_source = "清华"

    def initUI(self):
        self.setWindowTitle(" ")
        self.setWindowIcon(QIcon("img/tsinghua.ico"))
        self.setGeometry(600, 100, 1000, 700)
        
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(40, 40, 40, 40)

        # 顶部选择框和按钮
        top_layout = QHBoxLayout()
        top_layout.setSpacing(40)
        
        font = QFont()
        font.setFamily("黑体")  # 设置字体
        font.setPointSize(12)  # 设置字体大小
        font2 = QFont(font)
        font2.setPointSize(10)

        self.load_button = QPushButton("读取数据")
        set_button_style(self.load_button, 40)
        self.load_button.clicked.connect(self.on_load_data_clicked)
        
        top_layout.addWidget(self.load_button, 2)
        
        test_cases = [
            {"title": "案例A", "info": "2017年入库... 其他信息"},
            {"title": "案例B", "info": "2018年入库... 其他信息"},
            {"title": "案例C", "info": "2019年入库... 其他信息"},
            {"title": "案例D", "info": "2020年入库... 其他信息"}
        ]

        # 匹配不上案例列表
        match_layout = QHBoxLayout()
        self.unmatched_case_model, self.unmatched_case_list_view = get_case_list_widget()
        self.unmatched_case_list_view.clicked.connect(self.on_unmatched_case_clicked)
        # self.unmatched_case_list_view.setStyleSheet("""
        #     border: none;
        # """)
    

        match_layout.addWidget(self.unmatched_case_list_view)

        # 添加组件到主布局
        layout.addLayout(top_layout, 2)
        layout.addItem(QSpacerItem(0, 20, QSizePolicy.Minimum, QSizePolicy.Fixed))
        layout.addLayout(match_layout, 6)

        self.setLayout(layout)
        self.setObjectName("ImportWindow")
        self.setStyleSheet('''
            #ImportWindow {
                background-color: #f4f4f4;
            }
        ''')
         
    def on_unmatched_case_clicked(self, index):
        # 对应的case背景颜色改变
        case = self.unmatched_case_model.data(index, Qt.DisplayRole)
        if not case:
            return
        # search_results_model其他所有的case背景颜色恢复
        for row in range(self.unmatched_case_model.rowCount()):
            _case = self.unmatched_case_model.data(self.unmatched_case_model.index(row, 0), Qt.DisplayRole)
            if _case:
                _case["highlighted"] = False
        case["highlighted"] = True
        self.unmatched_case_model.layoutChanged.emit()

    def on_load_data_clicked(self):  
        self.unmatched_case_model.update_data([])
        
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xls *.xlsx *.csv)")
        if not file_path:
            print("读取文件失败")
            return

        wrong_cases = readCaseList(file_path)
        if len(wrong_cases) > 0:
            self.exist_unmatched = True
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle(" ")
            msg_box.setText("表格中部分信息有误，请参照列表检查！")
            msg_box.exec_()
            # 展示列表，列出错误信息
            wrong_cases = cases_dict_to_widget_list(wrong_cases)
            self.unmatched_case_model.update_data(wrong_cases)
        else:
            self.exist_unmatched = False
                
        # self.exist_unmatched = True
