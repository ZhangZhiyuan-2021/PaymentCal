from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QLabel, QSizePolicy, QSpacerItem
)
from PyQt5.QtGui import QFont, QIcon
from PyQt5.QtCore import Qt

from src.frontend.caselist import get_case_list_widget
from src.frontend.utils import *
from src.backend.read_case import *

class WrongCaseListWidget(QWidget):
    def __init__(self, wrong_cases):
        super().__init__()

        self.wrong_cases = wrong_cases
        self.initUI()

    def initUI(self):
        # self.setWindowTitle(" ")
        # self.setWindowIcon(QIcon("img/tsinghua.ico"))
        # self.setGeometry(600, 100, 1000, 700)
        
        layout = QVBoxLayout()
        layout.setSpacing(20)
        layout.setContentsMargins(40, 40, 40, 40)

        
        font = QFont()
        font.setFamily("黑体")  # 设置字体
        font.setPointSize(12)  # 设置字体大小
        font2 = QFont(font)
        font2.setPointSize(10)
        
        wrong_label = QLabel("表格中部分信息有误，请参照列表检查！")
        wrong_label.setFont(font)
        wrong_label.setAlignment(Qt.AlignCenter)  # 设置文本居中对齐

        self.wrong_case_model, self.wrong_case_list_view = get_case_list_widget(self.wrong_cases)
        self.wrong_case_list_view.clicked.connect(self.on_wrong_case_clicked)

        # 添加组件到主布局
        layout.addWidget(wrong_label)
        layout.addItem(QSpacerItem(0, 20, QSizePolicy.Minimum, QSizePolicy.Fixed))
        layout.addWidget(self.wrong_case_list_view)

        self.setLayout(layout)
        self.setObjectName("ImportWindow")
        self.setStyleSheet('''
            #ImportWindow {
                background-color: #f4f4f4;
            }
        ''')
         
    def on_wrong_case_clicked(self, index):
        # 对应的case背景颜色改变
        case = self.wrong_case_model.data(index, Qt.DisplayRole)
        if not case:
            return
        # search_results_model其他所有的case背景颜色恢复
        for row in range(self.wrong_case_model.rowCount()):
            _case = self.wrong_case_model.data(self.wrong_case_model.index(row, 0), Qt.DisplayRole)
            if _case:
                _case["highlighted"] = False
        case["highlighted"] = True
        self.wrong_case_model.layoutChanged.emit()

class WrongCaseListWindow(WrongCaseListWidget):
    def __init__(self, wrong_cases):
        super().__init__(wrong_cases)

        self.setWindowTitle(" ")
        self.setWindowIcon(QIcon("img/tsinghua.ico"))
        self.setGeometry(600, 100, 1000, 700)
