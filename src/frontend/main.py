import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QLineEdit, QListWidget, QLabel, QFrame, QGraphicsDropShadowEffect, 
    QListWidgetItem, QStyledItemDelegate, QListView, QSizePolicy
)
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt, QSize, QAbstractListModel
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
import numpy as np

from caselist import CaseListModel, CaseItemDelegate, get_case_list_widget
from importDataWindow import ImportDataWindow
from utils import load_data, set_button_style

class MatplotlibCanvas(FigureCanvas):
    def __init__(self, parent=None):
        plt.rcParams['font.sans-serif'] = ['SimHei']  # 使用黑体支持中文
        plt.rcParams['axes.unicode_minus'] = False  # 解决负号显示问题
        
        self.fig, self.ax = plt.subplots()
        super().__init__(self.fig)
        self.plot_example()

    def plot_example(self):
        years = np.array([2017, 2018, 2019, 2020, 2021])
        values = np.array([140, 160, 155, 158, 165])
        self.ax.plot(years, values, marker='', linewidth=2, linestyle='-', color='#9b69ff', alpha=0.6)
        self.ax.set_title("案例A稿酬历史发放记录")
        self.ax.set_xlabel("年份", fontsize=18)
        self.ax.set_ylabel("金额", fontsize=18)
        self.ax.set_xticks(years)  # 设置横轴刻度
        self.ax.grid(True)
        
        # 在每个点的上方显示 y 轴的具体值
        for i, value in enumerate(values):
            self.ax.annotate(f'{value}', xy=(years[i], values[i]), xytext=(0, 5),
                             textcoords='offset points', ha='center', fontsize=12)

        self.fig.subplots_adjust(left=0.1, right=0.9, top=0.9, bottom=0.15)
        self.draw()

class MainWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle("稿酬计算小程序")
        self.setGeometry(400, 400, 1800, 600)
        # self.init_shadow_effect()

        base_layout = QHBoxLayout()
        case_layout = QVBoxLayout()
        top_layout = QHBoxLayout()      # 顶部栏
        right_layout = QHBoxLayout()     # 案例信息
        case_layout.setSpacing(0)       # 设置布局之间的间距为0
        case_layout.setContentsMargins(0, 0, 0, 0)  # 设置布局的外边距为0
        top_layout.setSpacing(0)       # 设置布局之间的间距为0
        top_layout.setContentsMargins(0, 0, 0, 0)  # 设置布局的外边距为0
        right_layout.setSpacing(0)       # 设置布局之间的间距为0
        right_layout.setContentsMargins(0, 0, 0, 0)  # 设置布局的外边距为0
        
        
        buttonlist_container = QWidget()
        buttonlist_container.setFixedHeight(500)
        button_layout = QVBoxLayout(buttonlist_container)   # 左侧按钮
        button_layout.setContentsMargins(30, 50, 30, 50)  # 设置边距
        
        button_container = QWidget()
        button_base_layout = QVBoxLayout(button_container)
        button_base_layout.addWidget(buttonlist_container)
        
        # 创建阴影效果
        # button_container.setAttribute(Qt.WA_TranslucentBackground)
        # button_container.setGraphicsEffect(self.shadow_effect)
        
        list_container = QWidget()
        list_container.setStyleSheet("background-color: white;")
        list_layout = QVBoxLayout(list_container)       # 案例列表
        list_layout.setContentsMargins(20, 20, 20, 20)
  
        
        chart_container = QWidget()
        chart_container.setStyleSheet("background-color: white;")
        chart_layout = QVBoxLayout(chart_container)     # 右侧图表
        
        
        # 左侧按钮
        self.import_button = QPushButton("导入数据")
        set_button_style(self.import_button)
        self.import_button.clicked.connect(self.open_import_window)
        
        # 分界线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #e0e0e0;")
        
        self.royalty_button = QPushButton("导入历年版税")
        set_button_style(self.royalty_button)
        self.royalty_button.clicked.connect(self.import_royalty)
        
        self.migrate_button = QPushButton("导入软件数据")
        set_button_style(self.migrate_button)
        self.migrate_button.clicked.connect(self.import_migration_data)
        
        self.export_button = QPushButton("导出数据")
        set_button_style(self.export_button)
        self.export_button.clicked.connect(self.export_data)
        
        button_layout.addWidget(self.import_button)
        button_layout.addWidget(line)
        button_layout.addWidget(self.royalty_button)
        button_layout.addWidget(self.migrate_button)
        button_layout.addWidget(self.export_button)
        

        # 搜索框
        self.search_bar = QLineEdit()
        self.search_bar.setPlaceholderText("搜索案例")
        self.search_bar.setFixedHeight(60)
        self.search_bar.setStyleSheet("""
            font-size: 24px;
            border: none;
            border-bottom: 2px solid rgb(200, 200, 200);
            border-radius: 8px;
            padding: 10px;
        """)
        top_layout.addWidget(self.search_bar)

        # 案例列表
        cases = [
            {"title": "案例A", "info": "2017年入库... 其他信息"},
            {"title": "案例B", "info": "2018年入库... 其他信息"},
            {"title": "案例C", "info": "2019年入库... 其他信息"},
            {"title": "案例D", "info": "2020年入库... 其他信息"}
        ]
        # self.case_list_view.model().layoutChanged.emit()
        # self.case_list_view.update()
        self.case_list_model, self.case_list_view = get_case_list_widget(cases)
        self.case_list_view.setStyleSheet("""
            border: none;
            border-right: 1px solid #dbdbdb;
            padding-right: 20px;
        """)
        self.case_list_view.clicked.connect(self.on_case_clicked)
        
        list_layout.addWidget(self.case_list_view, 2)

        # Matplotlib 图表
        self.canvas = MatplotlibCanvas()
        chart_layout.setContentsMargins(30, 50, 30, 50)
        chart_layout.addWidget(self.canvas, 4)


        right_layout.addWidget(list_container, 1)
        right_layout.addWidget(chart_container, 1)
        right_layout.setContentsMargins(0, 20, 0, 0)
      
        case_layout.addLayout(top_layout)
        case_layout.addLayout(right_layout)
        
        base_layout.addWidget(button_container, 2)
        base_layout.addLayout(case_layout, 8)

        self.setLayout(base_layout)
        self.setObjectName("MainWindow")
        self.setStyleSheet('''
            #MainWindow {
                background-color: #f6f6f6;
            }
        ''')
        
    # def init_shadow_effect(self):
    #     self.shadow_effect = QGraphicsDropShadowEffect()
    #     self.shadow_effect.setBlurRadius(30)
    #     self.shadow_effect.setOffset(0, 0)
    #     self.shadow_effect.setColor(Qt.gray)
        
    def open_import_window(self):
        self.import_window = ImportDataWindow()
        self.import_window.show()
    
    def import_royalty(self):
        self.royalty_data = load_data(self)
        if self.royalty_data is None:
            return
    
    def import_migration_data(self):
        self.migration_data = load_data(self)
        if self.migration_data is None:
            return
    
    def export_data(self):
        pass
        
    def on_case_clicked(self, index):
        """点击案例项时的回调函数"""
        case = self.case_list_model.data(index, Qt.DisplayRole)
        if not case:
            return
        # case_list_model其他所有的case背景颜色恢复
        for row in range(self.case_list_model.rowCount()):
            _case = self.case_list_model.data(self.case_list_model.index(row, 0), Qt.DisplayRole)
            if _case:
                _case["highlighted"] = False
        case["highlighted"] = True
        self.case_list_model.layoutChanged.emit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
