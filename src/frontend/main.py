import sys
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QFrame, QFileDialog, QMessageBox
)
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
import matplotlib.pyplot as plt
import numpy as np

from src.frontend.caselist import get_case_list_widget
from src.frontend.importBrowseDownloadData import ImportBrowseDownloadDataWindow
from src.frontend.importCaseData import ImportCaseDataWindow 
from src.frontend.searchbar import SearchBar
from src.frontend.utils import *
from src.frontend.wrongCaseListWidget import WrongCaseListWindow
from src.frontend.importExclusiveAndBatch import ImportExclusiveAndBatchWindow
from src.frontend.overlayWidget import OverlayWidget
from src.backend.read_case import *
from src.db.init_db import init_db

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
        
        self.wrong_case_list_widget = None

    def initUI(self):
        self.setWindowTitle(" ")
        self.setWindowIcon(QIcon("img/tsinghua.ico"))
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
        button_layout.setSpacing(20)
        
        button_container = QWidget()
        button_base_layout = QVBoxLayout(button_container)
        button_base_layout.addWidget(buttonlist_container)
        
        list_container = QWidget()
        list_container.setObjectName("list_container")
        list_container.setStyleSheet('''
            #list_container {
                background-color: white;
                border-right: 1px solid #dbdbdb;
            }
        ''')
        list_layout = QVBoxLayout(list_container)       # 案例列表
        list_layout.setContentsMargins(10, 10, 10, 10)
        
        
        chart_container = QWidget()
        chart_container.setStyleSheet("background-color: white;")
        chart_layout = QVBoxLayout(chart_container)     # 右侧图表
        
        
        # 左侧按钮
        self.import_case_button = QPushButton("导入案例文档列表")
        set_button_style(self.import_case_button)
        self.import_case_button.clicked.connect(lambda: self.open_import_window(type='case'))
        
        self.import_bd_button = QPushButton("导入浏览下载量")
        set_button_style(self.import_bd_button)
        self.import_bd_button.clicked.connect(lambda: self.open_import_window(type='browse_download'))
        
        # 分界线
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        line.setStyleSheet("color: #e0e0e0;")
        
        # self.royalty_button = QPushButton("导入历年版税")
        # set_button_style(self.royalty_button)
        # self.royalty_button.clicked.connect(self.import_royalty)
        
        self.otherSchool_button = QPushButton("导入学校案例批次")
        set_button_style(self.otherSchool_button)
        self.otherSchool_button.clicked.connect(self.import_other_school)
        
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)
        line2.setStyleSheet("color: #e0e0e0;")
        
        self.migrate_button = QPushButton("导入软件数据")
        set_button_style(self.migrate_button)
        self.migrate_button.clicked.connect(self.import_migration_data)
        
        self.export_button = QPushButton("导出数据")
        set_button_style(self.export_button)
        self.export_button.clicked.connect(self.export_data)
        
        button_layout.addWidget(self.import_case_button)
        button_layout.addWidget(self.import_bd_button)
        button_layout.addWidget(line)
        # button_layout.addWidget(self.royalty_button)
        button_layout.addWidget(self.otherSchool_button)
        button_layout.addWidget(line2)
        button_layout.addWidget(self.migrate_button)
        button_layout.addWidget(self.export_button)
        

        # 搜索框
        self.search_bar = SearchBar(placeholder_text="搜索案例", search_callback=self.on_search_clicked)
        top_layout.addWidget(self.search_bar)

        # 案例列表
        self.case_list_model, self.case_list_view = get_case_list_widget()
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
        
        self.overlay = OverlayWidget(self)
        
    def resizeEvent(self, event):
        """调整遮罩层的大小以匹配窗口"""
        self.overlay.setGeometry(self.rect())
        super().resizeEvent(event)  # 确保父类的 resizeEvent 也被调用
        
    def open_import_window(self, type):
        if type == 'case':
            file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xls *.xlsx *.csv)")
            if not file_path:
                print("读取文件失败")
                return
            
            self.overlay.show_loading_animation()
            self.thread: LoadingUIThread = LoadingUIThread(readCaseList, file_path)
            self.thread.data_loaded.connect(self.load_data_finished)
            self.thread.start()
            
        elif type == 'browse_download':
            self.import_window = ImportBrowseDownloadDataWindow()
            self.import_window.show()
    
    def load_data_finished(self, returns):
        wrong_cases = returns[0]
        
        if len(wrong_cases) > 0:
            wrong_cases = cases_dict_to_widget_list(wrong_cases)
            
            if not self.wrong_case_list_widget:
                self.wrong_case_list_widget = WrongCaseListWindow(wrong_cases)
            else:
                self.wrong_case_list_widget.close()     # 关闭之前的窗口
                self.wrong_case_list_widget = WrongCaseListWindow(wrong_cases)
            self.wrong_case_list_widget.show()
            
        self.overlay.setVisible(False)
        self.overlay.timer.stop()
    
    def import_royalty(self):
        # self.royalty_data = load_data(self)
        # if self.royalty_data is None:
        #     return
        pass
        
    def import_other_school(self):    
        self.import_other_school_window = ImportExclusiveAndBatchWindow()
        self.import_other_school_window.show()
    
    def import_migration_data(self):
        # self.migration_data = load_data(self)
        # if self.migration_data is None:
        #     return
        folder_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if not folder_path:
            print("读取文件失败！")
            return
        
        caselist_path = folder_path + "/案例文档列表.xlsx"
        browse_and_download_path = folder_path + "/浏览下载量.xlsx"
        huatu_path = folder_path + "/华图数据.xlsx"
        
        # 检查三个文件是否都存在
        if not os.path.exists(caselist_path) or not os.path.exists(browse_and_download_path) or not os.path.exists(huatu_path):
            # 弹窗提示
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle("文件缺失")
            msg_box.setText("请确保文件夹中包含以下文件：\n1. 案例文档列表.xlsx\n2. 浏览下载量.xlsx\n3. 华图数据.xlsx")
            msg_box.exec_()
            return
        
        # !清空数据库
        # 删除../../PaymentCal.db文件
        db_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", "PaymentCal.db"))
        if os.path.exists(db_path):
            os.remove(db_path)
        # 重新创建数据库
        init_db()
        
        # TODO 加载圆圈
        
        readCaseList(caselist_path)
        readBrowsingAndDownloadRecord_Tsinghua(browse_and_download_path)
        readBrowsingAndDownloadData_HuaTu(huatu_path)
    
    def export_data(self):
        folder_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if not folder_path:
            print("读取文件失败！")
            return
        
        caselist_path = folder_path + "/案例文档列表.xlsx"
        browse_and_download_path = folder_path + "/浏览下载量.xlsx"
        huatu_path = folder_path + "/华图数据.xlsx"
        
        exportCaseList(caselist_path)
        exportBrowsingAndDownloadRecord(browse_and_download_path)
        exportHuaTuData(huatu_path)
        
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
        
    def on_search_clicked(self):
        search_text = self.search_bar.get_text()
        print(f"搜索内容: {search_text}")
        
        matched_strs, similar_cases = zip(*getSimilarCases(search_text))
         
        if len(similar_cases) > 0:
            similar_caseslist = cases_class_to_widget_list(similar_cases)
            self.case_list_model.update_data(similar_caseslist)
            
            for row in range(self.case_list_model.rowCount()):
                _case = self.case_list_model.data(self.case_list_model.index(row, 0), Qt.DisplayRole)
                _case["matched_str"] = f"匹配项：{matched_strs[row]}"
            self.case_list_model.layoutChanged.emit()
        else:
            self.case_list_model.update_data([])

def app():
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    app()
