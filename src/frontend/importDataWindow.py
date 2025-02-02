from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
    QLabel, QComboBox, QProgressBar, QListWidget, QLineEdit, QCheckBox, QListView,
    QSizePolicy, QFileDialog
)
from PyQt5.QtCore import Qt, QTimer
import pandas as pd

from caselist import CaseListModel, CaseItemDelegate, get_case_list_widget
from utils import load_data

class ImportDataWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("导入数据")
        self.setGeometry(600, 100, 800, 1200)

        self.initUI()
        self.exist_unmatched = False
        self.data_source = "中国工商案例库"
        self.copyright_source = "清华"

    def initUI(self):
        layout = QVBoxLayout()

        # 顶部选择框和按钮
        top_layout = QHBoxLayout()
        self.source_combo = QComboBox()
        self.source_combo.addItems(["中国工商案例库", "华图"])
        self.source_combo.currentIndexChanged.connect(self.on_source_selected)
        self.copyright_combo = QComboBox()
        self.copyright_combo.addItems(["清华", "浙大", "人大"])
        self.copyright_combo.currentIndexChanged.connect(self.on_copyright_selected)

        self.load_button = QPushButton("读取数据")
        self.load_button.clicked.connect(self.on_load_data_clicked)

        top_layout.addWidget(QLabel("选择数据来源"))
        top_layout.addWidget(self.source_combo)
        top_layout.addWidget(QLabel("选择版权方"))
        top_layout.addWidget(self.copyright_combo)
        top_layout.addWidget(self.load_button)

        # 数据表格
        self.table = QTableWidget()
        self.table.setColumnCount(3)  # 假设有3列
        self.table.setHorizontalHeaderLabels(["列1", "列2", "列3"])

        # 计算 & 进度条
        calc_layout = QVBoxLayout()
        self.calc_button = QPushButton("开始计算")
        self.calc_button.clicked.connect(self.on_calc_clicked)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        calc_layout.addWidget(self.calc_button)
        calc_layout.addWidget(self.progress_bar)
        
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
        
        self.search_box = QVBoxLayout()
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入关键词搜索")

        self.search_results_model, self.search_results_list_view = get_case_list_widget(test_cases)
        self.search_results_list_view.clicked.connect(self.on_search_results_clicked)
        
        self.confirm_button = QPushButton("确认")
        self.confirm_button.clicked.connect(self.on_confirm_clicked)
        
        self.search_box.addWidget(self.search_input)
        self.search_box.addWidget(self.search_results_list_view)
        self.search_box.addWidget(self.confirm_button)

        match_layout.addWidget(self.unmatched_case_list_view)
        match_layout.addLayout(self.search_box)

        # 添加组件到主布局
        layout.addLayout(top_layout, 2)
        layout.addWidget(self.table, 4)
        layout.addLayout(calc_layout, 2)
        layout.addLayout(match_layout, 6)

        self.setLayout(layout)
        
    def on_unmatched_case_clicked(self, index):
        pass
    
    def on_search_results_clicked(self, index):
        # 对应的case背景颜色改变
        case = self.search_results_model.data(index, Qt.DisplayRole)
        if not case:
            return
        # search_results_model其他所有的case背景颜色恢复
        for row in range(self.search_results_model.rowCount()):
            _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
            if _case:
                _case["highlighted"] = False
        case["highlighted"] = True
        self.search_results_model.layoutChanged.emit()

    def on_source_selected(self, index):
        """当选择数据来源时触发"""
        self.data_source = self.source_combo.currentText()
        print(f"选择了数据来源: {self.data_source}")

    def on_copyright_selected(self, index):
        """当选择版权方时触发"""
        self.copyright_source = self.copyright_combo.currentText()
        print(f"选择了版权方: {self.copyright_source}")

    def on_load_data_clicked(self):
        data = load_data(self)
        self.case_data = data
            
        if isinstance(data, dict):  
            sheet_names = list(data.keys())  
            print(f"文件包含多个 Sheet: {sheet_names}")
            data = data[sheet_names[0]]  # 读取第一个 sheet
            
        data = data.head(10)
            
        self.table.setRowCount(len(data))
        self.table.setColumnCount(len(data.columns))
        self.table.setHorizontalHeaderLabels(data.columns)

        for i, row in data.iterrows():
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value))
                item.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)  # 只读
                self.table.setItem(i, j, item)

        # self.exist_unmatched = True

    def on_calc_clicked(self):
        """模拟计算过程，并更新进度条"""
        self.progress_bar.setValue(0)
        self.calc_button.setEnabled(False)  # 禁用按钮，防止多次点击

        def update_progress():
            current_value = self.progress_bar.value()
            if current_value < 100:
                self.progress_bar.setValue(current_value + 10)
                QTimer.singleShot(300, update_progress)  # 300ms 后递归调用
            else:
                self.calc_button.setEnabled(True)
                print("计算完成！")

        QTimer.singleShot(300, update_progress)  # 开始更新进度条

    def on_confirm_clicked(self):
        """确认按钮点击事件"""
        selected_cases = []
        for row in range(self.search_results_model.rowCount()):
            case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
            if case and case.get("highlighted", False):
                selected_cases.append(case["title"])

        print(f"已确认的案例: {selected_cases}")
