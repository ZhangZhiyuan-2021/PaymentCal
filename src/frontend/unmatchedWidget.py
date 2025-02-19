
'''
半成品，不使用
'''

from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
    QLabel, QComboBox, QProgressBar, QSizePolicy, QSpacerItem, QFileDialog, QMessageBox,
    QLineEdit
)
from PyQt5.QtGui import QFont, QIcon, QIntValidator
from PyQt5.QtCore import Qt, QTimer

from src.frontend.caselist import get_case_list_widget
from src.frontend.searchbar import SearchBar
from src.frontend.utils import *
from src.frontend.wrongCaseListWidget import WrongCaseListWindow
from src.backend.read_case import *

class UnmatchedWidget(QWidget):
    def __init__(self, unmatched_data):
        super().__init__()
        
        self.initUI()
        self.matching_case_dict = {}    # 键是匹配错误的案例名，值是选择的对应的数据库中的案例名
        self.unmatched_casename_to_class_dict = {}
        self.current_unmatched_case = ""    # 当前选中的未匹配案例名 +
        self.unmatched_case_num = 0     # +
        self.matching_case_num = 0  # +
        
        unmatched_cases_widget_list = cases_to_widget_list(unmatched_data)
         
        # 这行是错的，不应该用 getCase，因为unmatched_casename_to_class_dict是为读华图数据准备的
        # self.unmatched_casename_to_class_dict = {case_dict["title"]: getCase(case_dict["title"]) for case_dict in unmatched_cases_widget_list}
        self.unmatched_case_model.update_data(unmatched_cases_widget_list)
        self.unmatched_case_model.layoutChanged.emit()
        self.unmatched_case_num = len(unmatched_cases_widget_list)
    
        
    def initUI(self):
        font = QFont()
        font.setFamily("黑体")  # 设置字体
        font.setPointSize(12)  # 设置字体大小
        font2 = QFont(font)
        font2.setPointSize(10)
        
        unmatch_label = QLabel("下方列表中的案例无法自动匹配。请按照下列步骤操作：\n1. 选择左侧列表中某个案例；\n2. 在右侧搜索框中输入关键词搜索匹配的案例；\n3. 选择右侧列表中的案例；\n4. 所有案例手动选择完成后，点击确认按钮。")
        unmatch_label.setFont(font2)

        # 匹配不上案例列表
        match_layout = QHBoxLayout()
        self.unmatched_case_model, self.unmatched_case_list_view = get_case_list_widget()
        self.unmatched_case_list_view.clicked.connect(self.on_unmatched_case_clicked)
        
        self.unmatch_box = QVBoxLayout()
        self.unmatch_box.setSpacing(10)
        self.unmatch_box.addWidget(unmatch_label, 2)
        self.unmatch_box.addWidget(self.unmatched_case_list_view, 6)

    
        self.search_input = SearchBar(placeholder_text="输入关键词搜索", search_callback=self.on_search_clicked)

        self.search_results_model, self.search_results_list_view = get_case_list_widget()
        self.search_results_list_view.clicked.connect(self.on_search_results_clicked)
        
        self.search_box = QVBoxLayout()
        self.search_box.setSpacing(10)
        self.search_box.addWidget(self.search_input, 2)
        self.search_box.addItem(QSpacerItem(0, 10, QSizePolicy.Minimum, QSizePolicy.Expanding))
        self.search_box.addWidget(self.search_results_list_view, 6)

        match_layout.addLayout(self.unmatch_box, 1)
        match_layout.addLayout(self.search_box, 1)
        
    def on_unmatched_case_clicked(self, index):
        # 对应的case背景颜色改变
        case = self.unmatched_case_model.data(index, Qt.DisplayRole)
        if not case:
            return
        # unmatched_case_model其他所有的case背景颜色恢复
        for row in range(self.unmatched_case_model.rowCount()):
            _case = self.unmatched_case_model.data(self.unmatched_case_model.index(row, 0), Qt.DisplayRole)
            if _case:
                _case["highlighted"] = False
                
        # 如果当前case已经匹配，则高亮search_results_model中对应的case；如果未匹配，则不高亮
        # 先恢复search_results_model中所有case的背景颜色
        for row in range(self.search_results_model.rowCount()):
            _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
            if _case:
                _case["highlighted"] = False
        # 如果当前case已经匹配，则高亮search_results_model中对应的case
        if case["title"] in self.matching_case_dict:
            print(case["title"], self.matching_case_dict)
            for row in range(self.search_results_model.rowCount()):
                _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
                if _case and _case["title"] == self.matching_case_dict[case["title"]]:
                    _case["highlighted"] = True                     
        
        case["highlighted"] = True
        self.unmatched_case_model.layoutChanged.emit()
        self.search_results_model.layoutChanged.emit()
        self.current_unmatched_case = case["title"]
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
        
        if self.current_unmatched_case != "":
            self.matching_case_dict[self.current_unmatched_case] = case["title"]
            self.matching_case_num = len(self.matching_case_dict)
                
    def on_search_clicked(self):
        """搜索按钮点击事件"""
        keyword = self.search_input.get_text()
        print(f"搜索关键词: {keyword}")
        matched_strs, similar_cases = zip(*getSimilarCases(keyword))
        
        if len(similar_cases) > 0:
            similar_caseslist = cases_class_to_widget_list(similar_cases)
            self.search_results_model.update_data(similar_caseslist)
            
            # 如果当前case已经匹配，则高亮search_results_model中对应的case；如果未匹配，则不高亮
            # 先恢复search_results_model中所有case的背景颜色
            for row in range(self.search_results_model.rowCount()):
                _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
                if _case:
                    _case["highlighted"] = False
                    
                _case['matched_str'] = f"匹配项：{matched_strs[row]}"
            # 如果当前case已经匹配，则高亮search_results_model中对应的case
            if self.current_unmatched_case in self.matching_case_dict:
                for row in range(self.search_results_model.rowCount()):
                    _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
                    if _case and _case["title"] == self.matching_case_dict[self.current_unmatched_case]:
                        _case["highlighted"] = True         
            
            self.search_results_model.layoutChanged.emit()
        else:
            print("未找到相关案例！")
            self.search_results_model.update_data([])
            
    def init_list(self):
        self.matching_case_dict = {}
        self.unmatched_casename_to_class_dict = {}
        self.unmatched_case_model.update_data([])
        self.unmatched_case_model.layoutChanged.emit()
        self.current_unmatched_case = ""
        self.unmatched_case_num = 0
        self.matching_case_num = 0
        self.init_search_area()
        
    def init_search_area(self):
        self.search_input.clear_text()
        self.search_results_model.update_data([])
        self.search_results_model.layoutChanged.emit()
                
    def all_matched(self):
        return self.unmatched_case_num == self.matching_case_num
    
    def update_browseAndDownload_database(self, data_source):
        if not self.all_matched():
            print("未完成所有案例的匹配！")
            return
        
        for alias, case_name in self.matching_case_dict.items():
            updateCase(case_name, alias)
            
        if data_source == "中国工商案例库":
            # TODO 添加浏览记录和下载记录到数据库
            pass
        elif data_source == "华图案例库":
            for key, value in self.unmatched_casename_to_class_dict.items():
                addBrowsingAndDownloadData_HuaTu(self.matching_case_dict[key], self.huatu_year, value['查看数'], value['邮件数'])