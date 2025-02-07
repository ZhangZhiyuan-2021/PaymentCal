from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
    QLabel, QComboBox, QProgressBar, QSizePolicy, QSpacerItem, QFileDialog, QMessageBox,
    QLineEdit
)
from PyQt5.QtGui import QFont, QIcon, QIntValidator
from PyQt5.QtCore import Qt, QTimer

import sys
import os
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))
from src.frontend.caselist import get_case_list_widget
from src.frontend.searchbar import SearchBar
from src.frontend.utils import *
from src.frontend.wrongCaseListWidget import WrongCaseListWindow
from src.backend.read_case import *

class ImportExclusiveAndBatchWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.data_source = "浙江大学管理学院"
        self.is_exclusive = True
        self.batch = None
        
        self.matching_case_dict = {}    # 键是匹配错误的案例名，值是选择的对应的数据库中的案例名
        self.unmatched_casename_to_class_dict_huatu = {}
        self.unmatched_casename_to_download_record_Thu = {}
        self.unmatched_casename_to_browse_record_Thu = {}
        self.current_unmatched_case = ""    # 当前选中的未匹配案例名
        self.unmatched_case_num = 0 
        self.matching_case_num = 0
        
        self.wrong_case_list_widget = None

    def initUI(self):
        self.setWindowTitle(" ")
        self.setWindowIcon(QIcon("img/tsinghua.ico"))
        self.setGeometry(600, 100, 1200, 1000)
        
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

        label1 = QLabel("数据来源")
        label1.setFont(font)
        label1.setSizePolicy(QSizePolicy.Fixed, QSizePolicy.Fixed)  # 设置QLabel为固定大小

        self.source_combo = QComboBox()
        set_combo_style(self.source_combo)
        self.source_combo.setFixedHeight(50)
        self.source_combo.addItems(["浙江大学管理学院", "中国人民大学商学院"])
        self.source_combo.currentIndexChanged.connect(self.on_source_selected)
        self.source_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)  # 自适应大小
        self.source_combo.setFont(font2)
        self.source_combo.view().setFont(font2)
  
        self.batch_input = QLineEdit()
        self.batch_input.setValidator(QIntValidator(0, 9999, self))  # 只允许输入 0-9999
        self.batch_input.setPlaceholderText("请输入第几批")  
        self.batch_input.setFont(font2)
        self.batch_input.setFixedHeight(50)
        self.batch_input.setStyleSheet("""
            QLineEdit {
                border: 2px solid #D1C4E9;
                border-radius: 8px;
                padding: 6px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid #7b56f0;
            }
        """)
        
        data_source_container = QWidget()
        data_source_container.setStyleSheet('margin-bottom: 10px;')
        data_source_layout = QHBoxLayout(data_source_container)
        data_source_layout.addWidget(label1)
        data_source_layout.addWidget(self.source_combo)
        data_source_layout.addWidget(self.batch_input)

        self.load_button = QPushButton("读取数据")
        set_button_style(self.load_button, 40)
        self.load_button.clicked.connect(self.on_load_data_clicked)
        
        top_layout.addWidget(data_source_container, 2)
        top_layout.addWidget(self.load_button, 2)
    

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
        
        self.confirm_button = QPushButton("确认")
        set_button_style(self.confirm_button, 40)
        self.confirm_button.clicked.connect(self.on_confirm_clicked)
        
        self.search_box = QVBoxLayout()
        self.search_box.setSpacing(10)
        self.search_box.addWidget(self.search_input, 2)
        self.search_box.addItem(QSpacerItem(0, 10, QSizePolicy.Minimum, QSizePolicy.Expanding))
        self.search_box.addWidget(self.search_results_list_view, 6)
        self.search_box.addItem(QSpacerItem(0, 10, QSizePolicy.Minimum, QSizePolicy.Expanding))
        self.search_box.addWidget(self.confirm_button, 2)

        match_layout.addLayout(self.unmatch_box, 1)
        match_layout.addLayout(self.search_box, 1)

        # 添加组件到主布局
        layout.addLayout(top_layout, 2)
        # layout.addWidget(self.table, 4)
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

    def on_source_selected(self, index):
        self.data_source = self.source_combo.currentText()
        if '中国人民大学' in self.data_source:
            is_exclusive = False
        elif '浙江大学' in self.data_source:
            is_exclusive = True

    def on_load_data_clicked(self):     
        try:
            self.init_list()
            
            self.batch = self.batch_input.text()
            if self.batch == "":
                QMessageBox.warning(self, "警告", "请输入批次！")
                return
            self.batch = int(self.batch)
            
            file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xls *.xlsx *.csv)")
            if not file_path:
                print("读取文件失败")
                return
                
            missingInformationCases, unmatchedCases = readCaseExclusiveAndBatch(file_path, self.data_source, self.batch)
            
            if len(missingInformationCases) > 0:
                missingInformationCases_Indexes = [f"序号：{i+1}，标题缺失！" for (i, case) in enumerate(missingInformationCases)]
                wrongcasedict = cases_name_to_widget_list(missingInformationCases_Indexes)
            
                if not self.wrong_case_list_widget:
                    self.wrong_case_list_widget = WrongCaseListWindow(wrongcasedict)
                else:
                    self.wrong_case_list_widget.close()  # 关闭旧的窗口，防止重复
                    self.wrong_case_list_widget = WrongCaseListWindow(wrongcasedict)
                self.wrong_case_list_widget.show()
                
            elif len(unmatchedCases) > 0:
                msg_box = QMessageBox()
                msg_box.setIcon(QMessageBox.Warning)
                msg_box.setWindowTitle(" ")
                msg_box.setText("表格中部分案例名称无法匹配，请手动匹配！")
                msg_box.exec_()
                
                for attr in unmatchedCases[0]:
                    if '标题' in attr:
                        title = attr
                        break
                    
                unmatchedCases_Names = [case[title] for case in unmatchedCases]
                unmatchedCaseslist = cases_name_to_widget_list(unmatchedCases_Names)
                
                self.unmatched_case_model.update_data(unmatchedCaseslist)
                
                self.unmatched_case_model.layoutChanged.emit()
                self.unmatched_case_num = len(unmatchedCases_Names)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"{str(e)}")
            return

    def on_confirm_clicked(self):
        """确认按钮点击事件"""
        print(self.matching_case_dict)
        
        self.batch = self.batch_input.text()
        if self.batch == "":
            QMessageBox.warning(self, "警告", "请输入批次！")
            return
        self.batch = int(self.batch)
        
        if len(self.matching_case_dict) == self.unmatched_case_num:
            for key, value in self.matching_case_dict.items():
                updateCase(value, alias=key, batch=self.batch, owner_name=self.data_source, is_exclusive=self.is_exclusive)
                
        self.init_list()
        
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
        self.unmatched_casename_to_class_dict_huatu = {}
        self.unmatched_casename_to_download_record_Thu = {}
        self.unmatched_casename_to_browse_record_Thu = {}
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
