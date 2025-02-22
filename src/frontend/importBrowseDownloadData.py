from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
    QLabel, QComboBox, QProgressBar, QSizePolicy, QSpacerItem, QFileDialog, QMessageBox,
    QLineEdit, QCheckBox
)
from PyQt5.QtGui import QFont, QIcon, QIntValidator, QPainter, QPen, QColor, QDoubleValidator
from PyQt5.QtCore import Qt, QTimer, QEvent

from src.frontend.caselist import get_case_list_widget
from src.frontend.searchbar import SearchBar
from src.frontend.utils import *
from src.frontend.wrongCaseListWidget import WrongCaseListWindow
from src.frontend.overlayWidget import OverlayWidget
from src.frontend.progressBar import ProgressBar
from src.backend.read_case import *

class ImportBrowseDownloadDataWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()
        self.data_source = "中国工商案例库"
        self.huatu_year = None
        
        self.matching_case_dict = {}    # 键是匹配错误的案例名，值是选择的对应的数据库中的案例名
        self.unmatched_casename_to_class_dict_huatu = {}
        self.unmatched_casename_to_download_record_Thu = {}
        self.unmatched_casename_to_browse_record_Thu = {}
        self.current_unmatched_case = ""    # 当前选中的未匹配案例名
        self.unmatched_case_num = 0 
        self.matching_case_num = 0
        
        self.wrong_case_list_widget = None
    
        # 启用事件过滤器
        self.installEventFilter(self)

    def eventFilter(self, obj, event):
        # 如果事件是按键事件并且按的是空格键
        if event.type() == QEvent.KeyPress and event.key() == Qt.Key_Space:
            return True  # 忽略空格键事件
        return super().eventFilter(obj, event)  # 否则正常处理事件

    def initUI(self):
        self.setWindowTitle(" ")
        self.setWindowIcon(QIcon("img/tsinghua.ico"))
        self.setGeometry(600, 100, 1200, 1000)
        
        layout = QVBoxLayout()
        layout.setSpacing(10)
        layout.setContentsMargins(40, 40, 40, 40)

        # 顶部选择框和按钮
        top_layout = QHBoxLayout()
        top_layout.setSpacing(40)
        top_layout.setContentsMargins(0, 0, 0, 0)
        
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
        self.source_combo.addItems(["中国工商案例库", "华图"])
        self.source_combo.currentIndexChanged.connect(self.on_source_selected)
        self.source_combo.setSizeAdjustPolicy(QComboBox.AdjustToContents)  # 自适应大小
        self.source_combo.setFont(font2)
        self.source_combo.view().setFont(font2)
  
        self.year_input = QLineEdit()
        self.year_input.setValidator(QIntValidator(0, 9999, self))  # 只允许输入 0-9999
        self.year_input.setPlaceholderText("请输入年份")  
        self.year_input.setFont(font2)
        self.year_input.setFixedHeight(50)
        self.year_input.setStyleSheet("""
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
        self.year_input.setVisible(False)  # 默认隐藏
        
        data_source_container = QWidget()
        data_source_container.setStyleSheet('margin-bottom: 10px;')
        # data_source_container.setStyleSheet('border: 1px solid #D1C4E9; border-radius: 8px; padding: 10px;')
        data_source_layout = QHBoxLayout(data_source_container)
        data_source_layout.addWidget(label1)
        data_source_layout.addWidget(self.source_combo)
        data_source_layout.addWidget(self.year_input)

        self.load_button = QPushButton("读取数据")
        set_button_style(self.load_button, 40)
        self.load_button.clicked.connect(self.on_load_data_clicked)
        
        top_layout.addWidget(data_source_container, 2)
        top_layout.addWidget(self.load_button, 2)
        
        # 读取清华浏览下载记录的进度条
        self.load_progress_bar = ProgressBar()


        unmatch_label = QLabel("下方列表中的案例无法自动匹配。请按照下列步骤操作：\n1. 选择左侧列表中某个案例；\n2. 在右侧搜索框中输入关键词搜索匹配的案例；\n3. 选择右侧列表中的案例；\n4. 所有案例手动选择完成后，点击确认按钮。")
        unmatch_label.setFont(font2)

        # 匹配不上案例列表
        match_layout = QHBoxLayout()
        match_layout.setSpacing(20)
        self.unmatched_case_model, self.unmatched_case_list_view = get_case_list_widget()
        self.unmatched_case_list_view.clicked.connect(self.on_unmatched_case_clicked)
        # TODO 允许不匹配
        
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
        
        
        # 计算 & 进度条
        calc_layout = QVBoxLayout()
        calc_layout.setSpacing(10)
        calc_input_layout = QHBoxLayout()
        calc_input_layout.setSpacing(20)
        calc_button_layout = QHBoxLayout() 
        calc_button_layout.setSpacing(20)
        
        self.calc_year_input = QLineEdit()
        self.calc_year_input.setValidator(QIntValidator(0, 9999, self))  # 只允许输入 0-9999
        self.calc_year_input.setPlaceholderText("结算年份,默认为今年,结算去年稿酬")
        self.calc_year_input.setStyleSheet("""
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
        
        self.decimal_input = QLineEdit()
        validator = QDoubleValidator(0, 9999, 2)  # 允许输入最多2位小数的数字
        self.decimal_input.setValidator(validator)
        self.decimal_input.setStyleSheet("""
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
        self.decimal_input.setPlaceholderText("浏览量折扣因子,默认为0.3")
        # self.decimal_input.setFixedHeight(40)
        
        self.total_money_input = QLineEdit()
        self.total_money_input.setValidator(QIntValidator(0, 999999999, self))
        self.total_money_input.setPlaceholderText("总金额,单位为元")
        self.total_money_input.setStyleSheet("""
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
        
        self.square_root_checkbox = QCheckBox("浏览量开根号")
        # self.square_root_checkbox.setFixedHeight(40)
        self.square_root_checkbox.setStyleSheet("""
            QCheckBox {
                font-size: 20px;
                padding-left: 10px;
            }
            QCheckBox::indicator {
                width: 32px;
                height: 32px;
                border: 2px solid #D1C4E9;
                border-radius: 5px;
            }
            QCheckBox::indicator:checked {
                background-color: #7b56f0;
                border-radius: 5px;
                border: 2px solid #7b56f0;
            }
            QCheckBox::indicator:checked::before {
                font-size: 16px;
                margin-left: 2px;
                margin-top: 2px;
            }
        """)
        # self.square_root_checkbox.setFont(font)
        
        self.calc_button = QPushButton("开始计算")
        set_button_style(self.calc_button, 40)
        self.calc_button.clicked.connect(self.on_calc_clicked)
        self.calc_button.setFocusPolicy(Qt.NoFocus)
        
        self.export_button = QPushButton("导出版税")
        set_button_style(self.export_button, 40)
        self.export_button.clicked.connect(self.on_export_clicked)
        
        self.calc_progress_bar = ProgressBar()

        # calc_layout.addWidget(self.calc_button)
        calc_input_layout.addWidget(self.calc_year_input, 3)
        calc_input_layout.addWidget(self.decimal_input, 3)
        calc_input_layout.addWidget(self.total_money_input, 3)
        calc_input_layout.addWidget(self.square_root_checkbox, 2)
        
        calc_button_layout.addWidget(self.calc_button)
        calc_button_layout.addWidget(self.export_button)
        
        calc_layout.addLayout(calc_input_layout)
        calc_layout.addItem(QSpacerItem(0, 10, QSizePolicy.Minimum, QSizePolicy.Fixed))
        calc_layout.addLayout(calc_button_layout)
        calc_layout.addWidget(self.calc_progress_bar)
        

        # 添加组件到主布局
        layout.addLayout(top_layout, 2)
        layout.addWidget(self.load_progress_bar, 1)
        layout.addItem(QSpacerItem(0, 20, QSizePolicy.Minimum, QSizePolicy.Fixed))
        layout.addLayout(match_layout, 6)
        layout.addItem(QSpacerItem(0, 20, QSizePolicy.Minimum, QSizePolicy.Fixed))
        layout.addLayout(calc_layout, 2)

        self.setLayout(layout)
        self.setObjectName("ImportWindow")
        self.setStyleSheet('''
            #ImportWindow {
                background-color: #f4f4f4;
            }
        ''')
        
        self.overlay = OverlayWidget(self)
        
    def resizeEvent(self, event):
        """调整遮罩层的大小以匹配窗口"""
        self.overlay.setGeometry(self.rect())
        super().resizeEvent(event)  # 确保父类的 resizeEvent 也被调用
        
    def on_search_clicked(self):
        """搜索按钮点击事件"""
        keyword = self.search_input.get_text()
        matched_strs, similar_cases = zip(*getSimilarCases(keyword))
        
        if len(similar_cases) > 0:
            similar_caseslist = cases_class_to_widget_list(similar_cases)
            # 在similar_caseslist头部添加“不匹配”项
            similar_caseslist.insert(0, {"title": "未匹配", "highlighted": False, "matched_str": "", "info": "如果找不到匹配项，请选择此项"})
            self.search_results_model.update_data(similar_caseslist)
            
            # 如果当前case已经匹配，则高亮search_results_model中对应的case；如果未匹配，则不高亮
            # 从1开始，因为第0个是“未匹配”
            for row in range(1, self.search_results_model.rowCount()):
                _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
                if _case:
                    _case["highlighted"] = False
                    
                _case['matched_str'] = f"匹配项：{matched_strs[row-1]}"
            # 如果当前case已经匹配，则高亮search_results_model中对应的case
            if self.current_unmatched_case in self.matching_case_dict:
                for row in range(self.search_results_model.rowCount()):
                    _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
                    # if _case and _case["title"] == self.matching_case_dict[self.current_unmatched_case]:
                    #     _case["highlighted"] = True     
                    # 现在将self.matching_case_dict[self.current_unmatched_case]改为一个list，存储多个匹配项
                    if _case and _case["title"] in self.matching_case_dict[self.current_unmatched_case]:
                        _case["highlighted"] = True    
            
            self.search_results_model.layoutChanged.emit()
        else:
            print("未找到相关案例！")
            self.search_results_model.update_data([])
        
    def on_unmatched_case_clicked(self, index):
        self.unmatched_case_list_view.simulate_right_click(index)
        
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
            for row in range(self.search_results_model.rowCount()):
                _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
                # if _case and _case["title"] == self.matching_case_dict[case["title"]]:
                #     _case["highlighted"] = True        
                # 现在将self.matching_case_dict[self.current_unmatched_case]改为一个list，存储多个匹配项
                if _case and _case["title"] in self.matching_case_dict[case["title"]]:
                    _case["highlighted"] = True             
        
        case["highlighted"] = True
        self.unmatched_case_model.layoutChanged.emit()
        self.search_results_model.layoutChanged.emit()
        self.current_unmatched_case = case["title"]
        
        self.search_input.set_text(case["title"])
        self.on_search_clicked()
    
    def on_search_results_clicked(self, index):
        # 对应的case背景颜色改变
        case = self.search_results_model.data(index, Qt.DisplayRole)
        if not case:
            return
        
        # # 如果self.matching_case_dict[self.current_unmatched_case]有值，且当前index选中了同一个case，则取消匹配
        # if self.current_unmatched_case in self.matching_case_dict and self.matching_case_dict[self.current_unmatched_case] == case["title"]:
        #     del self.matching_case_dict[self.current_unmatched_case]
        #     self.matching_case_num = len(self.matching_case_dict)
        #     case["highlighted"] = False
        #     self.search_results_model.layoutChanged.emit()
        # 现在将self.matching_case_dict[self.current_unmatched_case]改为一个list，存储多个匹配项
        if self.current_unmatched_case in self.matching_case_dict and case["title"] in self.matching_case_dict[self.current_unmatched_case]:
            self.matching_case_dict[self.current_unmatched_case].remove(case["title"])
            if len(self.matching_case_dict[self.current_unmatched_case]) == 0:
                del self.matching_case_dict[self.current_unmatched_case]
            self.matching_case_num = len(self.matching_case_dict)
            case["highlighted"] = False
            self.search_results_model.layoutChanged.emit()
        elif self.current_unmatched_case != "":
            # self.matching_case_dict[self.current_unmatched_case] = case["title"] 
            # self.matching_case_num = len(self.matching_case_dict)
            # 现在将self.matching_case_dict[self.current_unmatched_case]改为一个list，存储多个匹配项
            if self.current_unmatched_case not in self.matching_case_dict:
                self.matching_case_dict[self.current_unmatched_case] = []
            if case["title"] not in self.matching_case_dict[self.current_unmatched_case]:
                self.matching_case_dict[self.current_unmatched_case].append(case["title"])
            self.matching_case_num = len(self.matching_case_dict)
            
        # search_results_model其他所有的case背景颜色恢复
        for row in range(self.search_results_model.rowCount()):
            _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
            if _case:
                _case["highlighted"] = False
        for row in range(self.search_results_model.rowCount()):
            _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
            if _case and _case["title"] in self.matching_case_dict[self.current_unmatched_case]:
                _case["highlighted"] = True             
        self.search_results_model.layoutChanged.emit()

    def on_source_selected(self, index):
        """当选择数据来源时触发"""
        self.data_source = self.source_combo.currentText()
        if self.data_source == "中国工商案例库":
            self.year_input.setVisible(False)
            self.load_progress_bar.setVisible(True)
        elif self.data_source == "华图":
            self.year_input.setVisible(True)
            self.load_progress_bar.setVisible(False)
        self.init_list()

    def on_load_data_clicked(self):     
        self.init_list()
        
        file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xls *.xlsx *.csv)")
        if not file_path:
            print("读取文件失败")
            return
             
        self.overlay.show_loading_animation()     
        if self.data_source == "中国工商案例库":
            self.read_thread: ReadTsinghuaBrowsingAndDownloadThread = ReadTsinghuaBrowsingAndDownloadThread(file_path)
            self.read_thread.result.connect(self.readBrowsingAndDownloadRecord_Tsinghua_finished)
            self.read_thread.progress.connect(self.load_progress_bar.update_progress)
            self.read_thread.start()    
        else:     
            self.thread: LoadingUIThread = LoadingUIThread(readBrowsingAndDownloadData_HuaTu, file_path, self.huatu_year)
            self.thread.data_loaded.connect(self.readBrowsingAndDownloadData_HuaTu_finished)
            self.thread.start()
            
    def readBrowsingAndDownloadRecord_Tsinghua_finished(self, returns):       
        (missingInformationBrowsingRecords, missingInformationDownloadRecords, 
            wrongBrowsingRecords, wrongDownloadRecords) = returns
        
        # *不管缺失信息了，因为读取的表是案例库系统导出的，即使有缺失信息也是案例库系统的问题
        # if len(missingInformationBrowsingRecords) > 0 or len(missingInformationDownloadRecords) > 0:
        #     print("缺失信息：")
        #     print(missingInformationBrowsingRecords)
        #     print(missingInformationDownloadRecords)
        #     QMessageBox.warning(self, "警告", "表格中信息不全！")
        #     self.overlay.setVisible(False)
        #     self.overlay.timer.stop()
        #     return
        
        # 提取wrongBrowsingRecords和wrongDownloadRecords的案例名，存储在set中
        wrongCases = set()
        for record in wrongBrowsingRecords:
            wrongCases.add(record['案例名称'])
        for record in wrongDownloadRecords:
            wrongCases.add(record['案例名称'])
        # 调用getCase函数返回Case全部信息对应的class
        
        if len(wrongCases) > 0:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle(" ")
            msg_box.setText("表格中部分案例名称无法匹配，请手动匹配！")
            msg_box.exec_()
            
            wrongCaseslist = cases_name_to_widget_list(wrongCases)
            
            self.unmatched_casename_to_download_record_Thu = {case['案例名称']: case for case in wrongDownloadRecords}
            self.unmatched_casename_to_browse_record_Thu = {case['案例名称']: case for case in wrongBrowsingRecords}
            
            self.unmatched_case_model.update_data(wrongCaseslist)
            self.unmatched_case_model.layoutChanged.emit()
            
            self.unmatched_case_num = len(wrongCases)
        else:
            # self.on_confirm_clicked()
            pass
        
        self.overlay.setVisible(False)
        self.overlay.timer.stop()
        
    def readBrowsingAndDownloadData_HuaTu_finished(self, returns):
        (missingInformationData, wrongData) = returns
        
        if len(missingInformationData) > 0:
            wrongcasedict = cases_huatu_to_widget_list(missingInformationData) 
            
            if not self.wrong_case_list_widget:
                self.wrong_case_list_widget = WrongCaseListWindow(wrongcasedict)
            else:
                self.wrong_case_list_widget.close()  # 关闭旧的窗口，防止重复
                self.wrong_case_list_widget = WrongCaseListWindow(wrongcasedict)
            self.wrong_case_list_widget.show()
        elif len(wrongData) > 0:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle(" ")
            msg_box.setText("表格中部分案例名称无法匹配，请手动匹配！")
            msg_box.exec_()
            
            unmatchedCaseslist = cases_huatu_to_widget_list(wrongData)
            self.unmatched_casename_to_class_dict_huatu = {case['标题']: case for case in wrongData}
            
            self.unmatched_case_model.update_data(unmatchedCaseslist)
            self.unmatched_case_model.layoutChanged.emit()
            
            self.unmatched_case_num = len(wrongData)
        else:
            # self.on_confirm_clicked()
            pass
        
        self.overlay.setVisible(False)
        self.overlay.timer.stop()
    
    def on_confirm_clicked(self):
        """确认按钮点击事件"""
        if len(self.matching_case_dict) == 0:
            return
        if len(self.matching_case_dict) != self.unmatched_case_num:
            QMessageBox.warning(self, "警告", "请先完成所有案例的匹配！")
              
        self.overlay.show_loading_animation()
        self.thread: LoadingUIThread = LoadingUIThread(self.update_Records)
        self.thread.data_loaded.connect(self.update_Records_finished)
        self.thread.start()
        
    def update_Records(self):
        if self.data_source == "中国工商案例库":
            for key, values in self.matching_case_dict.items():
                # if value == "未匹配":
                #     continue
                # updateCase(value, alias=key)
                for value in values:
                    if value == "未匹配":
                        continue
                    updateCase(value, alias=key)

            # 添加浏览记录和下载记录到数据库
            for key, value in self.unmatched_casename_to_download_record_Thu.items():
                for case in self.matching_case_dict[key]:
                    addDownloadRecord_Tsinghua(case, value['下载人账号'], value['下载时间'])
            for key, value in self.unmatched_casename_to_browse_record_Thu.items():
                for case in self.matching_case_dict[key]:
                    addBrowsingRecord_Tsinghua(case, value['浏览人账号'], value['浏览时间'])
        else:
            self.huatu_year = self.year_input.text()
            if self.huatu_year == "":
                QMessageBox.warning(self, "警告", "请输入年份！")
                return
            
            for key, values in self.matching_case_dict.items():
                # if value == "未匹配":
                #     continue
                # updateCase(value, alias=key)
                for value in values:
                    if value == "未匹配":
                        continue
                    updateCase(value, alias=key)
                    
            for key, value in self.unmatched_casename_to_class_dict_huatu.items():
                for case in self.matching_case_dict[key]:
                    addBrowsingAndDownloadData_HuaTu(case, self.huatu_year, value['查看数'], value['邮件数'])
            
        return None
            
    def update_Records_finished(self, returns):
        QMessageBox.information(self, "提示", "数据导入成功！")
        self.overlay.setVisible(False)
        self.overlay.timer.stop()
        self.init_list()
            
    def on_calc_clicked(self):
        year = self.calc_year_input.text()
        try:
            year = int(year)
        except ValueError:
            # 默认为今年
            year = datetime.datetime.now().year
            
        decimal_value = self.decimal_input.text()
        try:
            decimal_value = float(decimal_value)
        except ValueError:
            decimal_value = 0.3
            
        total_money = self.total_money_input.text()
        try:
            total_money = int(total_money)
        except ValueError:
            QMessageBox.warning(self, "警告", "请输入正确的总金额！")
            return
            
        square_root_selected = self.square_root_checkbox.isChecked()
        
        self.overlay.show_loading_animation()
        # 获取往年是否被计算过，确保过去每年都被计算过
        paymentCalcMessage: dict = getPaymentCalculatedYear()
        years = []
        for check_year in range(2001, year):
            if paymentCalcMessage.get(check_year-1, False) == False:
                years.append(check_year-1)
        years.append(year-1)
            
        self.calc_thread: calculatePaymentThread = calculatePaymentThread(years, total_money, decimal_value, square_root_selected)
        self.calc_thread.finished.connect(self.calculatePayment_finished)
        self.calc_thread.progress.connect(self.calc_progress_bar.update_progress)
        self.calc_thread.start()
        
    def on_export_clicked(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "保存文件", "", "Excel文件 (*.xlsx)")
        if not file_path:
            return
        
        exportCalculatedPayment(file_path)
        
    def calculatePayment_finished(self):
        QMessageBox.information(self, "提示", "计算完成！")
        self.overlay.setVisible(False)
        self.overlay.timer.stop()      
              
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
