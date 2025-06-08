from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
    QLabel, QComboBox, QProgressBar, QSizePolicy, QSpacerItem, QFileDialog, QMessageBox,
    QLineEdit
)
from PyQt5.QtGui import QFont, QIcon, QIntValidator
from PyQt5.QtCore import Qt, QTimer, QEvent

from src.frontend.caselist import get_case_list_widget
from src.frontend.searchbar import SearchBar
from src.frontend.utils import *
from src.frontend.wrongCaseListWidget import WrongCaseListWindow
from src.frontend.overlayWidget import OverlayWidget
from src.backend.read_case import *

class ImportPaymentWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.initUI()

        self.matching_case_dict = {}    # 键是匹配错误的案例名，值是选择的对应的数据库中的案例名
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
        
        top_layout.addWidget(self.load_button)
    
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
        
        self.overlay = OverlayWidget(self)
        
    def resizeEvent(self, event):
        """调整遮罩层的大小以匹配窗口"""
        self.overlay.setGeometry(self.rect())
        super().resizeEvent(event)  # 确保父类的 resizeEvent 也被调用
    
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
        
        # 如果self.matching_case_dict[self.current_unmatched_case]有值，且当前index选中了同一个case，则取消匹配
        if self.current_unmatched_case in self.matching_case_dict and case["title"] in self.matching_case_dict[self.current_unmatched_case]:
            self.matching_case_dict[self.current_unmatched_case].remove(case["title"])
            if len(self.matching_case_dict[self.current_unmatched_case]) == 0:
                del self.matching_case_dict[self.current_unmatched_case]
            self.matching_case_num = len(self.matching_case_dict)
            case["highlighted"] = False
            self.search_results_model.layoutChanged.emit()
        elif self.current_unmatched_case != "":
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

    def on_load_data_clicked(self):     
        try:
            self.init_list()
            
            file_path, _ = QFileDialog.getOpenFileName(self, "选择文件", "", "Excel文件 (*.xls *.xlsx *.csv)")
            if not file_path:
                print("读取文件失败")
                return
              
            self.overlay.show_loading_animation()
            self.thread: LoadingUIThread = LoadingUIThread(readRealPaymentData, file_path)
            self.thread.data_loaded.connect(self.readRealPaymentData_finfished)
            self.thread.start()        
                
        except Exception as e:
            QMessageBox.warning(self, "错误", f"{str(e)}")
            return
        
    def readRealPaymentData_finfished(self, returns):
        (missingInformationData, wrongData) = returns
        
        if len(missingInformationData) > 0:
            wrong_cases = cases_dict_to_widget_list(missingInformationData)
            
            if not self.wrong_case_list_widget:
                self.wrong_case_list_widget = WrongCaseListWindow(wrong_cases)
            else:
                self.wrong_case_list_widget.close()     # 关闭之前的窗口
                self.wrong_case_list_widget = WrongCaseListWindow(wrong_cases)
            self.wrong_case_list_widget.show()
            
        elif len(wrongData) > 0:
            msg_box = QMessageBox()
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle(" ")
            msg_box.setText("表格中部分案例名称无法匹配，请手动匹配！")
            msg_box.exec_()
            
            title = None
            for attr in wrongData[0]:
                if '标题' in attr:
                    title = attr
                    break
                
            unmatchedCases_Names = [case[title] for case in wrongData]
            unmatchedCaseslist = cases_name_to_widget_list(unmatchedCases_Names)
            
            self.unmatched_case_model.update_data(unmatchedCaseslist)
            
            self.unmatched_case_model.layoutChanged.emit()
            self.unmatched_case_num = len(unmatchedCases_Names)
            
        else:
            QMessageBox.information(self, " ", "数据加载成功！")
            
        self.overlay.setVisible(False)
        self.overlay.timer.stop()

    def on_confirm_clicked(self):
        if len(self.matching_case_dict) == self.unmatched_case_num:
            for key, values in self.matching_case_dict.items():
                for value in values:
                    print(f"{key} -> {value}")
                    if value == "未匹配":
                        continue
                    updateCase(value, alias=key)
                
        self.init_list()
        
    def on_search_clicked(self):
        """搜索按钮点击事件"""
        keyword = self.search_input.get_text()
        print(f"搜索关键词: {keyword}")
        matched_strs, similar_cases = zip(*getSimilarCases(keyword))
        
        if len(similar_cases) > 0:
            similar_caseslist = cases_class_to_widget_list(similar_cases)
            # 在similar_caseslist头部添加“不匹配”项
            similar_caseslist.insert(0, {"title": "未匹配", "highlighted": False, "matched_str": "", "info": "如果找不到匹配项，请选择此项"})
            self.search_results_model.update_data(similar_caseslist)
            
            # 如果当前case已经匹配，则高亮search_results_model中对应的case；如果未匹配，则不高亮
            # 从1开始，因为第0个是“未匹配”项
            for row in range(1, self.search_results_model.rowCount()):
                _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
                if _case:
                    _case["highlighted"] = False
                    
                _case['matched_str'] = f"匹配项：{matched_strs[row-1]}"
            # 如果当前case已经匹配，则高亮search_results_model中对应的case
            if self.current_unmatched_case in self.matching_case_dict:
                for row in range(self.search_results_model.rowCount()):
                    _case = self.search_results_model.data(self.search_results_model.index(row, 0), Qt.DisplayRole)
                    if _case and _case["title"] in self.matching_case_dict[self.current_unmatched_case]:
                        _case["highlighted"] = True   
            
            self.search_results_model.layoutChanged.emit()
        else:
            print("未找到相关案例！")
            self.search_results_model.update_data([])
            
    def init_list(self):
        self.matching_case_dict = {}
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
