from PyQt5.QtWidgets import QFileDialog, QPushButton, QGraphicsDropShadowEffect, QComboBox, QListView, QProgressBar, QAbstractItemView
from PyQt5.QtCore import Qt, QThread, pyqtSignal
import pandas as pd

from src.db.init_db import *

class LoadingUIThread(QThread):
    data_loaded = pyqtSignal(tuple)
    
    def __init__(self, func, *args, **kargs):
        super().__init__()
        self.func = func
        self.args = args
        self.kargs = kargs

    def run(self):
        print("开始加载数据...")

        returns = self.func(*self.args, **self.kargs)
        if not isinstance(returns, tuple):
            returns = (returns, )

        print("数据加载完成")
        # 数据加载完成后，发出信号通知主线程
        self.data_loaded.emit(returns)

def load_data(widget, file_path=None):
    """加载 Excel/CSV 数据并填充表格"""
    if file_path is None:
        file_path, _ = QFileDialog.getOpenFileName(widget, "选择文件", "", "Excel文件 (*.xls *.xlsx *.csv)")
    if not file_path:
        print("读取文件失败！")
        return
    
    # 读取数据
    if file_path.endswith(".csv"):
        data = pd.read_csv(file_path)
    elif file_path.endswith(".xls"):
        data = pd.read_excel(file_path, engine="xlrd", sheet_name=None)  # 读取所有 sheets
    else:
        data = pd.read_excel(file_path, engine="openpyxl", sheet_name=None)  # 读取所有 sheets

    if data is None or (isinstance(data, pd.DataFrame) and data.empty) or (isinstance(data, dict) and len(data) == 0):
        print("读取的数据为空！")
        return

    print(f"数据加载成功！文件: {file_path}")
    return data, file_path

def set_button_style(button: QPushButton, height=60):
    button.setFixedHeight(60)
    button.setStyleSheet("""
        QPushButton {
            font-size: 24px;
            font-family: '黑体';
            margin-bottom: 10px;
            background-color: white;
            border: none;
            border-radius: 8px;
        }
        QPushButton:hover {
            background-color: #e0e0e0;
        }
        QPushButton:pressed {
            background-color: #c0c0c0;
        }
    """)
    
    shadow_effect = QGraphicsDropShadowEffect()
    shadow_effect.setBlurRadius(10)
    shadow_effect.setOffset(0, 0)
    shadow_effect.setColor(Qt.gray)
    
    button.setGraphicsEffect(shadow_effect)
    
def set_combo_style(combo: QComboBox):
    combo.setStyleSheet("""
        QComboBox {
            background-color: #ffffff;
            border: 2px solid #d7d7d7;
            border-radius: 8px;
            padding: 5px;
            color: #717171;
            min-width: 150px;
        }
        QComboBox:hover {
            border-color: #7b56f0;
        }
        QComboBox::drop-down {
            border: none;
            background: transparent;
        }
        QComboBox::down-arrow {
            image: url(down_arrow.png);
            width: 14px;
            height: 14px;
        }
        QComboBox QAbstractItemView {
            border: 1px solid #aaa;
            background: #ffffff;
            selection-background-color: #a68fff;
            selection-color: white;
            padding: 4px;
            border-radius: 5px;
        }
        QComboBox QAbstractItemView::item {
            min-height: 40px;
            padding: 10px;
        }
    """)
    
def set_scrollbar_style(list_view: QListView):
    list_view.setStyleSheet("""
        QListView {
            border: none;
        }
        QScrollBar:vertical {
            border: none;
            background: white;
            width: 12px;  
            margin: 0px 0px 0px 0px;
        }
        QScrollBar::handle:vertical {
            background: rgba(160, 160, 160, 150);
            border-radius: 4px;
            min-height: 20px;
        }
        QScrollBar::handle:vertical:hover {
            background: rgba(120, 120, 120, 200);
        }
        QScrollBar::add-line:vertical, 
        QScrollBar::sub-line:vertical {
            border: none;
            background: none;
        }
        QScrollBar::add-page:vertical, 
        QScrollBar::sub-page:vertical {
            background: none;
        }
    """)
    
    list_view.setVerticalScrollMode(QAbstractItemView.ScrollPerPixel)
    list_view.verticalScrollBar().setSingleStep(20)

# 添加前端展示需要的字段
def case_dict_to_widget_list(case_dict):
    if case_dict is None:
        return ""
    # 检查case_dict是否有键"错误信息"
    if '序号' not in case_dict and '发布时间' not in case_dict and '案例编号' not in case_dict:
        # 把case_dict中的所有键值信息转换为info
        case_dict['info'] = "\n".join([f"{key}: {value}" for key, value in case_dict.items()])
    else:
        if "错误信息" in case_dict:
            case_dict['info'] = f"序号：{case_dict['序号']}\n错误信息：{case_dict['错误信息']}"
        else:
            case_dict['info'] = f"序号：{case_dict['序号']}\n发布时间：{case_dict['发布时间']}，案例编号：{case_dict['案例编号']}"
    # case_dict['title'] = case_dict['案例标题']
    title = None
    for attr in case_dict:
        if '标题' in attr:
            title = attr
    case_dict['title'] = case_dict.get(title, "无标题")
    return case_dict

# 批量添加前端展示需要的字段
def cases_dict_to_widget_list(cases_dict):
    return [case_dict_to_widget_list(case) for case in cases_dict]

# 特指中国工商案例库的类别
def case_class_to_widget_list(case):
    if case is None:
        return ""
    
    case_dict = dict()
    case_dict['title'] = case.name
    case_dict['info'] = f"发布时间：{case.release_time}，案例编号：{case.submission_number}"
    return case_dict
# 特指中国工商案例库的类别
def cases_class_to_widget_list(cases):
    return [case_class_to_widget_list(case) for case in cases]

def case_name_to_widget_list(case_name):
    if case_name is None:
        return ""
    
    case_dict = dict()
    case_dict['title'] = case_name
    case_dict['info'] = case_name
    return case_dict

def cases_name_to_widget_list(cases):
    return [case_name_to_widget_list(case) for case in cases]

def case_huatu_to_widget_list(case_dict):
    if case_dict is None:
        return ""
    
    case_dict['title'] = case_dict['标题']
    if '错误信息' in case_dict:
        case_dict['info'] = f"序号：{case_dict['序号']}\n错误信息：{case_dict['错误信息']}"
    else:
        case_dict['info'] = f"序号：{case_dict['序号']}\n出版时间：{case_dict['出版时间']}，邮件数：{case_dict['邮件数']}，查看数：{case_dict['查看数']}"
    return case_dict

def cases_huatu_to_widget_list(cases_dict):
    return [case_huatu_to_widget_list(case) for case in cases_dict]

# 大概率不完整
# def cases_to_widget_list(cases):
#     cases_widget_list = []
#     if isinstance(cases, dict):
#         cases_widget_list = cases_dict_to_widget_list(cases)
#     elif isinstance(cases, list):
#         cases_widget_list = cases_name_to_widget_list(cases)
#     elif isinstance(cases, Case):
#         cases_widget_list = cases_class_to_widget_list(cases)
#     elif isinstance(cases, HuaTuData):
#         cases_widget_list = cases_huatu_to_widget_list(cases)
#     return cases_widget_list