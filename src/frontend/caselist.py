import sys
from PyQt5.QtWidgets import ( QApplication, QWidget, QVBoxLayout, QListView, QLabel, 
    QStyledItemDelegate, QSizePolicy, QTextEdit
)
from PyQt5.QtGui import QFontMetrics, QTextOption, QTextDocument, QFont, QColor, QBrush, QPainterPath
from PyQt5.QtCore import Qt, QSize, QAbstractListModel, QRectF

from src.frontend.utils import set_scrollbar_style

# 数据模型
class CaseListModel(QAbstractListModel):
    def __init__(self, cases=None):
        super().__init__()
        self.cases = cases or []

    def rowCount(self, parent=None):
        return len(self.cases)

    def data(self, index, role):
        if not index.isValid() or index.row() >= len(self.cases):
            return None
        case = self.cases[index.row()]
        if role == Qt.DisplayRole:
            return case  # 这里只存储案例数据
        return None
    
    def update_data(self, cases):
        self.beginResetModel()
        self.cases = cases
        self.endResetModel()

# 自定义渲染代理
import math

class CaseItemDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        case = index.data(Qt.DisplayRole)
        if case:
            painter.save()

            rect = option.rect.adjusted(20, 20, -20, -20)  # 设置内边距
            text_width = option.rect.width() - 40  # 让文本适应 `QListView`

            if case.get("highlighted", False):
                color_rect = option.rect.adjusted(10, 10, -10, -10)
                _color_rect_ = QRectF(color_rect)  # 用于绘制圆角矩形
                path = QPainterPath()
                path.addRoundedRect(_color_rect_, 10, 10)  # 10 是圆角的半径
                painter.fillPath(path, QBrush(QColor(246, 242, 255)))
            
            # 使用 QTextDocument 渲染标题
            title_doc = QTextDocument()
            title_doc.setTextWidth(text_width)
            title_doc.setHtml(f"""
                <span style="font-family: '黑体'; font-size: 12pt; color: black;">
                    {case["title"]}
                </span>
            """)
            
            title_height = title_doc.documentLayout().documentSize().height()
            title_rect = QRectF(rect.left(), rect.top(), text_width, title_height)
            
            painter.save()
            painter.translate(title_rect.topLeft())
            title_doc.drawContents(painter)
            painter.restore()

            # 使用 QTextDocument 渲染内容
            content_doc = QTextDocument()
            content_doc.setTextWidth(text_width)
            content_doc.setHtml(f"""
                <span style="font-family: Arial; font-size: 10pt; color: rgb(100, 100, 100);">
                    {case["info"]}
                </span>
            """)
            
            content_height = content_doc.documentLayout().documentSize().height()
            content_pos = rect.adjusted(0, int(title_height) + 5, 0, 0)  # 让内容在标题下方
            painter.translate(content_pos.topLeft())
            content_doc.drawContents(painter)
            
            painter.restore()

    def sizeHint(self, option, index):
        case = index.data(Qt.DisplayRole)
        if case:
            content_width = option.rect.width() - 40  # 让文本适应 `QListView`

            # 使用 QTextDocument 计算标题高度
            title_doc = QTextDocument()
            title_doc.setTextWidth(content_width)
            title_doc.setHtml(f"""
                <span style="font-family: '黑体'; font-size: 12pt; color: black;">
                    {case["title"]}
                </span>
            """)
            title_height = title_doc.documentLayout().documentSize().height()

            # 使用 QTextDocument 计算内容高度
            content_doc = QTextDocument()
            content_doc.setTextWidth(content_width)
            content_doc.setHtml(f"""
                <span style="font-family: Arial; font-size: 10pt; color: rgb(100, 100, 100);">
                    {case["info"]}
                </span>
            """)
            content_height = content_doc.documentLayout().documentSize().height()

            total_height = math.ceil(title_height + content_height + 40 + 5)  # 标题 + 内容 + 内边距
            return QSize(option.rect.width(), total_height)
        return QSize(300, 50)  # 默认大小

def get_case_list_widget(cases = None):    
    # 返回设置好的model和view
    case_list_model = CaseListModel(cases)
    case_list_view = QListView()
    case_list_view.setModel(case_list_model)
    case_list_view.setItemDelegate(CaseItemDelegate())
    case_list_view.setWordWrap(True)  # 启用自动换行
    case_list_view.setResizeMode(QListView.Adjust)  # 适应内容大小
    case_list_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
    case_list_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    case_list_view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    
    set_scrollbar_style(case_list_view)
    
    return case_list_model, case_list_view