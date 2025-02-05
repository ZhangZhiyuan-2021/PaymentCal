import sys
from PyQt5.QtWidgets import ( QApplication, QWidget, QVBoxLayout, QListView, QLabel, 
    QStyledItemDelegate, QSizePolicy
)
from PyQt5.QtGui import QFontMetrics, QTextOption, QTextDocument, QFont, QColor, QBrush
from PyQt5.QtCore import Qt, QSize, QAbstractListModel, QRectF

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

# 自定义渲染代理
class CaseItemDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        case = index.data(Qt.DisplayRole)
        if case:
            painter.save()

            rect = option.rect.adjusted(10, 10, -10, -10)  # 设置内边距
            font_metrics = painter.fontMetrics()
            
            if case.get("highlighted", False):
                painter.fillRect(option.rect, QBrush(QColor(249, 247, 255)))
            
            text_width = option.rect.width() - 20  # 让文本适应 `QListView`
            title_height = font_metrics.height()
            title_rect = rect.adjusted(0, 0, -10, 0)
            
            title_font = QFont("Arial", 12)  # 加粗、12号字体
            painter.setFont(title_font)
            painter.setPen(QColor(0, 0, 0))
            painter.drawText(title_rect, Qt.AlignLeft | Qt.AlignTop, case["title"])

            # 设置内容字体
            content_doc = QTextDocument()
            content_doc.setTextWidth(text_width)  # 确保内容不会超出 text_width
            content_doc.setHtml(f"""
                <span style="font-family: Arial; font-size: 10pt; color: rgb(100, 100, 100);">
                    {case["info"]}
                </span>
            """)

            content_pos = rect.adjusted(0, title_height + 5, 0, 0)  # 让内容在标题下方
            painter.translate(content_pos.topLeft())  # 移动绘制起点
            content_doc.drawContents(painter)

            painter.restore()


    def sizeHint(self, option, index):
        case = index.data(Qt.DisplayRole)
        if case:
            font_metrics = option.fontMetrics
            title_height = font_metrics.height()
            
            content_width = option.rect.width() - 20  # 让文本适应 `QListView`

            # 改用 QTextDocument 计算高度
            doc = QTextDocument()
            doc.setTextWidth(content_width)  # **指定文本宽度，启用自动换行**
            doc.setPlainText(case["info"])
            content_height = doc.size().height()  # 获取计算出的文本高度

            total_height = title_height + content_height + 40  # 标题 + 内容 + 内边距
            return QSize(option.rect.width(), int(total_height))
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
    return case_list_model, case_list_view