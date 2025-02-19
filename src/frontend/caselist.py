import sys
from PyQt5.QtWidgets import ( QApplication, QWidget, QVBoxLayout, QListView, QLabel, 
    QStyledItemDelegate, QSizePolicy, QTextEdit, QMenu, QGraphicsOpacityEffect
)
from PyQt5.QtGui import QFontMetrics, QTextOption, QTextDocument, QFont, QColor, QBrush, QPainterPath
from PyQt5.QtCore import Qt, QSize, QAbstractListModel, QRectF, QTimer, QTime

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
            # painter.save()

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
            
            matched_height = -5
            if case.get("matched_str", "") != "":
                matched_doc = QTextDocument()
                matched_doc.setTextWidth(text_width)
                matched_doc.setHtml(f"""
                    <span style="font-family: Arial; font-size: 10pt; color: rgb(155, 101, 223);">
                        {case.get("matched_str", "")}
                    </span>
                """)
                
                matched_height = matched_doc.documentLayout().documentSize().height()
                matched_pos = rect.adjusted(0, int(title_height) + 5, 0, 0)
            
                painter.save()
                painter.translate(matched_pos.topLeft())
                matched_doc.drawContents(painter)
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
            content_pos = rect.adjusted(0, int(title_height) + int(matched_height) + 10, 0, 0)  # 让内容在匹配文本下方
            
            painter.save()
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
            
            matched_height = -5
            if case.get("matched_str", "") != "":
                matched_doc = QTextDocument()
                matched_doc.setTextWidth(content_width)
                matched_doc.setHtml(f"""
                    <span style="font-family: Arial; font-size: 10pt; color: rgb(186, 85, 211);">
                        {case.get("matched_str", "")}
                    </span>
                """)
                matched_height = matched_doc.documentLayout().documentSize().height()

            # 使用 QTextDocument 计算内容高度
            content_doc = QTextDocument()
            content_doc.setTextWidth(content_width)
            content_doc.setHtml(f"""
                <span style="font-family: Arial; font-size: 10pt; color: rgb(100, 100, 100);">
                    {case["info"]}
                </span>
            """)
            content_height = content_doc.documentLayout().documentSize().height()

            total_height = math.ceil(title_height + matched_height + content_height + 40 + 10)  # 标题 + 内容 + 内边距
            return QSize(option.rect.width(), total_height)
        return QSize(300, 50)  # 默认大小

class CaseListView(QListView):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.fade_duration = 1000  # 总共的渐变持续时间（1秒）
        self.fade_in_time = 500  # 前0.5秒（不透明度保持1.0）

    def contextMenuEvent(self, event):
        index = self.indexAt(event.pos())
        if index.isValid():
            case = index.data(Qt.DisplayRole)
            self.copy_title_to_clipboard(case["title"])
            # menu = QMenu(self)
            # copy_action = menu.addAction("复制标题")
            
            # # 复制标题到剪贴板
            # copy_action.triggered.connect(lambda: self.copy_title_to_clipboard(case["title"]))
            # menu.exec_(event.globalPos())

    def copy_title_to_clipboard(self, title):
        clipboard = QApplication.clipboard()
        clipboard.setText(title)
        
        # 显示复制成功提示
        self.show_copy_success_tip()

    def show_copy_success_tip(self):
        tip_label = QLabel("复制成功", self)
        tip_label.setFont(QFont("黑体", 12))
        tip_label.setStyleSheet("background-color: rgba(112, 71, 240, 150); color: white; padding: 10px; border-radius: 5px;")
        tip_label.setAlignment(Qt.AlignCenter)
        tip_label.move(self.width() // 2 - tip_label.width() // 2, self.height() // 2 - tip_label.height() // 2)
        tip_label.show()

        # 使用 QGraphicsOpacityEffect 实现渐变透明效果
        opacity_effect = QGraphicsOpacityEffect()
        tip_label.setGraphicsEffect(opacity_effect)

        # 设置初始透明度
        opacity_effect.setOpacity(1.0)

        # 定时器控制透明度渐变
        fade_out_timer = QTimer(self)
        fade_out_timer.timeout.connect(lambda: self.fade_out(tip_label, opacity_effect, fade_out_timer))
        fade_out_timer.start(50)  # 每50ms更新一次透明度

        # 记录开始时间
        self.start_time = self.current_time_ms()

    def fade_out(self, tip_label, opacity_effect, timer):
        elapsed_time = self.current_time_ms() - self.start_time
        total_time = self.fade_duration

        if elapsed_time < self.fade_in_time:
            # 前 0.5 秒，保持不变，透明度为 1.0
            opacity_effect.setOpacity(1.0)
        elif elapsed_time < total_time:
            # 在后0.5秒内渐变透明度
            progress = (elapsed_time - self.fade_in_time) / (total_time - self.fade_in_time)  # 计算剩余时间的进度
            opacity = 1.0 - progress  # 从1.0渐变到0.0
            opacity_effect.setOpacity(opacity)
        else:
            # 完成渐变，删除tip_label
            timer.stop()
            tip_label.deleteLater()

    def current_time_ms(self):
        """返回当前的毫秒时间戳"""
        return int(QTime.currentTime().msecsSinceStartOfDay())

def get_case_list_widget(cases = None):    
    # 返回设置好的model和view
    case_list_model = CaseListModel(cases)
    # case_list_view = QListView()
    case_list_view = CaseListView()
    case_list_view.setModel(case_list_model)
    case_list_view.setItemDelegate(CaseItemDelegate())
    case_list_view.setWordWrap(True)  # 启用自动换行
    case_list_view.setResizeMode(QListView.Adjust)  # 适应内容大小
    case_list_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
    case_list_view.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    case_list_view.setVerticalScrollBarPolicy(Qt.ScrollBarAsNeeded)
    
    set_scrollbar_style(case_list_view)
    
    return case_list_model, case_list_view