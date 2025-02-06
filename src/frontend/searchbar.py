from PyQt5.QtWidgets import QWidget, QLineEdit, QPushButton, QHBoxLayout, QSpacerItem, QSizePolicy
from PyQt5.QtGui import QFont
from PyQt5.QtCore import Qt

class SearchBar(QWidget):
    def __init__(self, placeholder_text="搜索案例", search_callback=None, parent=None):
        super().__init__(parent)

        # 主布局（水平布局）
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)  # 去除边距
        layout.setSpacing(0)  # 控件间距为 0

        # 搜索框 (QLineEdit)
        self.search_input = QLineEdit(self)
        self.search_input.setPlaceholderText(placeholder_text)
        self.search_input.setFixedHeight(50)
        self.search_input.setFont(QFont("黑体", 10))
        self.search_input.setStyleSheet("""
            QLineEdit {
                border: none;
                border: 2px solid rgb(200, 200, 200);
                border-radius: 8px;
                padding: 10px;
                background-color: white;
            }
            QLineEdit:focus {
                border: 2px solid rgb(120, 120, 255);
                outline: none;
            }
        """)

        # 搜索按钮 (QPushButton)
        self.search_button = QPushButton("搜索", self)
        self.search_button.setFixedSize(80, 50)  # 按钮大小
        self.search_button.setFont(QFont("黑体", 10))
        self.search_button.setCursor(Qt.PointingHandCursor)  # 鼠标悬停时显示手型
        self.search_button.setStyleSheet("""
            QPushButton {
                background-color: #783cd5;  /* 紫色背景 */
                color: white;
                border-radius: 8px;
                border: none;
            }
            QPushButton:hover {
                background-color: #5d27b0;
            }
            QPushButton:pressed {
                background-color: #451b9a;
            }
        """)

        # 绑定按钮点击事件
        if search_callback:
            self.search_button.clicked.connect(search_callback)

        # 将控件添加到布局
        layout.addWidget(self.search_input)
        layout.addItem(QSpacerItem(10, 0, QSizePolicy.Minimum, QSizePolicy.Expanding))
        layout.addWidget(self.search_button)

        # 设定整体大小策略
        self.setLayout(layout)
        self.setFixedHeight(50)

    # 获取用户输入
    def get_text(self):
        return self.search_input.text()

    # 清空搜索框
    def clear_text(self):
        self.search_input.clear()
