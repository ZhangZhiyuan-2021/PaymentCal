from PyQt5.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QTableWidget, QTableWidgetItem, 
    QLabel, QComboBox, QProgressBar, QSizePolicy, QSpacerItem, QFileDialog, QMessageBox,
    QLineEdit
)
from PyQt5.QtGui import QFont, QIcon, QIntValidator, QPainter, QPen, QColor
from PyQt5.QtCore import Qt, QTimer

from src.frontend.caselist import get_case_list_widget
from src.frontend.searchbar import SearchBar
from src.frontend.utils import *
from src.frontend.wrongCaseListWidget import WrongCaseListWindow
from src.frontend.overlayWidget import OverlayWidget
from src.backend.read_case import *

class ProgressBar(QProgressBar):
    def __init__(self, parent=None):
        super(ProgressBar, self).__init__(parent)
        self.setRange(0, 100)
        self.setValue(0)
        # self.setFormat("%v/%m")
        self.setAlignment(Qt.AlignCenter)
        self.setStyleSheet("""
            QProgressBar {
                border: 2px solid #c09cfe;  /* 进度条边框 */
                border-radius: 8px;         /* 圆角边框 */
                background-color: #c4b7e1;  /* 背景颜色 */
                text-align: center;         /* 文字居中 */
                font-size: 18px;            /* 文字大小 */
                color: white;             /* 文字颜色 */
                padding: 2px;
            }
            
            QProgressBar::chunk {
                background-color: #672aca;  /* 进度条颜色 */
                border-radius: 6px;         /* 进度条圆角 */
            }
        """)
        
        self.current_value = 0
        
    def update_progress(self, value):
        self.current_value = value
        self.setValue(value)
        self.update()