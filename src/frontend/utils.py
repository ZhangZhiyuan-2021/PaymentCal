from PyQt5.QtWidgets import QFileDialog, QPushButton, QGraphicsDropShadowEffect
from PyQt5.QtCore import Qt
import pandas as pd

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
    return data

def set_button_style(button: QPushButton, height=60):
    button.setFixedHeight(60)
    button.setStyleSheet("""
        QPushButton {
            font-size: 24px;
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
    shadow_effect.setBlurRadius(30)
    shadow_effect.setOffset(0, 0)
    shadow_effect.setColor(Qt.gray)
    
    button.setGraphicsEffect(shadow_effect)