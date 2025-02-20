from PyQt5.QtWidgets import (
    QWidget
)
from PyQt5.QtGui import QPainter, QPen, QColor
from PyQt5.QtCore import Qt, QTimer

class OverlayWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        if parent:
            self.setGeometry(parent.rect())
        self.angle = 0
        self.setVisible(False)  # 初始时隐藏遮罩层
        
        # 设置一个定时器来控制动画的更新
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.update_animation)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        painter.setPen(Qt.NoPen)  # 不需要边框
        painter.setBrush(QColor(0, 0, 0, 100))  # 设置半透明的黑色背景
        painter.drawRect(self.rect())  # 绘制填充整个窗口的矩形
          
        painter.setPen(QPen(QColor(255, 255, 255), 10))
        # painter.setBrush(QColor(0, 0, 0))
        rect = self.rect()
        painter.translate(rect.center())
        painter.rotate(self.angle)  # 从父窗口获取旋转角度
        painter.drawArc(-50, -50, 100, 100, 0, 180 * 16)  # 绘制半圆
        
    def update_animation(self):
        """更新旋转角度"""
        self.angle += 15  # 增加角度
        if self.angle >= 360:
            self.angle = 0
        # self.overlay.repaint()
        self.update()  # 重绘界面
        
    def show_loading_animation(self):
        """显示加载动画和遮罩层"""
        self.setVisible(True)  # 显示遮罩层
        self.timer.start(50)  # 启动定时器，更新动画