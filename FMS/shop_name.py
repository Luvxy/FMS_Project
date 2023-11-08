from PyQt5.QtCore import (QCoreApplication, QMetaObject, QObject, QPoint,
    QRect, QSize, QUrl, Qt, QStringListModel)
from PyQt5.QtGui import (QBrush, QColor, QConicalGradient, QCursor, QFont,
    QFontDatabase, QIcon, QLinearGradient, QPalette, QPainter, QPixmap,
    QRadialGradient)
from PyQt5.QtWidgets import *
import sys
from button_functions import *
import ui_FMS

class shop_name(ui_FMS):
    index = 0
    
    def __init__(self):
        self.index += 1
        self.widget = QWidget(self.tab_3)
        self.widget.setObjectName(u"widget")
        self.widget.setGeometry(QRect(10, 70, 351, 21))
        self.spinBox = QSpinBox(self.widget)
        self.spinBox.setObjectName(u"spinBox")
        self.spinBox.setGeometry(QRect(310, 0, 42, 22))
        self.label = QLabel(self.widget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 0, 131, 21))