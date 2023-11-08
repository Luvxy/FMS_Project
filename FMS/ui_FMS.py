# -*- coding: utf-8 -*-

################################################################################
## Form generated from reading UI file 'untitlednywweh.ui'
##
## Created by: Qt User Interface Compiler version 5.14.1
##
## WARNING! All changes made in this file will be lost when recompiling UI file!
################################################################################
from PyQt5.QtCore import (QCoreApplication, QMetaObject, QObject, QPoint,
    QRect, QSize, QUrl, Qt, QStringListModel,QDate)
from PyQt5.QtGui import (QBrush, QColor, QConicalGradient, QCursor, QFont,
    QFontDatabase, QIcon, QLinearGradient, QPalette, QPainter, QPixmap,
    QRadialGradient)
from PyQt5.QtWidgets import *
import sys
from button_functions import *


class Ui_MainWindow(object):
    listbox_1 = ''
    
    def setupUi(self, MainWindow):
        global listbox_1
        if MainWindow.objectName():
            MainWindow.setObjectName(u"MainWindow")
        MainWindow.resize(382, 426)
        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u"centralwidget")
        self.tabWidget = QTabWidget(self.centralwidget)
        self.tabWidget.setObjectName(u"tabWidget")
        self.tabWidget.setGeometry(QRect(0, 0, 381, 598))
        self.tab = QWidget()
        ################################################################################
        # 이용자등록 tab
        ################################################################################
        self.tab.setObjectName(u"tab")
        # 다음
        self.pushButton = QPushButton(self.tab)
        self.pushButton.setObjectName(u"pushButton")
        self.pushButton.setGeometry(QRect(90, 70, 75, 23))
        # 이전
        self.pushButton_2 = QPushButton(self.tab)
        self.pushButton_2.setObjectName(u"pushButton_2")
        self.pushButton_2.setGeometry(QRect(10, 70, 75, 23))
        # 파일선택
        self.pushButton_3 = QPushButton(self.tab)
        self.pushButton_3.setObjectName(u"pushButton_3")
        self.pushButton_3.setGeometry(QRect(10, 10, 75, 23))
        # 이용자정보 출력
        self.listView = QListView(self.tab)
        self.listView.setObjectName(u"listView")
        self.listView.setGeometry(QRect(10, 100, 361, 251))
        # 날짜선택
        self.dateEdit = QDateEdit(self.tab)
        self.dateEdit.setObjectName(u"dateEdit")
        self.dateEdit.setGeometry(QRect(90, 10, 110, 22))
        # 파일 이름 출력
        self.listView_2 = QListView(self.tab)
        self.listView_2.setObjectName(u"listView_2")
        self.listView_2.setGeometry(QRect(10, 40, 361, 21))
        #파일 정보
        self.textEdit = QTextEdit(self.tab)
        self.textEdit.setObjectName(u"textEdit")
        self.textEdit.setGeometry(QRect(10, 40, 361, 21))
        # 시작
        self.pushButton_4 = QPushButton(self.tab)
        self.pushButton_4.setObjectName(u"pushButton_4")
        self.pushButton_4.setGeometry(QRect(290, 70, 75, 23))
        # 자동화 check
        self.checkBox = QCheckBox(self.tab)
        self.checkBox.setObjectName(u"checkBox")
        self.checkBox.setGeometry(QRect(220, 70, 81, 21))
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QWidget()
        # 가져오기
        self.pushButton_12 = QPushButton(self.tab)
        self.pushButton_12.setObjectName(u"pushButton_12")
        self.pushButton_12.setGeometry(QRect(210, 10, 75, 23))
        self.tabWidget.addTab(self.tab, "")
        self.tab_2 = QWidget()
        
        ################################################################################
        # 제공등록 tab
        ################################################################################
        self.tab_2.setObjectName(u"tab_2")
        self.pushButton_5 = QPushButton(self.tab_2)
        self.pushButton_5.setObjectName(u"pushButton_5")
        self.pushButton_5.setGeometry(QRect(10, 10, 75, 23))
        self.dateEdit_2 = QDateEdit(self.tab_2)
        self.dateEdit_2.setObjectName(u"dateEdit_2")
        self.dateEdit_2.setGeometry(QRect(90, 10, 110, 22))
        self.listView_3 = QListView(self.tab_2)
        self.listView_3.setObjectName(u"listView_3")
        self.listView_3.setGeometry(QRect(10, 40, 361, 21))
        self.comboBox = QComboBox(self.tab_2)
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.addItem("")
        self.comboBox.setObjectName(u"comboBox")
        self.comboBox.setGeometry(QRect(260, 10, 101, 22))
        self.listView_4 = QListView(self.tab_2)
        self.listView_4.setObjectName(u"listView_4")
        self.listView_4.setGeometry(QRect(10, 100, 361, 251))
        self.checkBox_2 = QCheckBox(self.tab_2)
        self.checkBox_2.setObjectName(u"checkBox_2")
        self.checkBox_2.setGeometry(QRect(220, 70, 81, 21))
        self.pushButton_6 = QPushButton(self.tab_2)
        self.pushButton_6.setObjectName(u"pushButton_6")
        self.pushButton_6.setGeometry(QRect(290, 70, 75, 23))
        self.pushButton_7 = QPushButton(self.tab_2)
        self.pushButton_7.setObjectName(u"pushButton_7")
        self.pushButton_7.setGeometry(QRect(90, 70, 75, 23))
        self.pushButton_8 = QPushButton(self.tab_2)
        self.pushButton_8.setObjectName(u"pushButton_8")
        self.pushButton_8.setGeometry(QRect(10, 70, 75, 23))
        self.tabWidget.addTab(self.tab_2, "")
        ################################################################################
        # 접수등록 tab
        ################################################################################
        self.tab_3 = QWidget()
        self.tab_3.setObjectName(u"tab_3")
        self.dateEdit_3 = QDateEdit(self.tab_3)
        self.dateEdit_3.setObjectName(u"dateEdit_3")
        self.dateEdit_3.setMinimumDate(QDate(1900, 1, 1))
        self.dateEdit_3.setGeometry(QRect(10, 10, 110, 22))
        self.checkBox_3 = QCheckBox(self.tab_3)
        self.checkBox_3.setObjectName(u"checkBox_3")
        self.checkBox_3.setGeometry(QRect(210, 40, 81, 21))
        self.pushButton_9 = QPushButton(self.tab_3)
        self.pushButton_9.setObjectName(u"pushButton_9")
        self.pushButton_9.setGeometry(QRect(280, 40, 75, 23))
        self.pushButton_10 = QPushButton(self.tab_3)
        self.pushButton_10.setObjectName(u"pushButton_10")
        self.pushButton_10.setGeometry(QRect(210, 330, 75, 23))
        self.pushButton_11 = QPushButton(self.tab_3)
        self.pushButton_11.setObjectName(u"pushButton_11")
        self.pushButton_11.setGeometry(QRect(290, 330, 75, 23))
        self.line = QFrame(self.tab_3)
        self.line.setObjectName(u"line")
        self.line.setGeometry(QRect(10, 60, 351, 16))
        self.line.setFrameShape(QFrame.HLine)
        self.line.setFrameShadow(QFrame.Sunken)
        
        self.widget = QWidget(self.tab_3)
        self.widget.setObjectName(u"widget")
        self.widget.setGeometry(QRect(10, 70, 351, 21))
        self.spinBox = QSpinBox(self.widget)
        self.spinBox.setObjectName(u"spinBox")
        self.spinBox.setGeometry(QRect(310, 0, 42, 22))
        self.label = QLabel(self.widget)
        self.label.setObjectName(u"label")
        self.label.setGeometry(QRect(10, 0, 131, 21))
        
        self.widget_2 = QWidget(self.tab_3)
        self.widget_2.setObjectName(u"widget_2")
        self.widget_2.setGeometry(QRect(10, 90, 351, 21))
        self.spinBox_2 = QSpinBox(self.widget_2)
        self.spinBox_2.setObjectName(u"spinBox_2")
        self.spinBox_2.setGeometry(QRect(310, 0, 42, 22))
        self.label_2 = QLabel(self.widget_2)
        self.label_2.setObjectName(u"label_2")
        self.label_2.setGeometry(QRect(10, 0, 131, 21))
        
        self.tabWidget.addTab(self.tab_3, "")
        self.tab_4 = QWidget()
        self.tab_4.setObjectName(u"tab_4")
        self.tabWidget.addTab(self.tab_4, "")
        self.tab_5 = QWidget()
        self.tab_5.setObjectName(u"tab_5")
        self.tabWidget.addTab(self.tab_5, "")
        self.tab_6 = QWidget()
        self.tab_6.setObjectName(u"tab_6")
        self.tabWidget.addTab(self.tab_6, "")
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QMenuBar(MainWindow)
        self.menubar.setObjectName(u"menubar")
        self.menubar.setGeometry(QRect(0, 0, 382, 21))
        self.menuFNS = QMenu(self.menubar)
        self.menuFNS.setObjectName(u"menuFNS")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QStatusBar(MainWindow)
        self.statusbar.setObjectName(u"statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.menubar.addAction(self.menuFNS.menuAction())

        self.retranslateUi(MainWindow)

        self.tabWidget.setCurrentIndex(0)


        QMetaObject.connectSlotsByName(MainWindow)
    # setupUi

    def retranslateUi(self, MainWindow):
        MainWindow.setWindowTitle(QCoreApplication.translate("MainWindow", u"MainWindow", None))
        self.pushButton.setText(QCoreApplication.translate("MainWindow", u"\ub2e4\uc74c", None))
        self.pushButton_2.setText(QCoreApplication.translate("MainWindow", u"\uc774\uc804", None))
        self.pushButton_3.setText(QCoreApplication.translate("MainWindow", u"\ud30c\uc77c\uc120\ud0dd", None))
        self.pushButton_4.setText(QCoreApplication.translate("MainWindow", u"\uc2dc\uc791", None))
        self.checkBox.setText(QCoreApplication.translate("MainWindow", u"\uc790\ub3d9\ub4f1\ub85d", None))
        self.pushButton_12.setText(QCoreApplication.translate("MainWindow", u"\uac00\uc838\uc624\uae30", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab), QCoreApplication.translate("MainWindow", u"이용자등록", None))
        self.pushButton_5.setText(QCoreApplication.translate("MainWindow", u"\ud30c\uc77c\uc120\ud0dd", None))
        self.comboBox.setItemText(0, QCoreApplication.translate("MainWindow", u"\uc720\ud1b5\uae30\ud55c\uc21c", None))
        self.comboBox.setItemText(1, QCoreApplication.translate("MainWindow", u"\uae30\ubd80\uc77c\uc790\uc21c", None))
        self.comboBox.setItemText(2, QCoreApplication.translate("MainWindow", u"\uc218\ub7c9\uc801\uc740\uc21c", None))
        self.comboBox.setItemText(3, QCoreApplication.translate("MainWindow", u"\uc218\ub7c9\ub9ce\uc740\uc21c", None))

        self.checkBox_2.setText(QCoreApplication.translate("MainWindow", u"\uc790\ub3d9\ub4f1\ub85d", None))
        self.pushButton_6.setText(QCoreApplication.translate("MainWindow", u"\uc2dc\uc791", None))
        self.pushButton_7.setText(QCoreApplication.translate("MainWindow", u"\ub2e4\uc74c", None))
        self.pushButton_8.setText(QCoreApplication.translate("MainWindow", u"\uc774\uc804", None))
        
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_2), QCoreApplication.translate("MainWindow", u"제공등록", None))
        self.checkBox_3.setText(QCoreApplication.translate("MainWindow", u"\uc790\ub3d9\ub4f1\ub85d", None))
        self.pushButton_9.setText(QCoreApplication.translate("MainWindow", u"\uc2dc\uc791", None))
        self.pushButton_10.setText(QCoreApplication.translate("MainWindow", u"\ucd94\uac00", None))
        self.pushButton_11.setText(QCoreApplication.translate("MainWindow", u"\uc218\uc815", None))
    
        self.label.setText(QCoreApplication.translate("MainWindow", u"매장이름", None))
        self.label_2.setText(QCoreApplication.translate("MainWindow", u"매장이름2", None))
        
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_3), QCoreApplication.translate("MainWindow", u"접수등록", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_4), QCoreApplication.translate("MainWindow", u"접수수정", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_5), QCoreApplication.translate("MainWindow", u"제공수정", None))
        self.tabWidget.setTabText(self.tabWidget.indexOf(self.tab_6), QCoreApplication.translate("MainWindow", u"else", None))
        self.menuFNS.setTitle(QCoreApplication.translate("MainWindow", u"FNS", None))
    # retranslateUi
    
    # def setText_listview_1(self, text):
    #     # 리스트뷰에 표시할 데이터
    #     list_text_one = ''
    #     # 모델을 생성하고 데이터를 설정
    #     model = QStringListModel(list_text_one)
    #     # 리스트뷰에 모델 설정
    #     self.listView.setModel(model)

class MyMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Connect button clicks to functions
        # tab1
        self.ui.pushButton.clicked.connect(on_button1_clicked) # 이전
        self.ui.pushButton_2.clicked.connect(on_button2_clicked) # 다음
        self.ui.pushButton_3.clicked.connect(on_button3_clicked) # 파일선택
        # tab2
        self.ui.pushButton_5.clicked.connect(on_button5_clicked) # 파일선택
        self.ui.pushButton_6.clicked.connect(on_button6_clicked) # 시작
        self.ui.pushButton_7.clicked.connect(on_button7_clicked) # 다음
        self.ui.pushButton_8.clicked.connect(on_button8_clicked) # 이전
        # tab3
        self.ui.pushButton_9.clicked.connect(lambda idx:on_button9_clicked(self.ui.dateEdit_3.date())) # 시작
        
    def get_main_window(self):
        return self.ui

def main():
    app = QApplication(sys.argv)
    main_window = MyMainWindow()
    main_window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()