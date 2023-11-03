import sys
from PyQt5.QtWidgets import QApplication, QMainWindow
from ui_FMS import Ui_MainWindow  # Replace 'ui_file' with the actual module name.
from button_functions import *

class MyMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Connect button clicks to functions
        self.ui.pushButton.clicked.connect(on_button1_clicked)
        self.ui.pushButton_2.clicked.connect(on_button2_clicked)
        self.ui.pushButton_3.clicked.connect(on_button3_clicked)

def main():
    app = QApplication(sys.argv)
    main_window = MyMainWindow()
    main_window.show()
    sys.exit(app.exec_())

if __name__ == '__main__':
    main()
