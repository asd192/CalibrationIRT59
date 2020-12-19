import sys
from PyQt5 import QtWidgets
from settings import Ui_settings

class WindowSettings(QtWidgets.QMainWindow):
    def __init__(self):
        super(WindowSettings, self).__init__()
        self.st = Ui_settings()
        self.st.setupUi(self)

if __name__ == "__main__":
    s_app = QtWidgets.QApplication([])
    s_application = WindowSettings()
    s_application.setWindowTitle("Настройки")
    s_application.show()

    sys.exit(s_app.exec())