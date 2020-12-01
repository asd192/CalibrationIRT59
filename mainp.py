from PyQt5 import QtWidgets, QtCore
from main import Ui_MainWindow
import sys


class window(QtWidgets.QMainWindow):
    def __init__(self):
        super(window, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Прибор с ПВИ
        self.ui.checkBox_pvi.stateChanged.connect(lambda check=self.ui.checkBox_pvi.isChecked(): self.select_pvi(check))

        # Установка входных значений
        # TODO
        self.ui.lineEdit_out_irt_in_5.setText(self.ui.lineEdit_in_start_value.text())
        self.ui.lineEdit_in_start_value.textChanged.connect(self.out_irt_in)
        self.ui.lineEdit_in_end_value.textChanged.connect(self.out_irt_in)

    def select_pvi(self, pvi_on):
        if pvi_on == QtCore.Qt.Checked:
            self.ui.label_pvi_scale.setEnabled(True)
            self.ui.lineEdit_pvi_scale_start.setEnabled(True)
            self.ui.lineEdit_pvi_scale_end.setEnabled(True)
            self.ui.label_pvi_out.setEnabled(True)
            self.ui.comboBox_pvi_out.setEnabled(True)
            self.ui.label_pvi_out_r.setEnabled(True)
        else:
            self.ui.label_pvi_scale.setEnabled(False)
            self.ui.lineEdit_pvi_scale_start.setEnabled(False)
            self.ui.lineEdit_pvi_scale_end.setEnabled(False)
            self.ui.label_pvi_out.setEnabled(False)
            self.ui.comboBox_pvi_out.setEnabled(False)
            self.ui.label_pvi_out_r.setEnabled(False)

    def out_irt_in(self):
        in_start = self.ui.lineEdit_in_start_value.text()
        in_end = self.ui.lineEdit_in_end_value.text()

        values = []
        for i in (0.05, 0.25, 0.5, 0.75, 0.95):
            values.append(str((abs(float(in_start)) + abs(float(in_end))) * i))

        self.ui.lineEdit_out_irt_in_5.setText(values[0])
        self.ui.lineEdit_out_irt_in_25.setText(values[1])
        self.ui.lineEdit_out_irt_in_50.setText(values[2])
        self.ui.lineEdit_out_irt_in_75.setText(values[3])
        self.ui.lineEdit_out_irt_in_95.setText(values[4])


app = QtWidgets.QApplication([])
application = window()
application.setWindowTitle("Создание протокола калибровки ИРТ 59хх")
application.show()

sys.exit(app.exec())
