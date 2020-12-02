from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtGui import QDoubleValidator

from main import Ui_MainWindow
import sys


class window(QtWidgets.QMainWindow):
    def __init__(self):
        super(window, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Валидация
        self.ui.lineEdit_in_start_value.setValidator(self.validat_p())
        self.ui.lineEdit_in_end_value.setValidator(self.validat_p())
        self.ui.lineEdit_out_start_value.setValidator(self.validat_p())
        self.ui.lineEdit_out_end_value.setValidator(self.validat_p())

        self.ui.lineEdit_pvi_scale_start.setValidator(self.validat_p())
        self.ui.lineEdit_pvi_scale_start.setValidator(self.validat_p())

        self.ui.lineEdit_out_irt_value_5.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_irt_value_25.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_irt_value_50.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_irt_value_75.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_irt_value_95.setValidator(self.validat_in_out())

        self.ui.lineEdit_out_24_value_0.setValidator(self.valdat_24())
        self.ui.lineEdit_out_24_value_820.setValidator(self.valdat_24())

        self.ui.lineEdit_out_pvi_value_5.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_pvi_value_25.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_pvi_value_50.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_pvi_value_75.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_pvi_value_95.setValidator(self.validat_in_out())

        # Чекбокс ПВИ
        self.ui.checkBox_pvi.stateChanged.connect(lambda check=self.ui.checkBox_pvi.isChecked(): self.select_pvi(check))

        # Установка входных значений выхода
        self.ui.lineEdit_in_start_value.textChanged.connect(self.out_irt_in)
        self.ui.lineEdit_in_end_value.textChanged.connect(self.out_irt_in)

    def validat_in_out(self):
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,5}(?:[\.,]\d{,3})?$"))

    def validat_para(self):
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,6}(?:[\.,]\d{,3})?$"))

    def valdat_24(self):
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,2}(?:[\.,]\d{,3})?$"))

    def select_pvi(self, pvi_on):
        """ Включение группы ПВИ """
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
        """ Расчет требуемых входных """
        try:
            in_start = self.ui.lineEdit_in_start_value.text().replace(',', '.')
            in_end = self.ui.lineEdit_in_end_value.text().replace(',', '.')

            values = []
            for i in (0.05, 0.25, 0.5, 0.75, 0.95):
                values.append(str((abs(float(in_start)) + abs(float(in_end))) * i))

            self.ui.lineEdit_out_irt_in_5.setText(values[0])
            self.ui.lineEdit_out_irt_in_25.setText(values[1])
            self.ui.lineEdit_out_irt_in_50.setText(values[2])
            self.ui.lineEdit_out_irt_in_75.setText(values[3])
            self.ui.lineEdit_out_irt_in_95.setText(values[4])
        except:
            pass


app = QtWidgets.QApplication([])
application = window()
application.setWindowTitle("Создание протокола калибровки ИРТ 59хх")
application.show()

sys.exit(app.exec())
