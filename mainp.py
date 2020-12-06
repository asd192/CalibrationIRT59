import sys, os, configparser

from PyQt5 import QtWidgets, QtCore, QtGui
from PyQt5.QtWidgets import QMessageBox, QTableWidgetItem
from main import Ui_MainWindow


class Window(QtWidgets.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # Загрузка настроек
        self.load_param()

        # Заполнение comboBox-ов
        self.comboBox_load(0)  # средства калибровки
        self.comboBox_load(1)  # выходные величины
        self.comboBox_load(2)  # сдал
        self.comboBox_load(3)  # принял

        # Валидация полей Вход, Выход, Шкала ПВИ параметров прибора
        self.ui.lineEdit_in_start_value.setValidator(self.validat_param())
        self.ui.lineEdit_in_end_value.setValidator(self.validat_param())
        self.ui.lineEdit_out_start_value.setValidator(self.validat_param())
        self.ui.lineEdit_out_end_value.setValidator(self.validat_param())

        self.ui.lineEdit_pvi_scale_start.setValidator(self.validat_param())
        self.ui.lineEdit_pvi_scale_start.setValidator(self.validat_param())

        # Валиддация показаний Выход ИРТ, Выход ПВИ
        self.ui.lineEdit_out_irt_value_5.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_irt_value_25.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_irt_value_50.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_irt_value_75.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_irt_value_95.setValidator(self.validat_in_out())

        self.ui.lineEdit_out_pvi_value_5.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_pvi_value_25.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_pvi_value_50.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_pvi_value_75.setValidator(self.validat_in_out())
        self.ui.lineEdit_out_pvi_value_95.setValidator(self.validat_in_out())

        # Валидация полей Выход 24В
        self.ui.lineEdit_out_24_value_0.setValidator(self.valdat_24())
        self.ui.lineEdit_out_24_value_820.setValidator(self.valdat_24())

        # Чекбокс ПВИ
        self.ui.checkBox_pvi.stateChanged.connect(lambda check=self.ui.checkBox_pvi.isChecked(): self.select_pvi(check))

        # TODO рассчет погрешностей
        # Установка входных значений выходов и рассчет допустимых погрешностей
        self.ui.lineEdit_in_start_value.textChanged.connect(self.out_irt_in)
        self.ui.lineEdit_in_end_value.textChanged.connect(self.out_irt_in)

        self.ui.lineEdit_pvi_scale_start.textChanged.connect(self.out_pvi_in)
        self.ui.lineEdit_pvi_scale_end.textChanged.connect(self.out_pvi_in)

        # Установка параметров входа
        self.ui.comboBox_in_signal_type.activated.connect(self.parametr_in_signal)

        # R ПВИ
        self.ui.comboBox_pvi_out.activated.connect(self.parametr_pvi_out_signal)

        # Сохранение настроек
        self.ui.pushButton_save_custom.clicked.connect(self.save_param)

        # Сохранение файла конфигурации прибора
        self.ui.pushButton_save_config.clicked.connect(self.save_config_file)
        self.ui.action_save_config.triggered.connect(self.save_config_file)

        # О программе
        self.ui.action_about.triggered.connect(self.about)

    def parametr_pvi_out_signal(self):
        """ Определяет требуемое R ПВИ """
        r_pvi = {
            "0-5": "1.8 кОм",
            "0-20": "470 Ом",
            "4-20": "470 Ом"
        }

        r_pvi_key = self.ui.comboBox_pvi_out.currentText()
        if r_pvi_key in r_pvi.keys():
            self.ui.label_pvi_out_r.setText(f"R={str(r_pvi[r_pvi_key])}±5%")

    def parametr_in_signal(self):
        """ Устанавливает параметры прибора(входные значения) """
        sensor_type = {
            "ТП-ХК": (-50, 600),
            "ТП-ХА": (-50, 1300),
            "ТСМ-50М": (-50, 200),
            "ТСМ-100М": (-50, 200),
            "ТСМ-50П": (-50, 600),
            "ТСМ-100П": (-50, 600),
        }

        sensor_type_key = self.ui.comboBox_in_signal_type.currentText()
        if sensor_type_key in sensor_type.keys():
            self.ui.lineEdit_in_start_value.setText(str(sensor_type.get(sensor_type_key)[0]))
            self.ui.lineEdit_in_end_value.setText(str(sensor_type.get(sensor_type_key)[1]))

            self.ui.lineEdit_out_start_value.setText(str(sensor_type.get(sensor_type_key)[0]))
            self.ui.lineEdit_out_end_value.setText(str(sensor_type.get(sensor_type_key)[1]))

        else:
            self.ui.lineEdit_in_start_value.setText("")
            self.ui.lineEdit_in_end_value.setText("")

            self.ui.lineEdit_out_start_value.setText("")
            self.ui.lineEdit_out_end_value.setText("")

    def validat_param(self):
        """ Валидация полей Вход, Выход, Шкала ПВИ параметров прибора """
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,6}(?:[\.,]\d{,3})?$"))

    def validat_in_out(self):
        """ Валиддация показаний Выход ИРТ, Выход ПВИ """
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,5}(?:[\.,]\d{,3})?$"))

    def valdat_24(self):
        """ Валидация полей Выход 24В """
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,2}(?:[\.,]\d{,3})?$"))

    def comboBox_load(self, col=0):
        """ Заполнение ComboBox из tableWidget """
        _translate = QtCore.QCoreApplication.translate

        table = self.ui.tableWidget_param
        container = ['']

        for row in range(0, table.rowCount()):
            try:
                device = table.item(row, col).text()
                container.append(device)
            except:
                pass

        if col == 0:
            for n, i in enumerate(container):
                self.ui.comboBox_calibr_name.addItem("")
                self.ui.comboBox_calibr_name.setItemText(n, _translate("MainWindow", i))

        if col == 1:
            for n, i in enumerate(container):
                self.ui.comboBox_out_signal_type.addItem("")
                self.ui.comboBox_out_signal_type.setItemText(n, _translate("MainWindow", i))

        if col == 2:
            for n, i in enumerate(container):
                self.ui.comboBox_passed.addItem("")
                self.ui.comboBox_passed.setItemText(n, _translate("MainWindow", i))

        if col == 3:
            for n, i in enumerate(container):
                self.ui.comboBox_adopted.addItem("")
                self.ui.comboBox_adopted.setItemText(n, _translate("MainWindow", i))

        container.clear()

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
        values = ['']
        try:
            in_start = self.ui.lineEdit_in_start_value.text().replace(',', '.')
            in_end = self.ui.lineEdit_in_end_value.text().replace(',', '.')

            for i in (0.05, 0.25, 0.5, 0.75, 0.95):
                values.append(str((abs(float(in_start)) + abs(float(in_end))) * i))

            self.ui.lineEdit_out_irt_in_5.setText(values[1])
            self.ui.lineEdit_out_irt_in_25.setText(values[2])
            self.ui.lineEdit_out_irt_in_50.setText(values[3])
            self.ui.lineEdit_out_irt_in_75.setText(values[4])
            self.ui.lineEdit_out_irt_in_95.setText(values[5])
        except:
            self.ui.lineEdit_out_irt_in_5.setText(values[0])
            self.ui.lineEdit_out_irt_in_25.setText(values[0])
            self.ui.lineEdit_out_irt_in_50.setText(values[0])
            self.ui.lineEdit_out_irt_in_75.setText(values[0])
            self.ui.lineEdit_out_irt_in_95.setText(values[0])

    def out_pvi_in(self):
        """ Расчет требуемых входных ПВИ"""
        values = ['']
        try:
            in_start = self.ui.lineEdit_pvi_scale_start.text().replace(',', '.')
            in_end = self.ui.lineEdit_pvi_scale_end.text().replace(',', '.')

            for i in (0.05, 0.25, 0.5, 0.75, 0.95):
                values.append(str((abs(float(in_start)) + abs(float(in_end))) * i))

            self.ui.lineEdit_out_pvi_in_5.setText(values[1])
            self.ui.lineEdit_out_pvi_in_25.setText(values[2])
            self.ui.lineEdit_out_pvi_in_50.setText(values[3])
            self.ui.lineEdit_out_pvi_in_75.setText(values[4])
            self.ui.lineEdit_out_pvi_in_95.setText(values[5])
        except:
            self.ui.lineEdit_out_pvi_in_5.setText(values[0])
            self.ui.lineEdit_out_pvi_in_25.setText(values[0])
            self.ui.lineEdit_out_pvi_in_50.setText(values[0])
            self.ui.lineEdit_out_pvi_in_75.setText(values[0])
            self.ui.lineEdit_out_pvi_in_95.setText(values[0])

    def load_param(self):
        """ загрузка параметров """
        try:
            config = configparser.ConfigParser()
            config.read("parameters.ini")
            table = self.ui.tableWidget_param

            for column in range(0, table.columnCount()):
                column_name = table.horizontalHeaderItem(column).text()
                param = config.items(column_name)

                for row in range(0, len(param)):
                    table.setItem(row, column, QTableWidgetItem(param[row][1]))
        except Exception as exeption:
            QMessageBox.critical(self, "Ошибка",
                                 f"Не удалось загрузить параметры из файла <parameters.ini>. Ошибка - {type(exeption).__name__}",
                                 QMessageBox.Ok)

    def save_param(self):
        """ Сохранение параметров """
        table = self.ui.tableWidget_param
        config = configparser.ConfigParser()

        for column in range(0, table.columnCount()):
            column_name = table.horizontalHeaderItem(column).text()
            config.add_section(column_name)

            for row in range(0, table.rowCount()):
                try:
                    value = table.item(row, column).text()
                    config.set(column_name, str(row), value)
                except:
                    pass

        try:
            with open("parameters.ini", "w", "utf8") as config_file:
                config.write(config_file)
            QMessageBox.information(self, "Параметры сохранены",
                                    "Необходимо перезагрузить программу для обновления параметров.", QMessageBox.Ok)
        except Exception as exeption:
            QMessageBox.critical(self, "Ошибка записи",
                                  f"Не удалось сохранить параметры. Ошибка - {type(exeption).__name__}", QMessageBox.Ok)

    def save_config_file(self):
        """ Сохраняет файл конфигурации прибора """
        file_path = "configurations"
        file_name = self.ui.lineEdit_parametr_position.text()

        config = configparser.ConfigParser()
        config.add_section("Средство калибровки")
        config.set("Средство калибровки", "Калибратор", self.ui.comboBox_calibr_name.currentText())

        config.add_section("Параметры прибора")
        config.set("Параметры прибора", "Тип", self.ui.comboBox_parametr_type.currentText())
        config.set("Параметры прибора", "Номер", self.ui.lineEdit_parametr_number.text())
        config.set("Параметры прибора", "Год выпуска", self.ui.comboBox_parametr_year.currentText())
        config.set("Параметры прибора", "Позиция", self.ui.lineEdit_parametr_position.text())
        config.set("Параметры прибора", "Тип входа", self.ui.comboBox_in_signal_type.currentText())
        config.set("Параметры прибора", "Вход начало шкалы", self.ui.lineEdit_in_start_value.text())
        config.set("Параметры прибора", "Вход конец шкалы", self.ui.lineEdit_in_end_value.text())
        config.set("Параметры прибора", "Тип выхода", self.ui.comboBox_in_signal_type.currentText())
        config.set("Параметры прибора", "Выход начало шкалы", self.ui.lineEdit_out_start_value.text())
        config.set("Параметры прибора", "Выход конец шкалы", self.ui.lineEdit_out_end_value.text())
        config.set("Параметры прибора", "Наличие ПВИ", str(self.ui.checkBox_pvi.isChecked()))
        config.set("Параметры прибора", "ПВИ начало шкалы", self.ui.lineEdit_pvi_scale_start.text())
        config.set("Параметры прибора", "ПВИ конец шкалы", self.ui.lineEdit_pvi_scale_end.text())
        config.set("Параметры прибора", "ПВИ тип выхода", self.ui.comboBox_pvi_out.currentText())

        config.add_section("Выход ИРТ")
        config.set("Выход ИРТ", "Выход 5", self.ui.lineEdit_out_irt_value_5.text())
        config.set("Выход ИРТ", "Выход 25", self.ui.lineEdit_out_irt_value_25.text())
        config.set("Выход ИРТ", "Выход 50", self.ui.lineEdit_out_irt_value_50.text())
        config.set("Выход ИРТ", "Выход 75", self.ui.lineEdit_out_irt_value_75.text())
        config.set("Выход ИРТ", "Выход 95", self.ui.lineEdit_out_irt_value_95.text())

        config.add_section("Выход 24В")
        config.set("Выход 24В", "Выход R0", self.ui.lineEdit_out_24_value_0.text())
        config.set("Выход 24В", "Выход R820", self.ui.lineEdit_out_24_value_820.text())

        config.add_section("Выход ПВИ")
        config.set("Выход ПВИ", "Выход ПВИ 5", self.ui.lineEdit_out_pvi_value_5.text())
        config.set("Выход ПВИ", "Выход ПВИ 25", self.ui.lineEdit_out_pvi_value_25.text())
        config.set("Выход ПВИ", "Выход ПВИ 50", self.ui.lineEdit_out_pvi_value_50.text())
        config.set("Выход ПВИ", "Выход ПВИ 75", self.ui.lineEdit_out_pvi_value_75.text())
        config.set("Выход ПВИ", "Выход ПВИ 95", self.ui.lineEdit_out_pvi_value_95.text())

        config.add_section("Сдал/Принял/Дата")
        config.set("Сдал/Принял/Дата", "Сдал", self.ui.comboBox_passed.currentText())
        config.set("Сдал/Принял/Дата", "Принял", self.ui.comboBox_adopted.currentText())
        print(self.ui.dateEdit_date_calibration.dateTime().toString('dd-MM-yyyy'))
        config.set("Сдал/Принял/Дата", "Дата калибровки(ГГГГ.ММ.ДД)", self.ui.dateEdit_date_calibration.dateTime().toString('yyyy.MM.dd'))

        try:
            if not os.path.isdir(file_path):
                os.mkdir(file_path)

            with open(f"{file_path}/{file_name}.clbr59", "w") as config_file:
                config.write(config_file)

            QMessageBox.information(self, "Сохранено", f"Конфигурация {file_name} успешно сохранена", QMessageBox.Ok)

        except Exception as exeption:
            QMessageBox.critical(self, "Ошибка записи",
                                 f"Не удалось сохранить параметры. Ошибка - {type(exeption).__name__}", QMessageBox.Ok)

    def about(self):
        QtWidgets.QMessageBox.aboutQt(application, title="О программе")

app = QtWidgets.QApplication([])
application = Window()
application.setWindowTitle("Создание протокола калибровки ИРТ 59хх")
application.show()

sys.exit(app.exec())
