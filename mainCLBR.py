"""
Главное окно программы
"""

import sys, os, configparser, decimal, subprocess, shutil, openpyxl, random

from PyQt5 import QtWidgets, QtCore, QtGui

from webbrowser import open_new as wb_open_new

from main import Ui_MainWindow
from slSettings import ClbrSettings


class ClbrMain(QtWidgets.QMainWindow):
    verification_error = {
        'acceptance_conditions_t': True,
        'acceptance_conditions_f': True,
        'acceptance_conditions_p': True,
        'acceptance_error_24_0': True,
        'acceptance_error_24_820': True,
        'acceptance_error_irt_5': True,
        'acceptance_error_irt_25': True,
        'acceptance_error_irt_50': True,
        'acceptance_error_irt_75': True,
        'acceptance_error_irt_95': True,
        'acceptance_error_pvi_5': True,
        'acceptance_error_pvi_25': True,
        'acceptance_error_pvi_50': True,
        'acceptance_error_pvi_75': True,
        'acceptance_error_pvi_95': True,
    }

    # Допуски для *.xlsx
    permissible_inaccuracy_irt = '0'
    permissible_inaccuracy_pvi = '0'
    permissible_inaccuracy_24v = '2'

    def __init__(self):
        super(ClbrMain, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        self.ui.progressBar.hide()

        # Загрузка параметров
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

        # Валидация полей Условия
        self.ui.lineEdit_t.setValidator(self.validat_conditions())
        self.ui.lineEdit_f.setValidator(self.validat_conditions())
        self.ui.lineEdit_p.setValidator(self.validat_conditions())

        # Замена ',' -> '.'
        widgets_repl = (self.ui.lineEdit_t, self.ui.lineEdit_f, self.ui.lineEdit_p,
                        self.ui.lineEdit_in_start_value, self.ui.lineEdit_in_end_value,
                        self.ui.lineEdit_out_start_value, self.ui.lineEdit_out_end_value,
                        self.ui.lineEdit_pvi_scale_start, self.ui.lineEdit_pvi_scale_end,
                        self.ui.lineEdit_out_irt_value_5, self.ui.lineEdit_out_irt_value_25,
                        self.ui.lineEdit_out_irt_value_50, self.ui.lineEdit_out_irt_value_75,
                        self.ui.lineEdit_out_irt_value_95, self.ui.lineEdit_out_24_value_0,
                        self.ui.lineEdit_out_24_value_820, self.ui.lineEdit_out_pvi_value_5,
                        self.ui.lineEdit_out_pvi_value_25, self.ui.lineEdit_out_pvi_value_50,
                        self.ui.lineEdit_out_pvi_value_75, self.ui.lineEdit_out_pvi_value_95)

        for wr in widgets_repl:
            self.repl(wr)

        # Дата сегодня
        self.ui.dateEdit_date_calibration.setDate(QtCore.QDate.currentDate())

        # Валидация полей Выход 24В
        self.ui.lineEdit_out_24_value_0.setValidator(self.valdat_24())
        self.ui.lineEdit_out_24_value_820.setValidator(self.valdat_24())

        # Группа ПВИ(чекбокс)
        self.ui.groupBox_out_pvi.setEnabled(False)
        self.ui.checkBox_pvi.stateChanged.connect(lambda check=self.ui.checkBox_pvi.isChecked(): self.select_pvi(check))

        # Установка входных значений выходов ИРТ
        self.ui.lineEdit_in_start_value.textChanged.connect(self.out_irt_in)
        self.ui.lineEdit_in_end_value.textChanged.connect(self.out_irt_in)

        # Установка выходных значений выходов ИРТ
        self.ui.lineEdit_out_start_value.textChanged.connect(self.out_irt_out)
        self.ui.lineEdit_out_end_value.textChanged.connect(self.out_irt_out)

        # Установка входных значений выходов ПВИ
        self.ui.lineEdit_pvi_scale_start.textChanged.connect(self.out_pvi_in)
        self.ui.lineEdit_pvi_scale_end.textChanged.connect(self.out_pvi_in)

        # Установка выходных значений выходов, R и допуска ПВИ
        self.ui.comboBox_pvi_out.currentTextChanged.connect(self.out_pvi_out)

        # Пересчет допуска ПВИ
        self.ui.lineEdit_in_start_value.textChanged.connect(self.out_pvi_out)
        self.ui.lineEdit_in_end_value.textChanged.connect(self.out_pvi_out)
        self.ui.lineEdit_pvi_scale_start.textChanged.connect(self.out_pvi_out)
        self.ui.lineEdit_pvi_scale_end.textChanged.connect(self.out_pvi_out)

        # Установка допуска ПВИ
        self.ui.lineEdit_out_start_value.textChanged.connect(self.acceptance_irt)
        self.ui.lineEdit_out_end_value.textChanged.connect(self.acceptance_irt)
        self.ui.comboBox_out_signal_type.currentTextChanged.connect(self.acceptance_irt)

        # Установка параметров входа
        self.ui.comboBox_in_signal_type.activated.connect(self.parametr_in_signal)

        # Цветовая индикация условий калибровки
        self.ui.lineEdit_t.textChanged.connect(self.acceptance_conditions_t)
        self.ui.lineEdit_f.textChanged.connect(self.acceptance_conditions_f)
        self.ui.lineEdit_p.textChanged.connect(self.acceptance_conditions_p)

        # Цветовая индикация допуска ИРТ
        self.ui.lineEdit_out_irt_value_5.textChanged.connect(
            lambda irt_value=self.ui.lineEdit_out_irt_value_5.text(): self.acceptance_error_irt(irt_value, 5))
        self.ui.lineEdit_out_irt_value_25.textChanged.connect(
            lambda irt_value=self.ui.lineEdit_out_irt_value_25.text(): self.acceptance_error_irt(irt_value, 25))
        self.ui.lineEdit_out_irt_value_50.textChanged.connect(
            lambda irt_value=self.ui.lineEdit_out_irt_value_50.text(): self.acceptance_error_irt(irt_value, 50))
        self.ui.lineEdit_out_irt_value_75.textChanged.connect(
            lambda irt_value=self.ui.lineEdit_out_irt_value_75.text(): self.acceptance_error_irt(irt_value, 75))
        self.ui.lineEdit_out_irt_value_95.textChanged.connect(
            lambda irt_value=self.ui.lineEdit_out_irt_value_95.text(): self.acceptance_error_irt(irt_value, 95))

        # Цветовая индикация допуска ПВИ
        self.ui.lineEdit_out_pvi_value_5.textChanged.connect(
            lambda pvi_value=self.ui.lineEdit_out_pvi_value_5.text(): self.acceptance_error_pvi(pvi_value, 5))
        self.ui.lineEdit_out_pvi_value_25.textChanged.connect(
            lambda pvi_value=self.ui.lineEdit_out_pvi_value_25.text(): self.acceptance_error_pvi(pvi_value, 25))
        self.ui.lineEdit_out_pvi_value_50.textChanged.connect(
            lambda pvi_value=self.ui.lineEdit_out_pvi_value_50.text(): self.acceptance_error_pvi(pvi_value, 50))
        self.ui.lineEdit_out_pvi_value_75.textChanged.connect(
            lambda pvi_value=self.ui.lineEdit_out_pvi_value_75.text(): self.acceptance_error_pvi(pvi_value, 75))
        self.ui.lineEdit_out_pvi_value_95.textChanged.connect(
            lambda pvi_value=self.ui.lineEdit_out_pvi_value_95.text(): self.acceptance_error_pvi(pvi_value, 95))

        # Цветовая индикация допуска 24В
        self.ui.lineEdit_out_24_value_0.textChanged.connect(self.acceptance_error_24_0)
        self.ui.lineEdit_out_24_value_820.textChanged.connect(self.acceptance_error_24_820)

        # Сохранение настроек
        self.ui.pushButton_save_custom.clicked.connect(self.save_param)

        # Очистить все поля(Новый)
        self.ui.action_empty.triggered.connect(self.cal_clear)

        # Сохранение файла конфигурации прибора
        self.ui.pushButton_save_config.clicked.connect(self.save_config_file)
        self.ui.action_save_config.triggered.connect(self.save_config_file)

        # Создание протокола
        self.ui.pushButton_create.clicked.connect(self.create_protocol_verify)

        # Настройки программы
        self.ui.action_settings.triggered.connect(self.open_settings)

        # Загрузка файла конфигурации прибора
        self.ui.action_load_config.triggered.connect(self.load_config_file)

        if sys.platform == "win32":
            # Help
            self.ui.action_help.setVisible(True)
            self.ui.action_help.triggered.connect(self.ui_help)
            # Send error
            self.ui.action_error.triggered.connect(self.help_error)

        # О программе
        self.ui.action_about.triggered.connect(self.about)

        # Выход из программы
        self.ui.action_exit.triggered.connect(self.exit)

        # Тестовый режим (заполняет поля позитивными значениями) Ctrl + T
        self.ui.action_testing.triggered.connect(self.positive_testing)

    def repl(self, p):
        p.textChanged.connect(lambda v=p.text(): p.setText(v.replace(",", ".")))

    def parametr_in_signal(self):
        """ Устанавливает параметры прибора(входные значения) """
        sensor_type = {
            "ТП-ХК": (-50, 600),
            "ТП-ХА": (-50, 1300),
            "0-5 мА": (0, 5),
            "0-20 мА": (0, 20),
            "4-20 мА": (4, 20),
            "ТСМ-50М": (-50, 200),
            "ТСМ-100М": (-50, 200),
            "ТСМ-50П": (-50, 600),
            "ТСМ-100П": (-50, 600),
        }

        sensor_type_key = self.ui.comboBox_in_signal_type.currentText()
        if sensor_type_key in sensor_type.keys():
            self.ui.lineEdit_in_start_value.setText(str(sensor_type.get(sensor_type_key)[0]))
            self.ui.lineEdit_in_end_value.setText(str(sensor_type.get(sensor_type_key)[1]))

            if "Т" in sensor_type_key:
                self.ui.lineEdit_out_start_value.setText(str(sensor_type.get(sensor_type_key)[0]))
                self.ui.lineEdit_out_end_value.setText(str(sensor_type.get(sensor_type_key)[1]))
            else:
                self.ui.lineEdit_out_start_value.setText("")
                self.ui.lineEdit_out_end_value.setText("")

        else:
            self.ui.lineEdit_in_start_value.setText("")
            self.ui.lineEdit_in_end_value.setText("")

            self.ui.lineEdit_out_start_value.setText("")
            self.ui.lineEdit_out_end_value.setText("")

        # Допуск ИРТ
        self.acceptance_irt()
        # Допуск ПВИ
        self.out_pvi_out()

    def validat_param(self):
        """ Валидация полей Вход, Выход, Шкала ПВИ параметров прибора """
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,6}(?:[\.,]\d{,3})?$"))

    def validat_in_out(self):
        """ Валиддация показаний Выход ИРТ, Выход ПВИ """
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,5}(?:[\.,]\d{,3})?$"))

    def valdat_24(self):
        """ Валидация полей Выход 24В """
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,2}(?:[\.,]\d{,3})?$"))

    def validat_conditions(self):
        return QtGui.QRegExpValidator(QtCore.QRegExp("^[+-]?\d{,3}(?:[\.,]\d{,2})?$"))

    def comboBox_load(self, col):
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

        field_comboBox = None
        if col == 0:
            field_comboBox = self.ui.comboBox_calibr_name
        if col == 1:
            field_comboBox = self.ui.comboBox_out_signal_type
        if col == 2:
            field_comboBox = self.ui.comboBox_passed
        if col == 3:
            field_comboBox = self.ui.comboBox_adopted

        for n, i in enumerate(container):
            field_comboBox.addItem("")
            field_comboBox.setItemText(n, _translate("MainWindow", i))

        container.clear()

    def select_pvi(self, pvi_on):
        """ Включение группы ПВИ """
        self.ui.groupBox_out_pvi.setEnabled(pvi_on)
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
            # Установка входных
            in_start = float(self.ui.lineEdit_in_start_value.text())
            in_end = float(self.ui.lineEdit_in_end_value.text())

            for i in (0.05, 0.25, 0.5, 0.75, 0.95):
                values.append(str(round((float(in_end) - float(in_start)) * i + float(in_start), 3)))

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

    def out_irt_out(self):
        """ Расчет требуемых выходных """
        values = ['']
        try:
            # Установка выходных
            in_start = float(self.ui.lineEdit_out_start_value.text())
            in_end = float(self.ui.lineEdit_out_end_value.text())

            for i in (0.05, 0.25, 0.5, 0.75, 0.95):
                values.append(str(round((float(in_end) - float(in_start)) * i + float(in_start), 3)))

            self.ui.lineEdit_out_irt_output_5.setText(values[1])
            self.ui.lineEdit_out_irt_output_25.setText(values[2])
            self.ui.lineEdit_out_irt_output_50.setText(values[3])
            self.ui.lineEdit_out_irt_output_75.setText(values[4])
            self.ui.lineEdit_out_irt_output_95.setText(values[5])
        except:
            self.ui.lineEdit_out_irt_output_5.setText(values[0])
            self.ui.lineEdit_out_irt_output_25.setText(values[0])
            self.ui.lineEdit_out_irt_output_50.setText(values[0])
            self.ui.lineEdit_out_irt_output_75.setText(values[0])
            self.ui.lineEdit_out_irt_output_95.setText(values[0])

    def out_pvi_in(self):
        """ Расчет требуемых входных ПВИ"""
        values = ['']
        try:
            in_start = float(self.ui.lineEdit_pvi_scale_start.text())
            in_end = float(self.ui.lineEdit_pvi_scale_end.text())

            for i in (0.05, 0.25, 0.5, 0.75, 0.95):
                values.append(str(round((in_end - in_start) * i + in_start, 3)))

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

    def out_pvi_out(self, acceptance_error=False):
        """ Расчет требуемых выходных ПВИ"""
        values = ['']
        try:
            type_out = self.ui.comboBox_pvi_out.currentText().split('-')

            in_start = int(type_out[0])
            in_end = int(type_out[1])

            for i in (0.05, 0.25, 0.5, 0.75, 0.95):
                values.append(str((in_end - in_start) * i + in_start))

            self.ui.lineEdit_out_pvi_output_5.setText(values[1])
            self.ui.lineEdit_out_pvi_output_25.setText(values[2])
            self.ui.lineEdit_out_pvi_output_50.setText(values[3])
            self.ui.lineEdit_out_pvi_output_75.setText(values[4])
            self.ui.lineEdit_out_pvi_output_95.setText(values[5])
        except:
            self.ui.lineEdit_out_pvi_output_5.setText(values[0])
            self.ui.lineEdit_out_pvi_output_25.setText(values[0])
            self.ui.lineEdit_out_pvi_output_50.setText(values[0])
            self.ui.lineEdit_out_pvi_output_75.setText(values[0])
            self.ui.lineEdit_out_pvi_output_95.setText(values[0])

        # R ПВИ
        r_pvi = {
            "0-5": "1.8 кОм",
            "0-20": "470 Ом",
            "4-20": "470 Ом"
        }

        r_pvi_key = self.ui.comboBox_pvi_out.currentText()
        if r_pvi_key in r_pvi.keys():
            self.ui.label_pvi_out_r.setText(f"R={str(r_pvi[r_pvi_key])}±5%")

        # Допуск ПВИ
        K = 0.2  # допускаемая основная погрешность
        K_percent = 0
        try:
            irt_start = float(self.ui.lineEdit_pvi_scale_start.text())
            irt_end = float(self.ui.lineEdit_pvi_scale_end.text())

            pvi_start = float(self.ui.comboBox_pvi_out.currentText().split("-")[0])
            pvi_end = float(self.ui.comboBox_pvi_out.currentText().split("-")[1])

            # _K - погрешность
            _K = 0.2
            # k - коэффициент, равный отношению диапазона измерений к диапазону преобразования ПВИ
            k = round((pvi_end - pvi_start) / (irt_end - irt_start), 3)
            # Y - предел основной приведенной погрешности ПВИ
            Y = self.acceptance_irt(True)

            K_percent = round(k * Y + _K, 3)

            K = abs(round(K_percent / 100 * (pvi_start - pvi_end), 3))

            accept = f"Допуск ±({k}*{Y}+{_K})% -> {K}мА"
        except:
            accept = f"Допуск ±(k γ0+0.2)%"

        self.ui.label_acceptance_error_pvi.setText(accept)
        ClbrMain.permissible_inaccuracy_pvi = K_percent

        # если запрошен допуск ПВИ
        if acceptance_error:
            return round(K, 3)

    def convert_decimal(self, num: str) -> decimal:
        """ Преобразует в Decimal для определения единицы последнего разряда """
        try:
            if float(num) or float(num) == 0:
                return decimal.Decimal(num)
        except:
            pass

    def acceptance_irt(self, acceptance_error=False):
        """ Устанавливает допуск ИРТ """
        out_start = self.convert_decimal(self.ui.lineEdit_out_start_value.text())
        out_end = self.convert_decimal(self.ui.lineEdit_out_end_value.text())

        one_unit_last_number = 0
        try:
            range_scale = str(out_end - out_start)

            dig = abs(range_scale.find('.') - len(range_scale)) - 1 if '.' in range_scale else 0
            if dig == 0:
                one_unit_last_number = 1
            if dig == 1:
                one_unit_last_number = 0.1
            if dig == 2:
                one_unit_last_number = 0.01
            if dig == 3:
                one_unit_last_number = 0.001

            # одна единица последнего разряда, выраженная в процентах от диапазона измерений
            one_unit_last_number = abs(round(one_unit_last_number / float(range_scale) * 100, 3))
        except:
            pass

        in_signal_type = self.ui.comboBox_in_signal_type.currentText()
        if 'мА' in in_signal_type:
            _K = 0.2
        elif 'ТП' in in_signal_type:
            _K = 0.5
        elif 'ТСМ' in in_signal_type:
            _K = 0.2
        else:
            _K = "0.0"
            one_unit_last_number = "*"

        K = _K + one_unit_last_number

        # Допуск в установленных единицах
        acceptance = '*'
        try:
            acceptance = round(K / 100 * (float(out_end) - float(out_start)), 3)
        except:
            pass

        signal_type = self.ui.comboBox_out_signal_type.currentText()
        in_signal_text = f"Допуск ±({_K} + {one_unit_last_number})% -> {acceptance}{signal_type}"
        self.ui.label_acceptance_error_irt.setText(in_signal_text)

        ClbrMain.permissible_inaccuracy_irt = K

        # если запрошен допуск ИРТ
        if acceptance_error:
            return acceptance

    def acceptance_conditions_t(self):
        # температура окружающего воздуха
        t = 0 if self.ui.lineEdit_t.text() == '' else float(self.ui.lineEdit_t.text())
        if 15 <= t <= 25:
            color = u"color: black"
            ClbrMain.verification_error['acceptance_conditions_t'] = True
        else:
            color = u"color: red"
            ClbrMain.verification_error['acceptance_conditions_t'] = False
        self.ui.lineEdit_t.setStyleSheet(color)

    def acceptance_conditions_f(self):
        # относительная влажность воздуха
        f = 0 if self.ui.lineEdit_f.text() == '' else float(self.ui.lineEdit_f.text())
        if 30 <= f <= 80:
            color = u"color: black"
            ClbrMain.verification_error['acceptance_conditions_f'] = True
        else:
            color = u"color: red"
            ClbrMain.verification_error['acceptance_conditions_f'] = False
        self.ui.lineEdit_f.setStyleSheet(color)

    def acceptance_conditions_p(self):
        # атмосферное давление
        p = 0 if self.ui.lineEdit_p.text() == '' else float(self.ui.lineEdit_p.text())
        if 84.0 <= p <= 106.7:
            color = u"color: black"
            ClbrMain.verification_error['acceptance_conditions_p'] = True
        else:
            color = u"color: red"
            ClbrMain.verification_error['acceptance_conditions_p'] = False
        self.ui.lineEdit_p.setStyleSheet(color)

    def error_irt_5(self):
        Ai = self.ui.lineEdit_out_irt_value_5.text()
        Ad = self.ui.lineEdit_out_irt_output_5.text()
        self.ui.lineEdit_out_irt_value_5.setStyleSheet(self.acceptance_error_irt(Ai, Ad))

    def error_irt_25(self):
        Ai = self.ui.lineEdit_out_irt_value_25.text()
        Ad = self.ui.lineEdit_out_irt_output_25.text()
        self.ui.lineEdit_out_irt_value_25.setStyleSheet(self.acceptance_error_irt(Ai, Ad))

    def error_irt_50(self):
        Ai = self.ui.lineEdit_out_irt_value_50.text()
        Ad = self.ui.lineEdit_out_irt_output_50.text()
        self.ui.lineEdit_out_irt_value_50.setStyleSheet(self.acceptance_error_irt(Ai, Ad))

    def error_irt_75(self):
        Ai = self.ui.lineEdit_out_irt_value_75.text()
        Ad = self.ui.lineEdit_out_irt_output_75.text()
        self.ui.lineEdit_out_irt_value_75.setStyleSheet(self.acceptance_error_irt(Ai, Ad))

    def error_irt_95(self):
        Ai = self.ui.lineEdit_out_irt_value_95.text()
        Ad = self.ui.lineEdit_out_irt_output_95.text()
        self.ui.lineEdit_out_irt_value_95.setStyleSheet(self.acceptance_error_irt(Ai, Ad))

    def acceptance_error_irt(self, irt_value, percent):
        """ Расчет допусков ИРТ """
        acceptance_error = self.acceptance_irt(True)

        try:
            irt_value = float(irt_value)
            irt_output = 0

            if percent == 5:
                irt_output = float(self.ui.lineEdit_out_irt_output_5.text())
            if percent == 25:
                irt_output = float(self.ui.lineEdit_out_irt_output_25.text())
            if percent == 50:
                irt_output = float(self.ui.lineEdit_out_irt_output_50.text())
            if percent == 75:
                irt_output = float(self.ui.lineEdit_out_irt_output_75.text())
            if percent == 95:
                irt_output = float(self.ui.lineEdit_out_irt_output_95.text())

            if round(abs(irt_value - irt_output), 5) > acceptance_error:
                color = u"color: red"
                acceptance = False
            else:
                color = u"color: black"
                acceptance = True

            if percent == 5:
                self.ui.lineEdit_out_irt_value_5.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_irt_5'] = acceptance
            if percent == 25:
                self.ui.lineEdit_out_irt_value_25.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_irt_25'] = acceptance
            if percent == 50:
                self.ui.lineEdit_out_irt_value_50.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_irt_50'] = acceptance
            if percent == 75:
                self.ui.lineEdit_out_irt_value_75.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_irt_75'] = acceptance
            if percent == 95:
                self.ui.lineEdit_out_irt_value_95.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_irt_95'] = acceptance
        except ValueError:
            pass

    def acceptance_error_pvi(self, pvi_value, percent):
        """ Расчет допусков ПВИ """
        acceptance_error = self.out_pvi_out(True)

        try:
            pvi_value = float(pvi_value)
            pvi_output = 0

            if percent == 5:
                pvi_output = float(self.ui.lineEdit_out_pvi_output_5.text())
            if percent == 25:
                pvi_output = float(self.ui.lineEdit_out_pvi_output_25.text())
            if percent == 50:
                pvi_output = float(self.ui.lineEdit_out_pvi_output_50.text())
            if percent == 75:
                pvi_output = float(self.ui.lineEdit_out_pvi_output_75.text())
            if percent == 95:
                pvi_output = float(self.ui.lineEdit_out_pvi_output_95.text())

            if round(abs(pvi_value - pvi_output), 5) > acceptance_error:
                color = u"color: red"
                acceptance = False
            else:
                color = u"color: black"
                acceptance = True

            if percent == 5:
                self.ui.lineEdit_out_pvi_value_5.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_pvi_5'] = acceptance
            if percent == 25:
                self.ui.lineEdit_out_pvi_value_25.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_pvi_25'] = acceptance
            if percent == 50:
                self.ui.lineEdit_out_pvi_value_50.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_pvi_50'] = acceptance
            if percent == 75:
                self.ui.lineEdit_out_pvi_value_75.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_pvi_75'] = acceptance
            if percent == 95:
                self.ui.lineEdit_out_pvi_value_95.setStyleSheet(color)
                ClbrMain.verification_error['acceptance_error_pvi_95'] = acceptance
        except ValueError:
            pass

    def acceptance_error_24_0(self):
        try:
            Ai = abs(float(self.ui.lineEdit_out_24_value_0.text()))
            Ad = float(self.ui.lineEdit_out_24_in_0.text())
            if abs(round(Ad - Ai, 3)) > 0.48:
                color = u"color: red"
                acceptance = False
            else:
                color = u"color: black"
                acceptance = True

            self.ui.lineEdit_out_24_value_0.setStyleSheet(color)
            ClbrMain.verification_error['acceptance_error_24_0'] = acceptance
        except:
            pass

    def acceptance_error_24_820(self):
        try:
            Ai = abs(float(self.ui.lineEdit_out_24_value_820.text().replace(",", ".")))
            Ad = float(self.ui.lineEdit_out_24_in_820.text().replace(",", "."))
            if abs(round(Ad - Ai, 3)) > 0.48:
                color = u"color: red"
                acceptance = False
            else:
                color = u"color: black"
                acceptance = True

            self.ui.lineEdit_out_24_value_820.setStyleSheet(color)
            ClbrMain.verification_error['acceptance_error_24_820'] = acceptance
        except:
            pass

    def load_param(self):
        """ Загрузка параметров """
        try:
            config = configparser.ConfigParser()
            config.read("parameters.ini", encoding="utf-8")
            table = self.ui.tableWidget_param

            for column in range(0, table.columnCount()):
                column_name = table.horizontalHeaderItem(column).text()
                param = config.items(column_name)

                for row in range(0, len(param)):
                    table.setItem(row, column, QtWidgets.QTableWidgetItem(param[row][1]))
        except Exception as exeption:
            QtWidgets.QMessageBox.critical(self, "Ошибка",
                                           f"Не удалось загрузить параметры из файла <parameters.ini>.Ошибка - {type(exeption).__name__}",
                                           QtWidgets.QMessageBox.Ok)

    def save_param(self):
        """ Сохранение параметров """
        file_parameters = "parameters.ini"

        config = configparser.ConfigParser()
        config.read(file_parameters, encoding="utf-8")

        table = self.ui.tableWidget_param
        for column in range(0, table.columnCount()):
            column_name = table.horizontalHeaderItem(column).text()

            if config.has_section(column_name) == False:
                config.add_section(column_name)

            for row in range(0, table.rowCount()):
                try:
                    value = table.item(row, column).text()
                except:
                    value = ''
                config.set(column_name, str(row), value)

        try:
            with open(file_parameters, "w", encoding="utf-8") as config_file:
                config.write(config_file)
                QtWidgets.QMessageBox.information(self, "Параметры сохранены",
                                                  "Необходимо перезапустить программу для обновления параметров.",
                                                  QtWidgets.QMessageBox.Ok)
        except Exception as exeption:
            QtWidgets.QMessageBox.critical(self, "Ошибка записи",
                                           f"Не удалось сохранить параметры. Ошибка - {type(exeption).__name__}",
                                           QtWidgets.QMessageBox.Ok)

    def save_config_file(self):
        """ Сохраняет файл конфигурации прибора """
        file_path = "configurations"
        file_name = self.ui.lineEdit_parametr_position.text()

        exemplary = "Средство калибровки"
        environment = "Условия калибровки"
        parameters = "Параметры прибора"
        out_irt = "Выход ИРТ"
        out_24 = "Выход 24В"
        out_pvi = "Выход ПВИ"
        participants = "Сдал/Принял/Дата"

        config = configparser.ConfigParser()
        config.add_section(exemplary)
        config.set(exemplary, "Калибратор", self.ui.comboBox_calibr_name.currentText())

        config.add_section(environment)
        config.set(environment, "Температура", self.ui.lineEdit_t.text())
        config.set(environment, "Влажность", self.ui.lineEdit_f.text())
        config.set(environment, "Давление", self.ui.lineEdit_p.text())

        config.add_section(parameters)
        config.set(parameters, "Тип", self.ui.comboBox_parametr_type.currentText())
        config.set(parameters, "Номер", self.ui.lineEdit_parametr_number.text())
        config.set(parameters, "Год выпуска", self.ui.comboBox_parametr_year.currentText())
        config.set(parameters, "Позиция", self.ui.lineEdit_parametr_position.text())
        config.set(parameters, "Тип входа", self.ui.comboBox_in_signal_type.currentText())
        config.set(parameters, "Вход начало шкалы", self.ui.lineEdit_in_start_value.text())
        config.set(parameters, "Вход конец шкалы", self.ui.lineEdit_in_end_value.text())
        config.set(parameters, "Тип выхода", self.ui.comboBox_out_signal_type.currentText())
        config.set(parameters, "Выход начало шкалы", self.ui.lineEdit_out_start_value.text())
        config.set(parameters, "Выход конец шкалы", self.ui.lineEdit_out_end_value.text())
        config.set(parameters, "Наличие ПВИ", str(self.ui.checkBox_pvi.isChecked()))
        config.set(parameters, "ПВИ начало шкалы", self.ui.lineEdit_pvi_scale_start.text())
        config.set(parameters, "ПВИ конец шкалы", self.ui.lineEdit_pvi_scale_end.text())
        config.set(parameters, "ПВИ тип выхода", self.ui.comboBox_pvi_out.currentText())

        config.add_section(out_irt)
        config.set(out_irt, "Показания 5", self.ui.lineEdit_out_irt_value_5.text())
        config.set(out_irt, "Показания 25", self.ui.lineEdit_out_irt_value_25.text())
        config.set(out_irt, "Показания 50", self.ui.lineEdit_out_irt_value_50.text())
        config.set(out_irt, "Показания 75", self.ui.lineEdit_out_irt_value_75.text())
        config.set(out_irt, "Показания 95", self.ui.lineEdit_out_irt_value_95.text())

        config.add_section(out_24)
        config.set(out_24, "Выход R0", self.ui.lineEdit_out_24_value_0.text())
        config.set(out_24, "Выход R820", self.ui.lineEdit_out_24_value_820.text())

        config.add_section("Выход ПВИ")
        config.set(out_pvi, "Показания 5", self.ui.lineEdit_out_pvi_value_5.text())
        config.set(out_pvi, "Показания 25", self.ui.lineEdit_out_pvi_value_25.text())
        config.set(out_pvi, "Показания 50", self.ui.lineEdit_out_pvi_value_50.text())
        config.set(out_pvi, "Показания 75", self.ui.lineEdit_out_pvi_value_75.text())
        config.set(out_pvi, "Показания 95", self.ui.lineEdit_out_pvi_value_95.text())

        config.set(out_pvi, "Выход 5", self.ui.lineEdit_out_pvi_output_5.text())
        config.set(out_pvi, "Выход 25", self.ui.lineEdit_out_pvi_output_25.text())
        config.set(out_pvi, "Выход 50", self.ui.lineEdit_out_pvi_output_50.text())
        config.set(out_pvi, "Выход 75", self.ui.lineEdit_out_pvi_output_75.text())
        config.set(out_pvi, "Выход 95", self.ui.lineEdit_out_pvi_output_95.text())

        config.add_section(participants)
        config.set(participants, "Сдал", self.ui.comboBox_passed.currentText())
        config.set(participants, "Принял", self.ui.comboBox_adopted.currentText())
        config.set(participants, "дата калибровки(ДД.ММ.ГГГГ)",
                   self.ui.dateEdit_date_calibration.dateTime().toString('dd.MM.yyyy'))

        try:
            if not os.path.isdir(file_path):
                os.mkdir(file_path)

            with open(f"{file_path}/{file_name}.clbr59", "w", encoding="UTF-8") as config_file:
                config.write(config_file)

            QtWidgets.QMessageBox.information(self, "Сохранено",
                                              f"Конфигурация {file_name} успешно сохранена", QtWidgets.QMessageBox.Ok)

        except Exception as exeption:
            QtWidgets.QMessageBox.critical(self, "Ошибка записи",
                                           f"Не удалось сохранить параметры. Ошибка - {type(exeption).__name__}",
                                           QtWidgets.QMessageBox.Ok)

    def load_config_file(self, empty=False):
        """ Загружает пользовательский файл конфигурации """
        try:
            path = f"{os.path.abspath(os.curdir)}\configurations"
            file = QtWidgets.QFileDialog.getOpenFileName(parent=application,
                                                         caption="Загрузить файл",
                                                         directory=path,
                                                         filter="All (*);;clbr59 (*.clbr59)",
                                                         initialFilter="clbr59 (*.clbr59)", )

            exemplary = "Средство калибровки"
            environment = "Условия калибровки"
            parameters = "Параметры прибора"
            out_irt = "Выход ИРТ"
            out_24 = "Выход 24В"
            out_pvi = "Выход ПВИ"
            participants = "Сдал/Принял/Дата"

            config = configparser.ConfigParser()
            config.read(file[0], encoding="UTF-8")

            _translate = QtCore.QCoreApplication.translate

            self.ui.comboBox_calibr_name.setItemText(0, _translate("MainWindow",
                                                                   config.get(exemplary, "калибратор")))

            self.ui.lineEdit_t.setText(config.get(environment, "Температура"))
            self.ui.lineEdit_f.setText(config.get(environment, "Влажность"))
            self.ui.lineEdit_p.setText(config.get(environment, "Давление"))

            self.ui.comboBox_parametr_type.setItemText(0,
                                                       _translate("MainWindow", config.get(parameters, "тип")))
            self.ui.lineEdit_parametr_number.setText(config.get(parameters, "номер"))
            self.ui.comboBox_parametr_year.setItemText(0, _translate("MainWindow",
                                                                     config.get(parameters, "год выпуска")))
            self.ui.lineEdit_parametr_position.setText(config.get(parameters, "позиция"))
            self.ui.comboBox_in_signal_type.setItemText(0, _translate("MainWindow",
                                                                      config.get(parameters, "тип входа")))
            self.ui.lineEdit_in_start_value.setText(config.get(parameters, "вход начало шкалы"))
            self.ui.lineEdit_in_end_value.setText(config.get(parameters, "вход конец шкалы"))
            self.ui.comboBox_out_signal_type.setItemText(0, _translate("MainWindow",
                                                                       config.get(parameters, "тип выхода")))
            self.ui.lineEdit_out_start_value.setText(config.get(parameters, "выход начало шкалы"))
            self.ui.lineEdit_out_end_value.setText(config.get(parameters, "выход конец шкалы"))
            if config.get(parameters, "наличие пви") == 'True':
                self.ui.checkBox_pvi.setChecked(True)
            else:
                self.ui.checkBox_pvi.setChecked(False)
            self.ui.lineEdit_pvi_scale_start.setText(config.get(parameters, "пви начало шкалы"))
            self.ui.lineEdit_pvi_scale_end.setText(config.get(parameters, "пви конец шкалы"))
            self.ui.comboBox_pvi_out.setItemText(0, _translate("MainWindow",
                                                               config.get(parameters, "пви тип выхода")))

            self.ui.lineEdit_out_irt_value_5.setText(config.get(out_irt, "показания 5"))
            self.ui.lineEdit_out_irt_value_25.setText(config.get(out_irt, "показания 25"))
            self.ui.lineEdit_out_irt_value_50.setText(config.get(out_irt, "показания 50"))
            self.ui.lineEdit_out_irt_value_75.setText(config.get(out_irt, "показания 75"))
            self.ui.lineEdit_out_irt_value_95.setText(config.get(out_irt, "показания 95"))

            self.ui.lineEdit_out_24_value_0.setText(config.get(out_24, "выход r0"))
            self.ui.lineEdit_out_24_value_820.setText(config.get(out_24, "выход r820"))

            self.ui.lineEdit_out_pvi_value_5.setText(config.get(out_pvi, "показания 5"))
            self.ui.lineEdit_out_pvi_value_25.setText(config.get(out_pvi, "показания 25"))
            self.ui.lineEdit_out_pvi_value_50.setText(config.get(out_pvi, "показания 50"))
            self.ui.lineEdit_out_pvi_value_75.setText(config.get(out_pvi, "показания 75"))
            self.ui.lineEdit_out_pvi_value_95.setText(config.get(out_pvi, "показания 95"))

            self.ui.lineEdit_out_pvi_output_5.setText(config.get(out_pvi, "выход 5"))
            self.ui.lineEdit_out_pvi_output_25.setText(config.get(out_pvi, "выход 25"))
            self.ui.lineEdit_out_pvi_output_50.setText(config.get(out_pvi, "выход 50"))
            self.ui.lineEdit_out_pvi_output_75.setText(config.get(out_pvi, "выход 75"))
            self.ui.lineEdit_out_pvi_output_95.setText(config.get(out_pvi, "выход 95"))

            self.ui.comboBox_passed.setItemText(0, _translate("MainWindow", config.get(participants, "сдал")))
            self.ui.comboBox_adopted.setItemText(0, _translate("MainWindow", config.get(participants, "принял")))

            date_c = tuple(map(int, config.get(participants, "дата калибровки(ДД.ММ.ГГГГ)").split('.')))
            self.ui.dateEdit_date_calibration.setDate(QtCore.QDate(date_c[0], date_c[1], date_c[2]))
        except configparser.NoSectionError:
            pass

    def open_settings(self):
        self.settings_app = ClbrSettings()
        self.settings_app.setWindowTitle("Clbr59xx - Настройки")
        self.settings_app.setFixedSize(800, 620)
        self.settings_app.show()
        # self.settings_app.exec_()  # модальное

    def cal_clear(self):
        """ Очищает редактируемые виджеты """
        w = self.ui

        widgets = (w.lineEdit_parametr_number, w.lineEdit_parametr_position, w.lineEdit_in_start_value,
                   w.lineEdit_in_end_value, w.lineEdit_out_start_value, w.lineEdit_out_end_value,
                   w.lineEdit_out_irt_value_5, w.lineEdit_out_irt_value_25, w.lineEdit_out_irt_value_50,
                   w.lineEdit_out_irt_value_75, w.lineEdit_out_irt_value_95, w.lineEdit_out_pvi_value_5,
                   w.lineEdit_out_pvi_value_25, w.lineEdit_out_pvi_value_50, w.lineEdit_out_pvi_value_75,
                   w.lineEdit_out_pvi_value_95, w.lineEdit_out_24_value_0, w.lineEdit_out_24_value_820,
                   w.lineEdit_pvi_scale_start, w.lineEdit_pvi_scale_end, w.lineEdit_t, w.lineEdit_f, w.lineEdit_p)
        for _w in widgets:
            _w.clear()

        self.ui.comboBox_calibr_name.setCurrentIndex(0)

        w.comboBox_calibr_name.setItemText(0, '')
        w.comboBox_parametr_type.setItemText(0, '')
        w.comboBox_parametr_year.setItemText(0, '')
        w.comboBox_in_signal_type.setItemText(0, '')
        w.comboBox_out_signal_type.setItemText(0, '')
        w.comboBox_pvi_out.setItemText(0, '')
        w.comboBox_passed.setItemText(0, '')
        w.comboBox_adopted.setItemText(0, '')

        w.checkBox_pvi.setChecked(False)

    def create_protocol_verify(self):
        """ Проверка на наличие отклонений от допусков, перед созданием книги протокола. """
        if False in ClbrMain.verification_error.values():
            dialog = QtWidgets.QMessageBox.question(application, "Отклонения в допусках",
                                                    "Обнаружены отклонения от допусков. Продолжить?",
                                                    buttons=QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No,
                                                    defaultButton=QtWidgets.QMessageBox.Yes)
            if dialog == 65536:
                pass
            elif dialog == 16384:
                self.create_protocol()
        else:
            self.create_protocol()

    def progress(self, value, format):
        self.ui.progressBar.setValue(value)
        self.ui.progressBar.setFormat(format)

    def create_protocol(self):
        self.ui.progressBar.show()
        self.progress(0, "Начинаю формирование протокола")

        position = self.ui.lineEdit_parametr_position.text()
        if position == '':
            position = 'blank'

        file_position = f"{position}.xlsx"
        position_name = f"protocols/{file_position}"

        self.progress(5, "Проверяю наличие файла шаблона")
        if os.path.exists("Template_CalibrationIRT59xx.xlsx") == False:
            QtWidgets.QMessageBox.critical(self, "Ошибка",
                                           "Отсутствует файл шаблона - Template_CalibrationIRT59xx.xlsx",
                                           QtWidgets.QMessageBox.Ok)
        else:
            try:
                self.progress(10, "Копирую файл шаблона")
                shutil.copy("Template_CalibrationIRT59xx.xlsx", position_name)
            except PermissionError:
                QtWidgets.QMessageBox.critical(self, "Ошибка",
                                               f"Не могу открыть {file_position} для записи, возможно файл открыт в \
                                               другой программе",
                                               QtWidgets.QMessageBox.Ok)

            try:
                self.progress(20, "Загружаю книгу")
                wb = openpyxl.load_workbook(position_name)
                ws = wb.active

                cells = configparser.ConfigParser()
                cells.read("parameters.ini", encoding="UTF-8")

                self.progress(25, "Проверяю секции выходных ячеек")
                if not cells.has_section('Выходные ячейки'):
                    dialog = QtWidgets.QMessageBox.question(application, "Настройки ячеек",
                                                            "Отсутствуют настройки ячеек. Открыть настройки?",
                                                            buttons=QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Yes,
                                                            defaultButton=QtWidgets.QMessageBox.Yes)
                    if dialog == 65536:
                        pass
                    if dialog == 16384:
                        self.open_settings()

                else:
                    self.progress(45, "Загружаю значения")
                    section = 'Выходные ячейки'

                    calibr_name, calibr_number = '', ''
                    try:
                        calibr_name = self.ui.comboBox_calibr_name.currentText().split()[0]
                        calibr_number = self.ui.comboBox_calibr_name.currentText().split()[1]
                    except:
                        pass

                    # передача без округления - условия и параметры прибора
                    cells_dict_txt = {
                        'calibr_type': calibr_name,
                        'calibr_number': calibr_number,
                        't': self.ui.lineEdit_t.text(),
                        'f': self.ui.lineEdit_f.text(),
                        'p': self.ui.lineEdit_p.text(),

                        'parametr_type': self.ui.comboBox_parametr_type.currentText(),
                        'parametr_number': self.ui.lineEdit_parametr_number.text(),
                        'parametr_year': self.ui.comboBox_parametr_year.currentText(),
                        'parametr_position': self.ui.lineEdit_parametr_position.text(),
                        'in_signal': self.ui.comboBox_in_signal_type.currentText(),
                        'in_signal_start': self.ui.lineEdit_in_start_value.text(),
                        'in_signal_end': self.ui.lineEdit_in_end_value.text(),
                        'out_signal': self.ui.comboBox_out_signal_type.currentText(),
                        'out_signal_start': self.ui.lineEdit_out_start_value.text(),
                        'out_signal_end': self.ui.lineEdit_out_end_value.text(),
                        'pvi_scale_start': self.ui.lineEdit_pvi_scale_start.text(),
                        'pvi_scale_end': self.ui.lineEdit_pvi_scale_end.text(),
                        'pvi_scale_out': self.ui.comboBox_pvi_out.currentText(),

                        'passed': self.ui.comboBox_passed.currentText(),
                        'adopted': self.ui.comboBox_adopted.currentText(),
                        'date_calibration': self.ui.dateEdit_date_calibration.dateTime().toString('dd.MM.yyyy'),

                        'result': ("не годен" if False in ClbrMain.verification_error.values() else "годен")
                    }

                    # передача с округлением, поля участвующие в расчетах
                    cells_dict_num = {
                        'out_irt_value_5': self.ui.lineEdit_out_irt_value_5.text(),
                        'out_irt_value_25': self.ui.lineEdit_out_irt_value_25.text(),
                        'out_irt_value_50': self.ui.lineEdit_out_irt_value_50.text(),
                        'out_irt_value_75': self.ui.lineEdit_out_irt_value_75.text(),
                        'out_irt_value_95': self.ui.lineEdit_out_irt_value_95.text(),
                        'out_irt_output_5': self.ui.lineEdit_out_irt_output_5.text(),
                        'out_irt_output_25': self.ui.lineEdit_out_irt_output_25.text(),
                        'out_irt_output_50': self.ui.lineEdit_out_irt_output_50.text(),
                        'out_irt_output_75': self.ui.lineEdit_out_irt_output_75.text(),
                        'out_irt_output_95': self.ui.lineEdit_out_irt_output_95.text(),
                        'out_irt_in_5': self.ui.lineEdit_out_irt_in_5.text(),
                        'out_irt_in_25': self.ui.lineEdit_out_irt_in_25.text(),
                        'out_irt_in_50': self.ui.lineEdit_out_irt_in_50.text(),
                        'out_irt_in_75': self.ui.lineEdit_out_irt_in_75.text(),
                        'out_irt_in_95': self.ui.lineEdit_out_irt_in_95.text(),

                        'out_24_value': self.ui.lineEdit_out_24_value_0.text(),
                        'out_24_value_820': self.ui.lineEdit_out_24_value_820.text(),
                        'out_24_in': self.ui.lineEdit_out_24_in_0.text(),
                        'out_24_in_820': self.ui.lineEdit_out_24_in_820.text(),

                        'out_pvi_value_5': self.ui.lineEdit_out_pvi_value_5.text(),
                        'out_pvi_value_25': self.ui.lineEdit_out_pvi_value_25.text(),
                        'out_pvi_value_50': self.ui.lineEdit_out_pvi_value_50.text(),
                        'out_pvi_value_75': self.ui.lineEdit_out_pvi_value_75.text(),
                        'out_pvi_value_95': self.ui.lineEdit_out_pvi_value_95.text(),
                        'out_pvi_output_5': self.ui.lineEdit_out_pvi_output_5.text(),
                        'out_pvi_output_25': self.ui.lineEdit_out_pvi_output_25.text(),
                        'out_pvi_output_50': self.ui.lineEdit_out_pvi_output_50.text(),
                        'out_pvi_output_75': self.ui.lineEdit_out_pvi_output_75.text(),
                        'out_pvi_output_95': self.ui.lineEdit_out_pvi_output_95.text(),
                        'out_pvi_in_5': self.ui.lineEdit_out_pvi_in_5.text(),
                        'out_pvi_in_25': self.ui.lineEdit_out_pvi_in_25.text(),
                        'out_pvi_in_50': self.ui.lineEdit_out_pvi_in_50.text(),
                        'out_pvi_in_75': self.ui.lineEdit_out_pvi_in_75.text(),
                        'out_pvi_in_95': self.ui.lineEdit_out_pvi_in_95.text(),

                        'acceptance_error_irt': ClbrMain.permissible_inaccuracy_irt,
                        'acceptance_error_24': ClbrMain.permissible_inaccuracy_24v,
                        'acceptance_error_pvi': ClbrMain.permissible_inaccuracy_pvi,
                    }

                    columns = ' ABCDEFGHIJKLMNOPQRSTUVWXYZ'  # столбцы, ws['A'] = value не работает для объединенных ячеек

                    self.progress(65, "Записываю значения в книгу")
                    # без округления cells_dict_txt
                    for key, value in cells_dict_txt.items():
                        for cell_p in cells.get(section, key).split():
                            if value == '':
                                value = '—'

                            ws.cell(int(cell_p[1:]), int(columns.index(cell_p[0])), str(value.replace(".", ",")))

                    # с округлением cells_dict_num
                    for key, value in cells_dict_num.items():
                        for cell_p in cells.get(section, key).split():
                            if value == '':
                                value = '—'
                            else:
                                try:
                                    value = f"{float(value):0>.3f}".replace(".", ",")
                                except:
                                    pass

                            ws.cell(int(cell_p[1:]), int(columns.index(cell_p[0])), value)

                    self.progress(85, "Сохраняю книгу")
                    wb.save(position_name)

                    self.progress(95, f"Открываю протокол калибровки")
                    file_path = f"{os.path.abspath(os.curdir)}/{position_name}"

                    if sys.platform == "win32":
                        os.startfile(file_path)
                    else:
                        dialog = QtWidgets.QMessageBox.question(application, str(position_name),
                                                                f"Протокол {file_position} создан в директории \
                                                                {os.path.abspath(os.curdir)}/protocols.\n\n \
    Открыть <Office online> на сайте microsoft.com в браузере по умолчанию?",
                                                                buttons=QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Yes,
                                                                defaultButton=QtWidgets.QMessageBox.Yes)
                        if dialog == 65536:
                            pass
                        if dialog == 16384:
                            wb_open_new("https://www.office.com/?flight=unauthrefresh&auth=1")

                    parametr_position = cells_dict_txt["parametr_position"]
                    self.progress(100, f"Протокол калибровки {parametr_position} готов")
                    self.ui.progressBar.hide()

            except Exception as exeption:
                self.ui.progressBar.hide()
                QtWidgets.QMessageBox.critical(self, "Ошибка",
                                               f"Не удалось Создать протокол. Ошибка - {type(exeption).__name__}",
                                               QtWidgets.QMessageBox.Ok)

    def help_error(self):
        try:
            path = "C:\Program Files\Microsoft Office\Office16\outlook.exe"
            subprocess.Popen([path, "/c", "ipm.note", "/m", "help@ya.ru"])
        except Exception as exeption:
            QtWidgets.QMessageBox.critical(self, "Ошибка",
                                           f"Не могу найти почтовую программу. Ошибка - {type(exeption).__name__}",
                                           QtWidgets.QMessageBox.Ok)

    def ui_help(self):
        os.startfile(f"{os.path.abspath(os.curdir)}\doc\Pasport_IRT_5920.pdf")

    def about(self):
        QtWidgets.QMessageBox.aboutQt(application, title="О программе")

    def exit(self):
        dialog = QtWidgets.QMessageBox.question(application, "Выход из программы",
                                                "Сохранить файл конфигурации прибора?",
                                                buttons=QtWidgets.QMessageBox.Cancel | QtWidgets.QMessageBox.No | QtWidgets.QMessageBox.Yes,
                                                defaultButton=QtWidgets.QMessageBox.Yes)
        if dialog == 65536:
            sys.exit(app.exec())
        if dialog == 16384:
            self.save_config_file()
            sys.exit(app.exec())
        if dialog == 4194304:
            pass

    def positive_testing(self):
        """ Заполнение полей случайными значениями, в пределах допуска """
        accept_irt = self.acceptance_irt(True) # запрос допуска ИРТ
        accept_pvi = self.out_pvi_out(True) # запрос допуска ПВИ

        try:
            # ИРТ
            irt_5 = float(self.ui.lineEdit_out_irt_output_5.text())
            irt_25 = float(self.ui.lineEdit_out_irt_output_25.text())
            irt_50 = float(self.ui.lineEdit_out_irt_output_50.text())
            irt_75 = float(self.ui.lineEdit_out_irt_output_75.text())
            irt_95 = float(self.ui.lineEdit_out_irt_output_95.text())

            irt_random_5 = str(round(random.uniform(irt_5 - accept_irt * 0.2, irt_5 + accept_irt * 0.15), 3))
            irt_random_25 = str(round(random.uniform(irt_25 - accept_irt * 0.1, irt_25), 3))
            irt_random_50 = str(round(random.uniform(irt_50 - accept_irt * 0.95, irt_50 + accept_irt * 0.1), 3))
            irt_random_75 = str(round(random.uniform(irt_75 - accept_irt * 0.11, irt_75 + accept_irt * 0.14), 3))
            irt_random_95 = str(round(random.uniform(irt_95 + accept_irt * 0.2, irt_95), 3))

            self.ui.lineEdit_out_irt_value_5.setText(irt_random_5)
            self.ui.lineEdit_out_irt_value_25.setText(irt_random_25)
            self.ui.lineEdit_out_irt_value_50.setText(irt_random_50)
            self.ui.lineEdit_out_irt_value_75.setText(irt_random_75)
            self.ui.lineEdit_out_irt_value_95.setText(irt_random_95)

            # 24В
            v24_r0_random = round(random.uniform(24.0, 24.479), 3)
            v24_r820_random = round(random.uniform(23.9, v24_r0_random), 3)

            self.ui.lineEdit_out_24_value_0.setText(str(v24_r0_random))
            self.ui.lineEdit_out_24_value_820.setText(str(v24_r820_random))

            # ПВИ
            pvi_5 = float(self.ui.lineEdit_out_pvi_output_5.text())
            pvi_25 = float(self.ui.lineEdit_out_pvi_output_25.text())
            pvi_50 = float(self.ui.lineEdit_out_pvi_output_50.text())
            pvi_75 = float(self.ui.lineEdit_out_pvi_output_75.text())
            pvi_95 = float(self.ui.lineEdit_out_pvi_output_95.text())

            pvi_random_5 = str(round(random.uniform(pvi_5 - accept_pvi * 0.23, pvi_5 + accept_pvi * 0.25), 3))
            pvi_random_25 = str(round(random.uniform(pvi_25 - accept_pvi * 0.2, pvi_25), 3))
            pvi_random_50 = str(round(random.uniform(pvi_50 - accept_pvi * 0.14, pvi_50 + accept_pvi * 0.17), 3))
            pvi_random_75 = str(round(random.uniform(pvi_75 - accept_pvi * 0.18, pvi_75 + accept_pvi * 0.14), 3))
            pvi_random_95 = str(round(random.uniform(pvi_95 + accept_pvi * 0.2, pvi_95), 3))

            self.ui.lineEdit_out_pvi_value_5.setText(pvi_random_5)
            self.ui.lineEdit_out_pvi_value_25.setText(pvi_random_25)
            self.ui.lineEdit_out_pvi_value_50.setText(pvi_random_50)
            self.ui.lineEdit_out_pvi_value_75.setText(pvi_random_75)
            self.ui.lineEdit_out_pvi_value_95.setText(pvi_random_95)
        except ValueError:
            pass


if __name__ == "__main__":
    app = QtWidgets.QApplication([])
    application = ClbrMain()
    application.setWindowTitle(" Clbr59xx - Создание протокола калибровки")
    application.setFixedSize(800, 670)
    application.show()

    sys.exit(app.exec())
