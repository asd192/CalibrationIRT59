import configparser
from PyQt5 import QtWidgets

from settings import Ui_settings


class ClbrSettings(QtWidgets.QDialog):
    def __init__(self):
        super(ClbrSettings, self).__init__()
        self.sl = Ui_settings()
        self.sl.setupUi(self)

        self.load_sett()
        self.sl.save_settings.clicked.connect(self.save_sett)

    def dict_sett(self):
        sett_p = {
            'calibr_type': self.sl.s_calibr_type,
            'calibr_number': self.sl.s_calibr_number,
            't': self.sl.s_t,
            'f': self.sl.s_f,
            'p': self.sl.s_p,

            'parametr_type': self.sl.s_parametr_type,
            'parametr_number': self.sl.s_parametr_number,
            'parametr_year': self.sl.s_parametr_year,
            'parametr_position': self.sl.s_parametr_position,
            'in_signal': self.sl.s_in_signal,
            'in_signal_start': self.sl.s_in_signal_start,
            'in_signal_end': self.sl.s_in_signal_end,
            'out_signal': self.sl.s_out_signal,
            'out_signal_start': self.sl.s_out_signal_start,
            'out_signal_end': self.sl.s_out_signal_end,
            'pvi_scale_start': self.sl.s_pvi_scale_start,
            'pvi_scale_end': self.sl.s_pvi_scale_end,
            'pvi_scale_out': self.sl.s_pvi_scale_out,

            'out_irt_value_5': self.sl.s_out_irt_value_5,
            'out_irt_value_25': self.sl.s_out_irt_value_25,
            'out_irt_value_50': self.sl.s_out_irt_value_50,
            'out_irt_value_75': self.sl.s_out_irt_value_75,
            'out_irt_value_95': self.sl.s_out_irt_value_95,
            'out_irt_output_5': self.sl.s_out_irt_output_5,
            'out_irt_output_25': self.sl.s_out_irt_output_25,
            'out_irt_output_50': self.sl.s_out_irt_output_50,
            'out_irt_output_75': self.sl.s_out_irt_output_75,
            'out_irt_output_95': self.sl.s_out_irt_output_95,
            'out_irt_in_5': self.sl.s_out_irt_in_5,
            'out_irt_in_25': self.sl.s_out_irt_in_25,
            'out_irt_in_50': self.sl.s_out_irt_in_50,
            'out_irt_in_75': self.sl.s_out_irt_in_75,
            'out_irt_in_95': self.sl.s_out_irt_in_95,
            'acceptance_error_irt': self.sl.s_acceptance_error_irt,

            'out_24_value': self.sl.s_out_24_value,
            'out_24_value_820': self.sl.s_out_24_value_820,
            'out_24_in': self.sl.s_out_24_in,
            'out_24_in_820': self.sl.s_out_24_in_820,
            'acceptance_error_24': self.sl.s_acceptance_error_24,

            'out_pvi_value_5': self.sl.s_out_pvi_value_5,
            'out_pvi_value_25': self.sl.s_out_pvi_value_25,
            'out_pvi_value_50': self.sl.s_out_pvi_value_50,
            'out_pvi_value_75': self.sl.s_out_pvi_value_75,
            'out_pvi_value_95': self.sl.s_out_pvi_value_95,
            'out_pvi_output_5': self.sl.s_out_pvi_output_5,
            'out_pvi_output_25': self.sl.s_out_pvi_output_25,
            'out_pvi_output_50': self.sl.s_out_pvi_output_50,
            'out_pvi_output_75': self.sl.s_out_pvi_output_75,
            'out_pvi_output_95': self.sl.s_out_pvi_output_95,
            'out_pvi_in_5': self.sl.s_out_pvi_in_5,
            'out_pvi_in_25': self.sl.s_out_pvi_in_25,
            'out_pvi_in_50': self.sl.s_out_pvi_in_50,
            'out_pvi_in_75': self.sl.s_out_pvi_in_75,
            'out_pvi_in_95': self.sl.s_out_pvi_in_95,
            'acceptance_error_pvi': self.sl.s_acceptance_error_pvi,

            'passed': self.sl.s_passed,
            'adopted': self.sl.s_adopted,
            'date_calibration': self.sl.s_date_calibration,

            'result': self.sl.s_result,
        }
        return sett_p

    def load_sett(self):
        try:
            dict_sett = self.dict_sett()

            settings = configparser.ConfigParser()
            settings.read("parameters.ini", encoding="utf-8")
            section = "Выходные ячейки"

            for key in dict_sett.keys():
                dict_sett[key].setText(settings.get(section, key))

        except Exception as exeption:
            QtWidgets.QMessageBox.critical(self, "Ошибка",
                                           f"Не удалось прочитать файл <parameters.ini>.Ошибка - {type(exeption).__name__}",
                                           QtWidgets.QMessageBox.Ok)

    def save_sett(self):
        try:
            dict_sett = self.dict_sett()

            settings = configparser.ConfigParser()
            settings.read("parameters.ini", encoding="utf-8")

            section = "Выходные ячейки"
            if not settings.has_section(section):
                settings.add_section(section)

            for key, value in dict_sett.items():
                settings.set(section, key, value.text())

            with open("parameters.ini", "w", encoding="utf8") as config_file:
                settings.write(config_file)

                QtWidgets.QMessageBox.information(self, "Сохранение настроек",
                                                  "Настроки успешно сохранены!",
                                                  QtWidgets.QMessageBox.Ok)
        except Exception as exeption:
            QtWidgets.QMessageBox.critical(self, "Ошибка записи",
                                           f"Не удалось сохранить настройки. Ошибка - {type(exeption).__name__}",
                                           QtWidgets.QMessageBox.Ok)
