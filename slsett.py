import configparser
from PyQt5 import QtWidgets

from settings import Ui_settings


class ClbrSettings(QtWidgets.QDialog):
    def __init__(self):
        super(ClbrSettings, self).__init__()
        self.st = Ui_settings()
        self.st.setupUi(self)

        self.load_sett()
        self.st.save_settings.clicked.connect(self.save_sett)

    def dict_sett(self):
        sett_p = {
            '0': self.st.s_calibr_type,
            '1': self.st.s_calibr_number,
            '2': self.st.s_t,
            '3': self.st.s_f,
            '4': self.st.s_p,
            '5': self.st.s_parametr_type,
            '6': self.st.s_parametr_number,
            '7': self.st.s_parametr_yeaar,
            '8': self.st.s_parametr_position,
            '9': self.st.s_in_signal,
            '10': self.st.s_in_signal_start,
            '11': self.st.s_in_signal_end,
            '12': self.st.s_out_signal,
            '13': self.st.s_out_signal_start,
            '14': self.st.s_out_signal_end,
            '15': self.st.s_pvi_scale_start,
            '16': self.st.s_pvi_scale_end,
            '17': self.st.s_pvi_scale_out,
            '18': self.st.s_calibr_type,
            '19': self.st.s_calibr_type,

            '20': self.st.s_passed,
            '21': self.st.s_adopted,
            '22': self.st.s_date_calibration,

            '23': self.st.s_out_irt_value_5,
            '24': self.st.s_out_irt_value_25,
            '25': self.st.s_out_irt_value_50,
            '26': self.st.s_out_irt_value_75,
            '27': self.st.s_out_irt_value_95,
            '28': self.st.s_out_irt_output_5,
            '29': self.st.s_out_irt_output_25,
            '30': self.st.s_out_irt_output_50,
            '31': self.st.s_out_irt_output_75,
            '32': self.st.s_out_irt_output_95,
            '33': self.st.s_out_irt_in_5,
            '34': self.st.s_out_irt_in_25,
            '35': self.st.s_out_irt_in_50,
            '36': self.st.s_out_irt_in_75,
            '37': self.st.s_out_irt_in_95,

            '38': self.st.s_out_24_value,
            '39': self.st.s_out_24_value_820,
            '40': self.st.s_out_24_in,
            '41': self.st.s_out_24_in_820,

            '42': self.st.s_out_pvi_value_5,
            '43': self.st.s_out_pvi_value_25,
            '44': self.st.s_out_pvi_value_50,
            '45': self.st.s_out_pvi_value_75,
            '46': self.st.s_out_pvi_value_95,
            '47': self.st.s_out_pvi_output_5,
            '48': self.st.s_out_pvi_output_25,
            '49': self.st.s_out_pvi_output_50,
            '50': self.st.s_out_pvi_output_75,
            '51': self.st.s_out_pvi_output_95,
            '52': self.st.s_out_pvi_in_5,
            '53': self.st.s_out_pvi_in_25,
            '54': self.st.s_out_pvi_in_50,
            '55': self.st.s_out_pvi_in_75,
            '56': self.st.s_out_pvi_in_95,
        }
        return sett_p

    def load_sett(self):
        settings = configparser.ConfigParser()
        try:
            settings.read("parameters.ini", encoding="utf-8")

            sett_values = settings.items("Выходные ячейки")
            dict_sett = self.dict_sett()

            for i in range(len(dict_sett)):
                dict_sett[str(i)].setText(sett_values[i][1])

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
