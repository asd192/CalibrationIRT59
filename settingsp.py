from PyQt5.QtWidgets import QDialog
from settings import Ui_settings


class WindowSettings(QDialog):
    def __init__(self):
        super(WindowSettings, self).__init__()
        self.st = Ui_settings()
        self.st.setupUi(self)

        self.st.save_settings.clicked.connect(self.save_sett)

    def save_sett(self):
        sett_p = {
            0: self.st.s_calibr_type.text(),
            1: self.st.s_calibr_number.text(),
            2: self.st.s_t.text(),
            3: self.st.s_f.text(),
            4: self.st.s_p.text(),
            5: self.st.s_parametr_type.text(),
            6: self.st.s_parametr_number.text(),
            7: self.st.s_parametr_yeaar.text(),
            8: self.st.s_parametr_position.text(),
            9: self.st.s_in_signal.text(),
            10: self.st.s_in_signal_start.text(),
            11: self.st.s_in_signal_end.text(),
            12: self.st.s_out_signal.text(),
            13: self.st.s_out_signal_start.text(),
            14: self.st.s_out_signal_end.text(),
            15: self.st.s_pvi_scale_start.text(),
            16: self.st.s_pvi_scale_end.text(),
            17: self.st.s_pvi_scale_out.text(),
            18: self.st.s_calibr_type.text(),
            19: self.st.s_calibr_type.text(),

            20: self.st.s_passed.text(),
            21: self.st.s_adopted.text(),
            22: self.st.s_date_calibration.text(),

            23: self.st.s_out_irt_value_5.text(),
            24: self.st.s_out_irt_value_25.text(),
            25: self.st.s_out_irt_value_50.text(),
            26: self.st.s_out_irt_value_75.text(),
            27: self.st.s_out_irt_value_95.text(),
            28: self.st.s_out_irt_output_5.text(),
            29: self.st.s_out_irt_output_25.text(),
            30: self.st.s_out_irt_output_50.text(),
            31: self.st.s_out_irt_output_75.text(),
            32: self.st.s_out_irt_output_95.text(),
            33: self.st.s_out_irt_in_5.text(),
            34: self.st.s_out_irt_in_25.text(),
            35: self.st.s_out_irt_in_50.text(),
            36: self.st.s_out_irt_in_75.text(),
            37: self.st.s_out_irt_in_95.text(),

            38: self.st.s_out_24_value.text(),
            39: self.st.s_out_24_value_820.text(),
            40: self.st.s_out_24_in.text(),
            41: self.st.s_out_24_in_820.text(),

            42: self.st.s_out_pvi_value_5.text(),
            44: self.st.s_out_pvi_value_25.text(),
            45: self.st.s_out_pvi_value_50.text(),
            46: self.st.s_out_pvi_value_75.text(),
            47: self.st.s_out_pvi_value_95.text(),
            48: self.st.s_out_pvi_output_5.text(),
            49: self.st.s_out_pvi_output_25.text(),
            50: self.st.s_out_pvi_output_50.text(),
            51: self.st.s_out_pvi_output_75.text(),
            52: self.st.s_out_pvi_output_95.text(),
            53: self.st.s_out_pvi_in_5.text(),
            54: self.st.s_out_pvi_in_25.text(),
            55: self.st.s_out_pvi_in_50.text(),
            56: self.st.s_out_pvi_in_75.text(),
            57: self.st.s_out_pvi_in_95.text(),
        }

        import configparser

        config = configparser.ConfigParser()
        config.read("parameters.ini", encoding="utf-8")
        section = "Выходные ячейки"
        config.add_section(section)

        for key, value in sett_p.items():
            config.set(section, key, value)

