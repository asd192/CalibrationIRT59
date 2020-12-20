import os, openpyxl, shutil


def create_clbr():
    if os.path.exists("Шаблон_CalibrationIRT59xx.xlsx") == False:
        return False
    else:
        shutil.copy("Шаблон_CalibrationIRT59xx.xlsx", "_temporary.xlsx")

        wb = openpyxl.load_workbook("_temporary.xlsx")
        ws = wb.active

        ws['C3'] = 42

        wb.save("_temporary.xlsx")

if __name__ == "__main__":
    create_clbr()
    # os.remove("_temporary.xlsx")




