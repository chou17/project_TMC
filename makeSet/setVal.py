import openpyxl
from effectSet import 功效集合
from check import checkarr

# open the excel file
data = openpyxl.load_workbook("原始藥效.xlsx")

# 確認方劑內的大小關係


def check_prescription_order(rowstart, rowfinal):
    effectval = 100
    for row in range(rowstart, rowfinal):
        medicinename = sheet.cell(row, 3).value
        effect = sheet.cell(row, 5).value
        effectindex = checkarr(effect)
        if effectindex != -1:
            # for index in range(len(功效集合[effectIndex]))
            for key in 功效集合[effectindex].keys():
                #medEffect = 功效集合
                if medicinename == key:
                    功效集合[effectindex][key] = effectval
                    effectval -= 10
                    break


for sheet in data.worksheets:
    # 第一個方劑在第二行(e.g.,from 麻黃湯)
    rowStart = 2
    for row in range(3, sheet.max_row, 1):
        prescription = sheet.cell(row, 1).value
        # 如果到了下一個方劑(e.g.,桂枝湯)
        if prescription != None:
            rowFinal = row-1
            check_prescription_order(rowStart, rowFinal)
            rowStart = row
