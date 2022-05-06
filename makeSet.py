import openpyxl

功效集合 = [
    # 發散表寒:0
    {"川烏": 0, "藁本": 0, "蔥": 0, "川芎": 0, "桂枝": 0, "麻黃": 0, "獨活": 0, "防風": 0, "紫蘇": 0,
        "香薷": 0, "蘇梗": 0, "蒼耳子": 0, "生薑": 0, "辛夷": 0, "建神麯": 0, "胡荽": 0, "浮萍": 0},
    # 祛風寒:1
    {"生薑": 0, "附子": 0, "蒼耳草": 0, "細辛": 0, "桂枝": 0, "川芎": 0}
]

data = openpyxl.load_workbook("原始藥效.xlsx")
sheet = data.worksheets[0]

for row in range(2, sheet.max_row, 1):
    medicineName = sheet.cell(row, 2).value
    effectIndex = int(sheet.cell(row, 4).value[0])
    structure = sheet.cell(row, 1).value

    if structure == "君":
        for key in 功效集合[effectIndex].keys():
            if medicineName == key:
                if 功效集合[effectIndex][key] < 1001:
                    功效集合[effectIndex][key] = 1001
                else:
                    功效集合[effectIndex][key] += 1
                break
    elif structure == "臣":
        for key in 功效集合[effectIndex].keys():
            if medicineName == key:
                if 功效集合[effectIndex][key] < 101:
                    功效集合[effectIndex][key] = 101
                else:
                    功效集合[effectIndex][key] += 1
                break
    elif structure == "佐":
        for key in 功效集合[effectIndex].keys():
            if medicineName == key:
                if 功效集合[effectIndex][key] < 11:
                    功效集合[effectIndex][key] = 11
                else:
                    功效集合[effectIndex][key] += 1
                break
    elif structure == "使":
        for key in 功效集合[effectIndex].keys():
            if medicineName == key:
                if 功效集合[effectIndex][key] == 0:
                    功效集合[effectIndex][key] = 1
                else:
                    功效集合[effectIndex][key] += 1
                break


# data.save('原始藥效.xlsx')
