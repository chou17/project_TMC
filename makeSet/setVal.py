from itertools import count
import openpyxl
from collections import defaultdict
from effectSet import 功效集合
from check import checkarr
from check import checkeffect

# open the excel file
data = openpyxl.load_workbook("方劑單一藥效.xlsx")


class Graph:
    # construct
    def __init__(self):
        self.graph = defaultdict(list)

    def add_edge(self, u, v):
        # edge<u, v>
        self.graph[u].append(v)

    #設定效力值
    def set_val(self, arr):
        # 有指向其他點的點
        key = list(self.graph.keys())
        # 如果有被指向表示不是最大的點
        for pointtonext in key:
            for i in key:
                for j in self.graph[i]:
                    if(j == pointtonext):
                        arr[pointtonext] += 1
        # 最大值為100
        upnum = 100
        for pointtonext in key:
            # 如果沒被指向arr[pointtonext]會=0 表示為最大的點
            if arr[pointtonext] == 0:
                arr[pointtonext] = upnum
        # 是否繼續執行
        con = True
        while(con):
            con = False
            num = upnum-10
            for i in key:
                # 如果有被上一層的人指向，表示為下一層
                if arr[i] == upnum:
                    for j in self.graph[i]:
                        arr[j] = num
                        con = True
            upnum -= 10


def check_prescription_order(rowstart, rowfinal, e):
    appear = False
    premed = ""
    for row in range(rowstart, rowfinal):
        medicinename = sheet.cell(row, 3).value
        effect = sheet.cell(row, 4).value
        effectindex = checkarr(effect)
        if effectindex == e:
            if appear == False:
                appear = True
                premed = medicinename
            else:
                g.add_edge(premed, medicinename)
                premed = medicinename


def get_key(dict, val):
    return [k for k, v in dict.items() if v == val]


for i in range(0, 2):
    g = Graph()
    for sheet in data.worksheets:
        # 第一個方劑在第二行(e.g.,from 麻黃湯)
        rowStart = 2
        for row in range(3, sheet.max_row, 1):
            prescription = sheet.cell(row, 1).value
            # 如果到了下一個方劑(e.g.,桂枝湯)
            if prescription != None:
                rowFinal = row
                check_prescription_order(rowStart, rowFinal, i)
                rowStart = row

    g.set_val(功效集合[i])
    only_set = get_key(功效集合[i], 0)
    for only in only_set:
        功效集合[i][only] = 100
    print(功效集合[i])


medData = openpyxl.Workbook()
medValSheet = medData.active  # Workbook.create_sheet()
# 功效跟藥材名都放在奇數行
keyIndex = 1
for effectIndex in range(len(功效集合)):
    # 藥效
    medValSheet.cell(1, keyIndex).value = checkeffect(effectIndex)
    # 藥材從第二行開始放
    rowindex = 2
    for key in 功效集合[effectIndex].keys():
        # 藥材名稱
        medValSheet.cell(row=rowindex, column=keyIndex).value = key
        # 效力值(放在偶數行)
        medValSheet.cell(row=rowindex, column=keyIndex +
                         1).value = 功效集合[effectIndex][key]
        rowindex += 1
    keyIndex = keyIndex + 2
    medData.save('效力值.xlsx')
