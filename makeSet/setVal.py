from itertools import count
import openpyxl
from collections import defaultdict
from effectSet import 功效集合
from check import checkarr
from check import checkeffect

# open the excel file
data = openpyxl.load_workbook("方劑單一藥效.xlsx")


class Graph:
    indegree = None
    # construct

    def __init__(self, vertices):
        # construct表示點之間關係的dict{v: [u, i]}(v,u,i都是點,表示邊<v, u>, <v, i>)
        self.graph = defaultdict(list)
        # 點的個數
        self.V = vertices

    # 邊
    def add_edge(self, u, v):
        # 邊<u, v>
        self.graph[u].append(v)

    # 拓樸排序
    # self.graph['a']->a連向的人
    def topological_sort(self, arr):
        key = list(self.graph.keys())

        for pointtonext in key:
            for i in key:
                for j in self.graph[i]:
                    if(j == pointtonext):
                        arr[pointtonext] += 1
                        print(arr[pointtonext])
        upnum = 100
        for pointtonext in key:

            if arr[pointtonext] == 0:
                arr[pointtonext] = upnum
        con = True
        while(con):
            con = False
            num = upnum-10
            for i in key:
                if arr[i] == upnum:
                    for j in self.graph[i]:
                        arr[j] = num
                        con = True
            upnum -= 10

    # 排序


def check_prescription_order(rowstart, rowfinal, e):
    effectval = 100
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
    g = Graph(100)
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

    g.topological_sort(功效集合[i])
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
