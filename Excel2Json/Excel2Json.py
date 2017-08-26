# -*- coding: UTF-8 -*- 
# 作者：林培华
# 策划的配置表格转为json格式存储
# 后续可能的功能转为protobuf二进制文件同时生成.proto 后续可能的功能转为protobuf二进制文件同时生成对应的C#脚本

import xlrd
import sys
import json

class FieldItem:
    def __init__(self):
        self.name = ""
        self.targets = "A"
        self.type = ""

fileOutput = open('buff.json', 'wb')

workbook = xlrd.open_workbook('buff.xlsx')
booksheet = workbook.sheet_by_name('Sheet1')

colnumer = booksheet.ncols
rownumber = booksheet.nrows
# 生成字段的信息
fieldInfo = []
for colIndex in range(0, colnumer):
    fieldItem = FieldItem()
    fieldItem.name = booksheet.cell(0, colIndex).value
    fieldItem.targets = booksheet.cell(1, colIndex).value
    fieldItem.type = booksheet.cell(2, colIndex).value
    fieldInfo.append(fieldItem)

def myStr(num):
    if type(num) == float and round(num) == num :
        return str(int(num))
    else:
        return str(num)

def parseCell(fieldItem, cell):
    print(cell.ctype)
    if fieldItem.type == 'string':
        return myStr(cell.value)
    elif fieldItem.type == 'int':
        return int(cell.value)
    elif fieldItem.type == 'float':
        return float(cell.value)
    elif fieldItem.type == 'vector<int>':
        split = cell.value.split(':')
        result = []
        for i in split:
            result.append(int(i))
        return result
OutData = {}
for rowIndex in range(3, rownumber):
    for colIndex in range(0, colnumer):
        fieldItem = fieldInfo[colIndex]
        cell = booksheet.cell(rowIndex, colIndex)
        if colIndex == 0:
            id = int(cell.value)
            OutData[id] = {}
        else:
            OutData[id][fieldItem.name] = parseCell(fieldItem, cell)

print(json.dumps(OutData,sort_keys=True, separators=(',', ': ')))


