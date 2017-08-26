# -*- coding: UTF-8 -*- 
# 作者：林培华
# 策划的配置表格转为json格式存储
# 后续可能的功能转为protobuf二进制文件同时生成.proto 后续可能的功能转为protobuf二进制文件同时生成对应的C#脚本

import xlrd
import sys
import json
import re

class FieldItem:
    def __init__(self):
        self.name = ""
        self.targets = "A"
        self.type = ""

fileOutput = open('buff.json', 'wb')

workbook = xlrd.open_workbook('buff.xlsx')
booksheet = workbook.sheet_by_name('buff')

colnumer = booksheet.ncols
rownumber = booksheet.nrows
# 生成字段的信息
fieldInfo = []
for colIndex in range(0, colnumer):
    fieldItem = FieldItem()
    fieldItem.name = booksheet.cell(3, colIndex).value
    fieldItem.targets = booksheet.cell(1, colIndex).value
    fieldItem.type = booksheet.cell(2, colIndex).value.replace(' ','')
    fieldInfo.append(fieldItem)

def myStr(num):
    if type(num) == float and round(num) == num:
        return str(int(num))
    elif type(num) == float or type(num) == int or type(num) == long:
        return str(num)
    else:
        return num.encode('utf-8')

VectorPatten = re.compile('vector<(.*?)>', re.DOTALL | re.IGNORECASE)
MapPatten = re.compile('map<(.*?),(.*?)>', re.DOTALL | re.IGNORECASE)

def parseCell(fieldItem, cell):
    def _innerParse(subType, value):
        if subType == 'string':
            return myStr(value)
        elif subType == 'int':
            print(value)
            return int(value)
        elif subType == 'float':
            return float(value)

    fieldType = fieldItem.type
    m = VectorPatten.search(fieldType)
    if m != None:
        subType = m.group(0)
        split = cell.value.split(':')
        result = []
        for i in split:
            result.append(_innerParse(subType, i))

        return result
    m = MapPatten.search(fieldType)
    if m != None:
        subType1 = m.group(1)
        subType2 = m.group(2)
        result = {}
        split = cell.value.split('|')
        for i in split:
            tmpsplit = i.split(':')
            result[_innerParse(subType1, tmpsplit[0])] = _innerParse(subType2, tmpsplit[1])
        return result
    else:
        return _innerParse(fieldType, cell.value)


OutData = {}
for rowIndex in range(4, rownumber):
    for colIndex in range(0, colnumer):
        fieldItem = fieldInfo[colIndex]
        cell = booksheet.cell(rowIndex, colIndex)
        if colIndex == 0:
            id = int(cell.value)
            OutData[id] = {}
            OutData[id]['id'] = id
        elif cell.ctype == 1 or cell.ctype == 2:
            try:
                parseCell(fieldItem, cell)
            except:
                continue
            else:
                OutData[id][fieldItem.name] = parseCell(fieldItem, cell)
            

jsonOut = json.dumps(OutData,sort_keys=True, indent=4 ,separators=(',', ': '))
fileOutput.write(jsonOut)
print(json.dumps(jsonOut))


