# -*- coding: UTF-8 -*- 
# 作者：林培华
# 策划的配置表格转为json格式存储
# 后续可能的功能转为protobuf二进制文件同时生成.proto 后续可能的功能转为protobuf二进制文件同时生成对应的C#脚本
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import xlrd
import json
import re
import os

def _Excel2Json(filePath):
    class FieldItem:
        def __init__(self):
            self.name = ""
            self.targets = "a"
            self.type = ""
            
    # 获取文件名和拓展名
    dirPath, fileName = os.path.split(filePath)
    shotName, extension = os.path.splitext(fileName)

    fileOutput = open('%s/json/%s.json' % (dirPath, shotName), 'w')

    workbook = xlrd.open_workbook(filePath)
    booksheets = workbook.sheets()
    if len(booksheets) <= 0:
        print('%s has not sheets' % filePath)
    booksheet = booksheets[0]

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
            return num
    # 针对读出来的是102010101.0
    def myint(num):
        return int(myStr(float(str(num))))

    def mySplite(value, seperate):
        value = str(value).replace(' ', '')
        if value.find(seperate) == -1:
            return [value]
        else:
            return value.split(seperate)

    VectorPatten = re.compile('vector<(.*?)>', re.DOTALL | re.IGNORECASE)
    MapPatten = re.compile('map<(.*?),(.*?)>', re.DOTALL | re.IGNORECASE)

    def parseCell(fieldItem, cell):
        if cell.value == None or str(cell.value).replace(' ','') == '':
            return None
        def _innerParse(subType, value):
            try:
                if subType == 'string':
                    return myStr(value)
                elif subType == 'int':
                    return myint(value)
                elif subType == 'float':
                    return float(value)
            except:
                print(fieldItem.name + ' ' +  subType + ' _innerParse ' + str(value) + ' #')   
                if subType == 'string':
                    return ''
                elif subType == 'int':
                    return 0
                elif subType == 'float':
                    return 0.0

        fieldType = fieldItem.type
        
        m = VectorPatten.search(fieldType)
        if m != None:
            subType = m.group(1)
            split = mySplite(cell.value, '|')
            result = []
            for i in split:
                result.append(_innerParse(subType, i))

            return result
        m = MapPatten.search(fieldType)
        if m != None:
            subType1 = m.group(1)
            subType2 = m.group(2)
            result = {}
            split = mySplite(cell.value, '|')
            for i in split:
                tmpsplit = mySplite(i, ':')
                result[_innerParse(subType1, tmpsplit[0])] = _innerParse(subType2, tmpsplit[1])
            return result
        else:
            return _innerParse(fieldType, cell.value)


    OutData = {}
    for rowIndex in range(4, rownumber):
        for colIndex in range(0, colnumer):
            fieldItem = fieldInfo[colIndex]
            cell = booksheet.cell(rowIndex, colIndex)
            if cell.ctype == 1 or cell.ctype == 2:
                if colIndex == 0:
                    if cell.ctype == 2:
                        id = int(cell.value)
                        OutData[id] = {}
                        OutData[id]['id'] = id
                    else:
                        id = myStr(cell.value)
                        OutData[id] = {}
                        OutData[id]['id'] = id

                else: 
                    if fieldItem.name == None or fieldItem.name == '':
                        continue
                    # 这里由于可能存在' a'以及' c'所以改成了find
                    if fieldItem.targets == None or fieldItem.targets == '' or (fieldItem.targets.lower().find('a') == -1 and fieldItem.targets.lower().find('c') == -1):
                        # print(fieldItem.name + ' ' + fieldItem.targets.lower())
                        continue
                    # try:
                        # parseCell(fieldItem, cell)
                    # except:
                        # continue
                    # else:
                    result = parseCell(fieldItem, cell)
                    if result != None:
                        OutData[id][fieldItem.name] = result
                        
                

    jsonOut = json.dumps(OutData,sort_keys=True, indent=4, ensure_ascii=False, separators=(',', ': '))
    fileOutput.write(jsonOut)
    fileOutput.close()

rootdir = sys.path[0] + '/Test'
for filePath in os.listdir(rootdir):
    if filePath.endswith('.xlsx') and filePath.find('~$') == -1:
        print(os.path.join(rootdir, filePath))
        _Excel2Json(os.path.join(rootdir, filePath))

# print(json.dumps(jsonOut))


