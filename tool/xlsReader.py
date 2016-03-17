# encoding=utf-8
import itertools
import xlrd

table = None

def initSheet(fileName, sheet):
    global table
    book = xlrd.open_workbook(fileName)
    table = book.sheet_by_index(sheet - 1)
    
    

def getCellValue(row , col):
    cell = table.cell(row , col)
    return cell.value

'''
根据/生成各种组合数据
'''
def genCellValues(begin_cell, end_cell):
    s1 = begin_cell[0] - 1
    s2 = charToNum(begin_cell[1])
    e1 = end_cell[0]
    e2 = charToNum(end_cell[1]) + 1
    
    statList = []
    print (s1, s2, e1, e2)
    for a in range(s1 , e1):
        statNode = []
        for b in range(s2, e2):
            statNode.append(getCellValue(a, b))
        statList.append(statNode)
    for statNode in statList:
        l1 = []
        l2 = []
        for s in statNode:
            if '/' in s:
                l2.append(s)
            else:
                if s.strip() != "":
                    l1.append(s)
        
        if len(l2) == 0:
            print "`".join(l1)
        elif len(l2) == 1:
            l3 = l2[0].split("/")
            l4 = itertools.product(["`".join(l1)], l3)
            l5 = ["`".join(x) for x in list(l4)]
            for x in l5:
                print x
        elif len(l2) > 1:
            l3 = []
            for s in l2:
                l3.append(s.split('/'))
            l4 = []
            l4 = itertools.product(*l3)
            l5 = ["`".join(x) for x in list(l4)]
            l6 = itertools.product(["`".join(l1)], l5)
            l7 = ["`".join(x) for x in list(l6)]
            for x in l7:
                print x

def charToNum(c):
    num = 0
    if c.isalpha():
        num = ord(c)
    else:
        print c, '只能是字母'
        return None
    if num < 65 or num > 90:
        print "超出范围"
        return None
    return num - 65

if __name__ == '__main__':
    filePath = r"E:\Code\PycharmProjects\StatTest\xls\0317.xls"
    sheet = 1
    initSheet(filePath, sheet)
    cell1 = (12, "D")
    cell2 = (26, "G")
    genCellValues(cell1, cell2)
#     print charToNum("Z")
