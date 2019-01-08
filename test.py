#python3

import xlrd
import pprint
import datetime
import re



def getfilename(pz,server):
    date = str(datetime.date.today().__format__('%Y%m%d'))
    print(date)
    filename = dict()
    j = 0
    for i in pz:
        for k in server:
            j += 1
            filename[j]= k + date + i + "tick.xls"
    return filename
        #filename = "107-" + date + i + "tick.xls"
        #filename = date + i + "tick.xls"


def calxslcell():

    wb = xlrd.open_workbook(r'D:\\Documents\\fabric\\爬虫\\228-20180908白糖tick.xls')
    sheet =wb.sheet_by_name('Sheet1')

    while  1:
        cell = wb.cell(row=num, column=1).value
        if cell:
            num = num + 1
        else:
            print(num)
            exit()

    return  num

server = ['','227-','107-']
pz = ['白糖','PTA','郑醇','白糖主连','PTA主连','郑醇主连']

num = calxslcell()
print(num)
