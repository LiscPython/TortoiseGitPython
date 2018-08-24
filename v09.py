#python3

import openpyxl
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


def calxslcell(filename):
    print("open {}".format(filename))
    wb = openpyxl.load_workbook(filename)
    sheet = wb.get_sheet_by_name('sheet1')
    row = sheet.get_hightest_row() - 1
    return  row

server = ['','227-','107-']
pz = ['白糖','PTA','郑醇','白糖主连','PTA主连','郑醇主连']
filename=getfilename(pz,server)
filename = list(filename.values())
print(filename)
data = dict()

for i in filename:
        #num = calxslcell(i)
    print(i)


