import time; 
import pandas as pd
import xlrd;
import xlwt;
import datetime;
from xlrd import open_workbook
from xlutils.copy import copy
import scipy.stats
import math
import xlsxwriter


XlList1 = ['10sec_window','227sec_window','300sec_window']
XlList2 = ['10sec_Y_or_N.xlsx','227sec_Y_or_N.xlsx','300sec_Y_or_N.xlsx']
for i in range(0,len(XlList1)):
    count_Y = 0
    count_N = 0
    #loc = r'C:\Users\DayaRani\.spyder-py3'+"\\"+XlList1[i]+".xlsx"
    loc = r'C:\Users\DayaRani\Desktop'+"\\"+XlList1[i]+".xlsx"
    #loc = r'C:\Users\DayaRani\Desktop\Infosec_Project.xlsx'
    List_P = []
    wb = open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    for j in range(0,sheet.nrows):
        row_List = []
        for z in range(0,sheet.nrows):
            if(sheet.cell_value(j,z) <= 0.05):
                row_List.append("Yes")
                count_Y = count_Y + 1
            else:
                row_List.append("No")
                count_N = count_N + 1
        List_P.append(row_List)
    print(count_Y)
    print(count_N)
    with xlsxwriter.Workbook(XlList2[i]) as workbook:
        worksheet = workbook.add_worksheet()

        for row_num, data in enumerate(List_P):
            worksheet.write_row(row_num, 0, data)
                
        
        