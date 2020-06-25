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

#To calculate the P value
def PFunction(Z):
    p = 0.3275911
    a1 = 0.254829592
    a2 = -0.2884496736
    a3 = 1.421413741
    a4 = -1.453152027
    a5 = 1.061405429
    
    sign = 0
    if(Z < 0):
        sign = -1
    else:
        sign = 1
    
    x = abs(Z)/math.sqrt(2)
    t = 1/(1+(p*x))
    erf = 1-(((((a5 * t + a4)*t)+a3)*t + a2)*t + a1) * t * math.exp(-x*x);
    return 0.5 * (1 + sign * erf)

#To calculate the Z value (Z test)
def ZTest(r1a2a,r1a2b,r2a2b,N):
    Z1a2a = 1/2*math.log((1+r1a2a)/(1-r1a2a))
    Z1a2b = 1/2*math.log((1+r1a2b)/(1-r1a2b))
    
    rm2 = ((r1a2a*r1a2a)+(r1a2b*r1a2b))/2
    f = (1-r2a2b)/(2*(1-rm2))
    h = (1-(f*rm2))/(1-rm2)
    
    a = (Z1a2a - Z1a2b)
    b = math.sqrt(N - 3)/(2*(1-r2a2b)*h)
    Z = a*b
    return Z
    

TotalCount = 0
#These lists contain the entire data for all the windows (i.e all the P values)
PList_10 = []
PList_227 = []
PList_300 = []

#List holding all the unique dates in the two weeks excluding saturday and sunday
defList = [datetime.datetime(year = 2013, month = 2, day = 4, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 5, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 6, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 7, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 8, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 11, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 12, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 13, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 14, hour = 8),
                   datetime.datetime(year = 2013, month = 2, day = 15, hour = 8)]

#List of file names
XlList1 = ['ajb9b3','ajdqnf','amsh7c','bay6g7','cjdk88','ctkbn4',
           'dlokv6','dsdxd8','dtqb7d','elnfhc','habc8','hemmb2',
           'hmnhx4','hnldm2','jac345','jahq76','jdf5g4','jfk535',
           'jhgn9c','jkn7fc','jlgwxc','jlkm94','jmd5h8','jrrgxf',
           'kab5zb','kah89f','kdcncc','kedgm9','krh5f5','lms8m4',
           'lpnb79','mjpg95','mmg9n6','mtf7w2','mwwcn6','ngtp8b',
           'nhm2p3','njd5zf','nm8c7','oro26b','pecxc9','pt7td',
           'rrg3md','sanfz9','sbpc42','smdvw5','tah5z2','tcdzb7',
           'tjdpzd','tlgdq9','vvmmvd','wahbf','wgocnc','zws2p8']
#First main for loop which takes one user file and compares it with every other user.
for i in range(0,len(XlList1)):
    loc = r'D:\InfoSecurity\Information Security _ Privacy Material-20200409T150945Z-001\Information Security _ Privacy Material'+"\\"+XlList1[i]+".xlsx"
    List_P_10 = [] #Lists for 3 windows (10,227,300)
    List_P_227 = []
    List_P_300 = []
    wb = open_workbook(loc)
    sheet = wb.sheet_by_index(0) 
    sheet.cell_value(0, 0) 
    dateList = []
    docktetsList = []
    durationList = []
    OperDList = []
    day = 0
    dayStart = 1
    #This loop is used to filter the data(Data Preprocessing) 
    for i in range(1,sheet.nrows):
        X = int(sheet.cell_value(i,5)/1000) #Real first packet column
        D = sheet.cell_value(i,9) #Duration column
        O = sheet.cell_value(i,3) #Octets column
        Y = datetime.datetime.fromtimestamp(X) #Converting epoch time to datetime
        if(datetime.date(Y.year, Y.month,1).weekday() == 0):
            day = 12
        elif(datetime.date(Y.year, Y.month,1).weekday() == 1):
            day = 11
        elif(datetime.date(Y.year, Y.month,1).weekday() == 2):
            day = 10
        elif(datetime.date(Y.year, Y.month,1).weekday() == 3):
            day = 9
        elif(datetime.date(Y.year, Y.month,1).weekday() == 4):
            day = 15
            dayStart = 4
        elif(datetime.date(Y.year, Y.month,1).weekday() == 5):
            day = 14
        else:
            day = 13
        #Condition to filter all the unnecessary data.
        if(Y.time() >= datetime.time(8,0,0) and Y.time() <= datetime.time(17,0,0) and Y.date() >= datetime.date(Y.year,Y.month,dayStart) and Y.date() <= datetime.date(Y.year,Y.month,day) and Y.weekday() != 5 and Y.weekday() != 6 and D != 0):
            OperD = O/D #calculating octets/duration
            OperDList.append(OperD)
            dateList.append(Y)
            durationList.append(D)
            docktetsList.append(O)
                        
    
    #10sec window lists for week 1 and week 2
    Sheet_1Week_1List_10 = []
    Sheet_1Week_2List_10 = []
        
    #Loop to calculating octets/duration averages in several 10sec duration slots
    for i in range(0,len(defList)):
        T2 = defList[i]
        T3 = T2.replace(second = 10, hour = 8, minute = 0)
        for j in range(0,3240):
            docktetsAve = 0  
            count = 0
            for z in range(0,len(dateList)):
                if(dateList[z] >= T2 and dateList[z] < T3):
                    docktetsAve = docktetsAve + OperDList[z]
                    count = count + 1
        
            if(count == 0):
                count = 1
            if(T2 < datetime.datetime(year = 2013, month = 2, day = 8, hour = 18)):
                Sheet_1Week_1List_10.append(docktetsAve/count)
            else:
                Sheet_1Week_2List_10.append(docktetsAve/count)
            
            T2 = T2 + datetime.timedelta(seconds = 10)
            T3 = T3 + datetime.timedelta(seconds = 10)
            
            
    #227sec window lists for week 1 and week 2
    Sheet_1Week_1List_227 = []
    Sheet_1Week_2List_227 = []
        
    #Loop to calculating octets/duration averages in several 227sec duration slots    
    for i in range(0,len(defList)):
        T2 = defList[i]
        T3 = T2.replace(second = 47, hour = 8, minute = 3)
        for j in range(0,143):
            docktetsAve = 0  
            count = 0
            for z in range(0,len(dateList)):
                if(dateList[z] >= T2 and dateList[z] < T3):
                    docktetsAve = docktetsAve + OperDList[z]
                    count = count + 1
        
            if(count == 0):
                count = 1
            if(T2 < datetime.datetime(year = 2013, month = 2, day = 8, hour = 18)):
                Sheet_1Week_1List_227.append(docktetsAve/count)
            else:
                Sheet_1Week_2List_227.append(docktetsAve/count)
            
            T2 = T2 + datetime.timedelta(seconds = 227)
            T3 = T3 + datetime.timedelta(seconds = 227)
    #300sec window lists for week 1 and week 2            
    Sheet_1Week_1List_300 = []
    Sheet_1Week_2List_300 = []
        
    #Loop to calculating octets/duration averages in several 300sec duration slots    
    for i in range(0,len(defList)):
        T2 = defList[i]
        T3 = T2.replace(second = 0, hour = 8, minute = 5)
        for j in range(0,108):
            docktetsAve = 0  
            count = 0
            for z in range(0,len(dateList)):
                if(dateList[z] >= T2 and dateList[z] < T3):
                    docktetsAve = docktetsAve + OperDList[z]
                    count = count + 1
        
            if(count == 0):
                count = 1
            if(T2 < datetime.datetime(year = 2013, month = 2, day = 8, hour = 18)):
                Sheet_1Week_1List_300.append(docktetsAve/count)
            else:
                Sheet_1Week_2List_300.append(docktetsAve/count)
            
            T2 = T2 + datetime.timedelta(seconds = 300)
            T3 = T3 + datetime.timedelta(seconds = 300)
                        
            
    #Spearman Correlation for all the 3 windows.     
    r1a2a_10 = scipy.stats.spearmanr(Sheet_1Week_1List_10,Sheet_1Week_2List_10,nan_policy='propagate')[0]
    r1a2a_227 = scipy.stats.spearmanr(Sheet_1Week_1List_227,Sheet_1Week_2List_227,nan_policy='propagate')[0]
    r1a2a_300 = scipy.stats.spearmanr(Sheet_1Week_1List_300,Sheet_1Week_2List_300,nan_policy='propagate')[0]

    if(r1a2a_10 == 1):
        r1a2a_10 = 0.99
    if(math.isnan(r1a2a_10)):
        r1a2a_10 = 0       
    if(r1a2a_227 == 1):
        r1a2a_227 = 0.99
    if(math.isnan(r1a2a_227)):
        r1a2a_227 = 0
    if(r1a2a_300 == 1):
        r1a2a_300 = 0.99
    if(math.isnan(r1a2a_300)):
        r1a2a_300 = 0 
    # Two main for loop which iterates through each user excel file one after the other. 
    for i in range(0,len(XlList1)):
        loc1 = (r'D:\InfoSecurity\Information Security _ Privacy Material-20200409T150945Z-001\Information Security _ Privacy Material'+"\\"+XlList1[i]+".xlsx")
        

        wb1 = open_workbook(loc1)
        sheet1 = wb1.sheet_by_index(0)
        dateList1 = []
        docktetsList1 = []
        durationList1 = []
        OperDList1 = []

        
        
        day1 = 0
        dayStart1 = 1      
        #This loop is used to filter the data(Data Preprocessing) 
        for i in range(1,sheet1.nrows):
            X1 = int(sheet1.cell_value(i,5)/1000) #Real first packet
            D1 = sheet1.cell_value(i,9) #Duration
            O1 = sheet1.cell_value(i,3) #Octets
            
            Y1 = datetime.datetime.fromtimestamp(X1) #Converting Epoch time to datetime
            if(datetime.date(Y1.year, Y1.month,1).weekday() == 0):
                day = 12
            elif(datetime.date(Y1.year, Y1.month,1).weekday() == 1):
                day = 11
            elif(datetime.date(Y1.year, Y1.month,1).weekday() == 2):
                day = 10
            elif(datetime.date(Y1.year, Y1.month,1).weekday() == 3):
                day = 9
            elif(datetime.date(Y1.year, Y1.month,1).weekday() == 4):
                day = 15
                dayStart = 4
            elif(datetime.date(Y1.year, Y1.month,1).weekday() == 5):
                day = 14
            else:
                day = 13
            if(Y1.time() >= datetime.time(8,0,0) and Y1.time() <= datetime.time(17,0,0) and Y1.date() >= datetime.date(Y1.year,Y1.month,dayStart1) and Y1.date() <= datetime.date(Y1.year,Y1.month,day1) and Y1.weekday() != 5 and Y1.weekday() != 6 and D1 != 0):
                OperD1 = O1/D1 #Octets/duration
                OperDList1.append(OperD1)
                dateList1.append(Y1)
                durationList1.append(D1)
                docktetsList1.append(O1)        
        
        
        
        #10sec window lists for week 1 and week 2
        Sheet_2Week_1List_10 = []
        Sheet_2Week_2List_10 = []
        #Loop to calculating octets/duration averages in several 10sec duration slots    
        for i in range(0,len(defList)):
            T4 = defList[i]
            T5 = T4.replace(second = 10, hour = 8, minute = 0)
            for j in range(0,3240):
                docktetsAve1 = 0
                count1 = 0
                for z in range(0,len(dateList1)):
                    if(dateList1[z] >= T4 and dateList1[z] < T5):
                        docktetsAve1 = docktetsAve1 + OperDList1[z]
                        count1 = count1 + 1
                if(count1 == 0):
                    count1 = 1
                if(T4 < datetime.datetime(year = 2013, month = 2, day = 8, hour = 18)):
                    Sheet_2Week_1List_10.append(docktetsAve1/count1)
                else:
                    Sheet_2Week_2List_10.append(docktetsAve1/count1)
                    
                T4 = T4 + datetime.timedelta(seconds = 10)
                T5 = T5 + datetime.timedelta(seconds = 10)
        #227sec window lists for week 1 and week 2
        Sheet_2Week_1List_227 = []
        Sheet_2Week_2List_227 = []
        #Loop to calculating octets/duration averages in several 227sec duration slots
        for i in range(0,len(defList)):
            T4 = defList[i]
            T5 = T4.replace(second = 47, hour = 8, minute = 3)
            for j in range(0,143):
                docktetsAve1 = 0
                count1 = 0
                for z in range(0,len(dateList1)):
                    if(dateList1[z] >= T4 and dateList1[z] < T5):
                        docktetsAve1 = docktetsAve1 + OperDList1[z]
                        count1 = count1 + 1
                if(count1 == 0):
                    count1 = 1
                if(T4 < datetime.datetime(year = 2013, month = 2, day = 8, hour = 18)):
                    Sheet_2Week_1List_227.append(docktetsAve1/count1)
                else:
                    Sheet_2Week_2List_227.append(docktetsAve1/count1)
                    
                T4 = T4 + datetime.timedelta(seconds = 227)
                T5 = T5 + datetime.timedelta(seconds = 227)
        #300sec window lists for week 1 and week 2     
        Sheet_2Week_1List_300 = []
        Sheet_2Week_2List_300 = []
        #Loop to calculating octets/duration averages in several 300sec duration slots
        for i in range(0,len(defList)):
            T4 = defList[i]
            T5 = T4.replace(second = 0, hour = 8, minute = 5)
            for j in range(0,108):
                docktetsAve1 = 0
                count1 = 0
                for z in range(0,len(dateList1)):
                    if(dateList1[z] >= T4 and dateList1[z] < T5):
                        docktetsAve1 = docktetsAve1 + OperDList1[z]
                        count1 = count1 + 1
                if(count1 == 0):
                    count1 = 1
                if(T4 < datetime.datetime(year = 2013, month = 2, day = 8, hour = 18)):
                    Sheet_2Week_1List_300.append(docktetsAve1/count1)
                else:
                    Sheet_2Week_2List_300.append(docktetsAve1/count1)
                    
                T4 = T4 + datetime.timedelta(seconds = 300)
                T5 = T5 + datetime.timedelta(seconds = 300)
        #Spearman Correlation using scipy.stats
        r1a2b_10 = scipy.stats.spearmanr(Sheet_1Week_1List_10,Sheet_2Week_2List_10,nan_policy='propagate')[0]
        r2a2b_10 = scipy.stats.spearmanr(Sheet_1Week_2List_10,Sheet_2Week_2List_10,nan_policy='propagate')[0]
        
        r1a2b_227 = scipy.stats.spearmanr(Sheet_1Week_1List_227,Sheet_2Week_2List_227,nan_policy='propagate')[0]
        r2a2b_227 = scipy.stats.spearmanr(Sheet_1Week_2List_227,Sheet_2Week_2List_227,nan_policy='propagate')[0]
        
        r1a2b_300 = scipy.stats.spearmanr(Sheet_1Week_1List_300,Sheet_2Week_2List_300,nan_policy='propagate')[0]
        r2a2b_300 = scipy.stats.spearmanr(Sheet_1Week_2List_300,Sheet_2Week_2List_300,nan_policy='propagate')[0]
        
        if(r1a2b_10 == 1):
            r1a2b_10 = 0.99
        if(math.isnan(r1a2b_10)):
            r1a2b_10 = 0
        if(r2a2b_10 == 1):
            r2a2b_10 = 0.99
        if(math.isnan(r2a2b_10)):
            r2a2b_10 = 0
        if(r1a2b_227 == 1):
            r1a2b_227 = 0.99
        if(math.isnan(r1a2b_227)):
            r1a2b_227 = 0
        if(r2a2b_227 == 1):
            r2a2b_227 = 0.99
        if(math.isnan(r2a2b_227)):
            r2a2b_227 = 0
        if(r1a2b_300 == 1):
            r1a2b_300 = 0.99
        if(math.isnan(r1a2b_300)):
            r1a2b_300 = 0
        if(r2a2b_300 == 1):
            r2a2b_300 = 0.99
        if(math.isnan(r2a2b_300)):
            r2a2b_300 = 0
        
        #Calculating the Z value, P value 
        Z_10 = ZTest(r1a2a_10,r1a2b_10,r2a2b_10,16200)
        P_10 = PFunction(Z_10)
        TotalCount = TotalCount + 1
        List_P_10.append(P_10) # appending all the P row wise.
        
        Z_227 = ZTest(r1a2a_227,r1a2b_227,r2a2b_227,715)
        P_227 = PFunction(Z_227)
        
        List_P_227.append(P_227)
        
        Z_300 = ZTest(r1a2a_300,r1a2b_300,r2a2b_300,540)
        P_300 = PFunction(Z_300)
        
        List_P_300.append(P_300)
        
    #Main lists containing all the P values for windows 10sec, 227sec, 300sec
    PList_10.append(List_P_10)
    PList_227.append(List_P_227)
    PList_300.append(List_P_300)

    
#These are used to write the P values present in the lists to Excel files.
with xlsxwriter.Workbook('10sec_window.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    for row_num, data in enumerate(PList_10):
        worksheet.write_row(row_num, 0, data)
        
with xlsxwriter.Workbook('227sec_window.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    for row_num, data in enumerate(PList_227):
        worksheet.write_row(row_num, 0, data)
        
with xlsxwriter.Workbook('300sec_window.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    for row_num, data in enumerate(PList_300):
        worksheet.write_row(row_num, 0, data)