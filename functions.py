import time

import daxDownloader as td
import daxQueries as daxQ
import xlwings as xw
import datetime
import pandas as pd
import numpy as np
import os
#from datetime import datetime, timedelta, date

dax_query_list = [daxQ.grading,daxQ.sku_plu,daxQ.md,daxQ.plu_available,
                     daxQ.promo_reg,daxQ.promo_tv,daxQ.prh_data,daxQ.perf_dep,daxQ.pcal,daxQ.prh]

def runDaxQueries(startWeek,endWeek,dep,dax_query_list,path):

    wb = xw.Book.caller()
    ws = wb.sheets[0]

    ws.range("G3:H20").clear_contents() # clear cells with times
    ws.range("K3:L20").clear_contents() # clear cells with info about the files

    for index, query in enumerate(dax_query_list):

        ws.range("C20").value = dep
        ws.range("C21").value = query.__name__
        ws.range("C23").value = "Not Ready"
        ws.range((index + 3, 7)).value = datetime.datetime.now()

        runDaxQuery(startWeek,endWeek,dep,query,path)

        ws.range((index + 3, 8)).value = datetime.datetime.now()
        ws.range((index + 3, 11)).value = query.__name__
        ws.range("C23").value = "Ready"

def runDaxQuery(startWeek,endWeek,dep,dax_query,path):
    df = td.dataFrameFromTabular(dax_query(startWeek, endWeek, dep))
    q_name = dax_query.__name__ + '.csv'
    df.to_csv(path + q_name, index=False)

# def runDaxInventory(startWeek,endWeek,dep,path):
#     wb = xw.Book.caller()
#     ws = wb.sheets[0]
#
#     weeksList = []
#
#     wList = pd.read_csv(path + "pcal.csv")
#     for week in wList.iloc[:, 5]:
#         weeksList.append(week)
#     del wList
#
#     ws.range("C20").value = dep
#     ws.range("C21").value = "inventory"
#     ws.range("C23").value = "Not Ready"
#
#     ws.range("G15").value = datetime.datetime.now()
#     col = ['SKU PLU', 'SKU Colour', 'STR Number', 'Pl Year', 'Pl Week', 'SalesU', 'SalesV', 'SalesM', 'CSOHU', 'CSOHV']
#     inventory_df = pd.DataFrame()
#     for week in weeksList:
#         ws.range("C20").value = dep
#         ws.range("C21").value = "inventory"
#         ws.range("C22").value = week
#
#         df = td.dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, dep))
#         inventory_df = inventory_df.append(df)
#         inventory_df.to_csv(path + "inventory.csv", index=False)
#
#     inventory_df.columns = col
#     inventory_df.to_csv(path + "inventory.csv", index=False)
#
#     ws.range("H15").value = datetime.datetime.now()
#     ws.range("K15").value = "inventory"
#     ws.range("C20:C22").clear_contents()  # clear cells with Information's table
#     ws.range("C23").value = "Ready"

def weeksCalculation(startWeek, endWeek, maxWeekNo):
    """
    Comments:
    NOT USED
        W raporcie pcal znajdują się wszystkie tygodnie.
        :param startWeek:
        :param endWeek:
        :param maxWeekNo: number of weeks in the year from startWeek variable
        :return:
    """

    list_of_weeks = []
    if int(str(startWeek)[-2:]) > int(str(endWeek)[-2:]):
        weekNumbers = maxWeekNo - int(str(startWeek)[-2:]) + int(str(endWeek)[-2:])
    else:
        weekNumbers = endWeek - startWeek

    weekRatio = 0
    for weeks in range(0,weekNumbers + 1):
        week = startWeek + weeks - weekRatio

        if int(str(week)[-2:]) == 52:
            year = int(str(week)[:4]) + 1
            startWeek = int(str(year) + "01")
            weekRatio = weeks + 1

        list_of_weeks.append(week)
        #print(week)
    return list_of_weeks

def getExcelData():
    wb = xw.Book.caller()
    ws = wb.sheets[0]
    dep = ws["C4"].value
    startWeek = ws["T3"].value
    endWeek = ws["T4"].value
    path = ws["C5"].value + "\\"
    return dep, startWeek,endWeek, path

def loadParameters():
    dir_ = r'\\10.2.5.140\zasoby\Planowanie\PERSONAL FOLDERS\Mariusz Borycki\HitMidKit_DataBase'
    f_name = '_main files\HitMidKit_parameters.xlsm'
    directory = dir_ + "\\" + f_name
    df = pd.read_excel(directory, sheet_name='py_inputs', skiprows=1)
    df = df[['Variables Header', 'Chosen Variables']]

    df2 = pd.read_excel(directory, sheet_name='QueryPar', skiprows=1)
    df2 = df2[['Chosen Hierarchy', 'Index']]

    startWeek = int(df[df['Variables Header'] == 'Week From']['Chosen Variables'])
    endWeek = int(df[df['Variables Header'] == 'Week To']['Chosen Variables'])
    minPar = int(df[df['Variables Header'] == 'Minimum Sales Units']['Chosen Variables'])
    dep = df[df['Variables Header'] == 'Departament']['Chosen Variables'].item()
    departmentList = list(df2['Chosen Hierarchy'].unique())
    departmentList = [x for x in departmentList if str(x) != 'nan']
    hierIdx = df[df['Variables Header'] == 'Hierarchy']['Chosen Variables'].item()
    indexList = list(df2['Index'].unique())
    indexList = [x for x in indexList if str(x) != 'nan']

    path = str(df[df['Variables Header'] == 'Path']['Chosen Variables'].item()) + "\\"

    if dep == 0:
        for d in departmentList:
            dep = departmentList
            hierIdx = indexList
        else:
            dep = dep
            hierIdx = hierIdx

    return dep, startWeek, endWeek, minPar, path, hierIdx

def loadInventory(startWeek,endWeek, minPar, dep, path, hierIdx):
    # showStatus(path, 1, dep, hierIdx, week=None) # 1 status means: bu#sy

    if type(dep) == list:
        for dep_index, dep_name in enumerate(dep):
            weeksList = []
            wList = td.dataFrameFromTabular(daxQ.pcal(startWeek, endWeek, dep_name))
            for week in wList.iloc[:, 5]:
                weeksList.append(week)
            del wList

            col = ['SKU PLU', 'SKU Colour', 'STR Number', 'Pl Year', 'Pl Week', 'SalesU', 'SalesV', 'SalesM', 'CSOHU', 'CSOHV']
            inventory_df = pd.DataFrame()
            for week in weeksList:
                # Need it for saving informations into .log file
                startQuery = datetime.datetime.now()
                startQuery = startQuery.strftime("%H:%M:%S")  # startQuery.strftime("%d/%m/%Y %H:%M:%S")
                showStatus(path, 1, dep, hierIdx, week)
                # print(f"\n-----\nTemporary (dep list-YES):\nStart Week: {startWeek}\nEnd Week: {endWeek}\nWeek: {week}\nDepartment: {dep_name}\nIndex: {hierIdx[dep_index]}")
                try:
                    df = td.dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, minPar, dep_name))
                except:
                    try:
                        time.sleep(60)
                        df = td.dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, minPar, dep_name))
                    except Exception as ex:
                        file = open(path + "ConnectionIssue.txt", "w")
                        file.write(f"Connection error:\n{ex}\n")
                        file.close()
                        return None

                df['total'] = np.where(df.iloc[:, 5:].sum(axis=1) > 0, 1, 0)
                df = df[df.total > 0]
                df.drop(columns={'total'}, inplace=True)

                inventory_df = inventory_df.append(df)
                #inventory_df.to_csv(path + f"{hierIdx[dep_index]}_HitMidKit.csv", index=False)

                endQuery = datetime.datetime.now()
                endQuery = endQuery.strftime("%H:%M:%S")
                saveLog(startQuery, endQuery, dep_name, hierIdx[dep_index], week, path)

            inventory_df.columns = col
            inventory_df.to_csv(path + f"{hierIdx[dep_index]}_HitMidKit.csv", index=False)
    else:
        weeksList = []
        wList = td.dataFrameFromTabular(daxQ.pcal(startWeek, endWeek, dep))
        for week in wList.iloc[:, 5]:
            weeksList.append(week)
        del wList

        col = ['SKU PLU','SKU Colour','STR Number','Pl Year','Pl Week','SalesU','SalesV','SalesM','CSOHU','CSOHV']
        inventory_df = pd.DataFrame()
        for week in weeksList:
            # Need it for saving informations into .log file
            startQuery = datetime.datetime.now()
            startQuery = startQuery.strftime("%H:%M:%S")  # startQuery.strftime("%d/%m/%Y %H:%M:%S")
            showStatus(path, 1, dep, hierIdx, week)
            # print(f"\n-----\nTemporary (dep list-NO):\nStart Week: {startWeek}\nEnd Week: {endWeek}\nWeek: {week}\nDepartment: {dep}\nIndex: {hierIdx}")
            try:
                df = td.dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, minPar, dep))
            except:
                try:
                    time.sleep(60)
                    df = td.dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, minPar, dep))
                except Exception as ex:
                    file = open(path + "ConnectionIssue.txt", "w")
                    file.write(f"Connection error:\n{ex}\n")
                    file.close()
                    return None
            df['total'] = np.where(df.iloc[:, 5:].sum(axis=1) > 0, 1, 0)
            df = df[df.total > 0]
            df.drop(columns={'total'}, inplace=True)

            inventory_df = inventory_df.append(df)
            #inventory_df.to_csv(path + f"{hierIdx}_HitMidKit.csv", index=False)

            endQuery = datetime.datetime.now()
            endQuery = endQuery.strftime("%H:%M:%S")
            saveLog(startQuery, endQuery, dep, hierIdx, week, path)

        inventory_df.columns = col
        inventory_df.to_csv(path + f"{hierIdx}_HitMidKit.csv", index=False)

    showStatus(path, 0, dep=None, hierIdx=None, week=None)  # 0 status means: NOT busy. Inventory file is ready

def showStatus(path,y_n,dep,hier,week):
    if y_n == 1:
        if os.path.exists(path + "Ready.txt"):
            os.remove(path + "Ready.txt")

        file = open(path + "Not Ready.txt", "w")
        file.write(f"Data are still downloading...\n\n"
                   f"Department: {dep}\n"
                   f"Hierarchy: {hier}\n"
                   f"Week: {week}")
        file.close()
    else:
        if os.path.exists(path + "Not Ready.txt"):
            os.remove(path + "Not Ready.txt")

        file = open(path + "Ready.txt", "w")
        file.write("Data have been downloaded")

        file.close()

def saveLog(startQuery, endQuery, dep, hierIdx, week, path):
    today = datetime.datetime.today()
    today = today.strftime("%d/%m/%Y")
    queryTime = str(datetime.datetime.strptime(endQuery, "%H:%M:%S") - datetime.datetime.strptime(startQuery, "%H:%M:%S"))
    logList = [today, dep, str(hierIdx), str(week), startQuery, endQuery, queryTime]

    with open(path + '../Logs/Log_detailed_HitMidKit.log', 'a+') as log_file:
        log_file.write("\n")
        log_file.writelines(';'.join(logList))