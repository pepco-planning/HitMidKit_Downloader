import pandas as pd
import numpy as np
import os
import daxQueries as daxQ
import datetime
import time
import xlwings as xw
from pyadomd import Pyadomd
from sys import path
path.append("c:\Mariusz\MyProjects\HitMidKit_Downloader\dll")
path.append(os.getcwd() + '\\dll')

def dataFrameFromTabular(query):
    """
    # w path.append jest istotna ścieżka
    # muszą się w niej znajdować 2 pliki:
    # 1. Microsoft.AnalysisServices.AdomdClient.dll
    # 2. Microsoft.AnalysisServices.dll
    # Plików szukaj w C:\Windows\assembly\GAC_MSIL\Microsoft.AnalysisServices.(nazwa pliku)\

    :param query: here are required dax queries we need to download data
    :return:
    """
    connStr = "Provider=MSOLAP;Data Source=LB-P-WE-AS;Catalog=PEPCODW"

    conn = Pyadomd(connStr)
    conn.open()
    cursor = conn.cursor()
    cursor.execute(query)
    col_names = [i[0] for i in cursor.description]
    cursor.arraysize = 5000
    df = pd.DataFrame(cursor.fetchall(), columns=col_names)
    conn.close()

    return df

dax_query_list = [daxQ.grading,daxQ.sku_plu,daxQ.md,daxQ.plu_available,
                     daxQ.promo_reg,daxQ.promo_tv,daxQ.prh_data,daxQ.perf_dep,daxQ.pcal,daxQ.prh]

def runDaxQueries(startWeek,endWeek,dep,dax_query_list,path):

    wb = xw.Book.caller()
    ws = wb.sheets["QueryPar2"]

    ws.range("G3:H20").clear_contents() # clear cells with times
    ws.range("K3:L20").clear_contents() # clear cells with info about the files

    for index, query in enumerate(dax_query_list):

        ws.range("C20").value = dep
        ws.range("C21").value = query.__name__
        ws.range("C23").value = "Not Ready"
        ws.range((index + 3, 7)).value = datetime.datetime.now()

        df = runDaxQuery(startWeek,endWeek,dep,query,path)
        q_name = query.__name__ + '.csv'
        df.to_csv(path + q_name.capitalize(), index=False)

        ws.range((index + 3, 8)).value = datetime.datetime.now()
        ws.range((index + 3, 11)).value = query.__name__
        ws.range("C23").value = "Ready"

def runDaxQueriesExe(startWeek,endWeek,dep,dax_query_list,path, hierIdx):
    if type(dep) == list:
        for index, query in enumerate(dax_query_list):
            df_temp = pd.DataFrame()
            for dep_index, dep_name in enumerate(dep):
                startQuery = datetime.datetime.now()
                startQuery = startQuery.strftime("%H:%M:%S")  # startQuery.strftime("%d/%m/%Y %H:%M:%S")
                showStatus(path, 1, dep_name, hierIdx[dep_index], week=None)
                try:
                    df = runDaxQuery(startWeek, endWeek, dep_name, query, path)
                except:
                    try:
                        time.sleep(60)
                        df = runDaxQuery(startWeek, endWeek, dep_name, query, path)
                    except Exception as ex:
                        file = open(path + "ConnectionIssue.txt", "w")
                        file.write(f"Connection error:\n{ex}\n")
                        file.close()
                        return None

                df['HierarchyIndex'] = hierIdx[dep_index]
                df_temp = df_temp.append(df)
                endQuery = datetime.datetime.now()
                endQuery = endQuery.strftime("%H:%M:%S")
                saveLog(startQuery, endQuery, dep_name, hierIdx[dep_index], week=None, path=path, query=query.__name__)

            q_name = query.__name__ + '.csv'
            df_temp.to_csv(path + q_name.capitalize(), index=False)
    else:
        for index, query in enumerate(dax_query_list):
            startQuery = datetime.datetime.now()
            startQuery = startQuery.strftime("%H:%M:%S")  # startQuery.strftime("%d/%m/%Y %H:%M:%S")
            showStatus(path, 1, dep, hierIdx, week=None)

            df_temp = pd.DataFrame()
            try:
                df = runDaxQuery(startWeek, endWeek, dep, query, path)
            except:
                try:
                    time.sleep(60)
                    df = runDaxQuery(startWeek, endWeek, dep, query, path)
                except Exception as ex:
                    file = open(path + "ConnectionIssue.txt", "w")
                    file.write(f"Connection error:\n{ex}\n")
                    file.close()
                    return None

            df['HierarchyIndex'] = hierIdx
            df_temp = df_temp.append(df)
            endQuery = datetime.datetime.now()
            endQuery = endQuery.strftime("%H:%M:%S")
            saveLog(startQuery, endQuery, dep, hierIdx, week=None, path=path, query=query.__name__)

        q_name = query.__name__ + '.csv'
        df_temp.to_csv(path + q_name.capitalize(), index=False)

    showStatus(path, 0, dep, hier=None, week=None) # 0 means Ready

def runDaxQuery(startWeek,endWeek,dep,dax_query,path):
    df = dataFrameFromTabular(dax_query(startWeek, endWeek, dep))

    return df

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

    # Downloader (QueryPar)
    ws = wb.sheets['QueryPar2'] # wb.sheets[3]
    dep = ws["C4"].value
    startWeek = ws["T3"].value
    endWeek = ws["T4"].value
    path = ws["C5"].value + "\\"

    # Calculations (py_inputs)
    ws2 = wb.sheets['py_inputs']
    fileName = ws2["O20"].value

    return dep, startWeek,endWeek, path, fileName

def getCalcParam(path,fileName):
    wb = xw.Book.caller()
    ws = wb.sheets["QueryPar2"]

    ws.range("C21").value = "parameters"
    ws.range("C23").value = "Not Ready"
    ws.range("G13").value = datetime.datetime.now()

    df = pd.read_excel(fileName, sheet_name='py_inputs', skiprows=1,
                       usecols=['Variables Header', 'Chosen Variables'],converters={'Variables Header':str,'Chosen Variables':str})
    df = df.dropna()
    df = df[16:]
    df.to_csv(path + "parameters.csv", index=False)

    ws.range("H13").value = datetime.datetime.now()
    ws.range("K13").value = "parameters"
    ws.range("C23").value = "Ready"

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
    reportType = df[df['Variables Header'] == 'Report Type']['Chosen Variables'].item()

    path = str(df[df['Variables Header'] == 'Path']['Chosen Variables'].item()) + "\\"

    if dep == 0:
        dep = departmentList
        hierIdx = indexList
    else:
        dep = dep
        hierIdx = hierIdx

    return dep, startWeek, endWeek, minPar, path, hierIdx, reportType

def loadInventory(startWeek,endWeek, minPar, dep, path, hierIdx):
    if type(dep) == list:
        for dep_index, dep_name in enumerate(dep):
            weeksList = []
            wList = dataFrameFromTabular(daxQ.pcal(startWeek, endWeek, dep_name))
            for week in wList.iloc[:, 5]:
                weeksList.append(week)
            del wList

            col = ['SKU PLU', 'SKU Colour', 'STR Number', 'Pl Year', 'Pl Week', 'SalesU', 'SalesV', 'SalesM', 'CSOHU', 'CSOHV']
            inventory_df = pd.DataFrame()
            for week in weeksList:
                startQuery = datetime.datetime.now()
                startQuery = startQuery.strftime("%H:%M:%S")  # startQuery.strftime("%d/%m/%Y %H:%M:%S")
                showStatus(path, 1, dep_name, hierIdx[dep_index], week)
                try:
                    df = dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, minPar, dep_name))
                except:
                    try:
                        time.sleep(60)
                        df = dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, minPar, dep_name))
                    except Exception as ex:
                        file = open(path + "ConnectionIssue.txt", "w")
                        file.write(f"Connection error:\n{ex}\n")
                        file.close()
                        return None

                df['total'] = np.where(df.iloc[:, 5:].sum(axis=1) > 0, 1, 0)
                df = df[df.total > 0]
                df.drop(columns={'total'}, inplace=True)

                inventory_df = inventory_df.append(df)
                endQuery = datetime.datetime.now()
                endQuery = endQuery.strftime("%H:%M:%S")
                saveLog(startQuery, endQuery, dep_name, hierIdx[dep_index], week, path, query='Inventory')

            inventory_df.columns = col
            inventory_df.to_csv(path + f"{hierIdx[dep_index]}_HitMidKit.csv", index=False)
    else:
        weeksList = []
        wList = dataFrameFromTabular(daxQ.pcal(startWeek, endWeek, dep))
        for week in wList.iloc[:, 5]:
            weeksList.append(week)
        del wList

        col = ['SKU PLU','SKU Colour','STR Number','Pl Year','Pl Week','SalesU','SalesV','SalesM','CSOHU','CSOHV']
        inventory_df = pd.DataFrame()
        for week in weeksList:
            startQuery = datetime.datetime.now()
            startQuery = startQuery.strftime("%H:%M:%S")  # startQuery.strftime("%d/%m/%Y %H:%M:%S")
            showStatus(path, 1, dep, hierIdx, week)
            try:
                df = dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, minPar, dep))
            except:
                try:
                    time.sleep(60)
                    df = dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, minPar, dep))
                except Exception as ex:
                    file = open(path + "ConnectionIssue.txt", "w")
                    file.write(f"Connection error:\n{ex}\n")
                    file.close()
                    return None
            df['total'] = np.where(df.iloc[:, 5:].sum(axis=1) > 0, 1, 0)
            df = df[df.total > 0]
            df.drop(columns={'total'}, inplace=True)
            inventory_df = inventory_df.append(df)

            endQuery = datetime.datetime.now()
            endQuery = endQuery.strftime("%H:%M:%S")
            saveLog(startQuery, endQuery, dep, hierIdx, week, path, query='Inventory')

        inventory_df.columns = col
        inventory_df.to_csv(path + f"{hierIdx}_HitMidKit.csv", index=False)

    showStatus(path, 0, dep=None, hier=None, week=None)  # 0 status means: NOT busy. Inventory file is ready

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

def saveLog(startQuery, endQuery, dep, hierIdx, week, path, query):
    today = datetime.datetime.today()
    today = today.strftime("%d/%m/%Y")
    queryTime = str(datetime.datetime.strptime(endQuery, "%H:%M:%S") - datetime.datetime.strptime(startQuery, "%H:%M:%S"))
    logList = [today, query, dep, str(hierIdx), str(week), startQuery, endQuery, queryTime]

    with open(path + '../Logs/Log_detailed_HitMidKit.log', 'a+') as log_file:
        log_file.write("\n")
        log_file.writelines(';'.join(logList))