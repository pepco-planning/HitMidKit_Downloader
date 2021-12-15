import daxDownloader as td
import daxQueries as daxQ
import xlwings as xw
import datetime
import pandas as pd
import os
import time

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

def runDaxInventory(startWeek,endWeek,dep,path):
    wb = xw.Book.caller()
    ws = wb.sheets[0]

    weeksList = []

    wList = pd.read_csv(path + "pcal.csv")
    for week in wList.iloc[:, 5]:
        weeksList.append(week)
    del wList

    ws.range("C20").value = dep
    ws.range("C21").value = "inventory"
    ws.range("C23").value = "Not Ready"

    ws.range("G15").value = datetime.datetime.now()
    col = ['SKU PLU', 'SKU Colour', 'STR Number', 'Pl Year', 'Pl Week', 'SalesU', 'SalesV', 'SalesM', 'CSOHU', 'CSOHV']
    inventory_df = pd.DataFrame()
    for week in weeksList:
        ws.range("C20").value = dep
        ws.range("C21").value = "inventory"
        ws.range("C22").value = week

        df = td.dataFrameFromTabular(daxQ.inventory(startWeek, endWeek, week, dep))
        inventory_df = inventory_df.append(df)
        inventory_df.to_csv(path + "inventory.csv", index=False)

    inventory_df.columns = col
    inventory_df.to_csv(path + "inventory.csv", index=False)

    ws.range("H15").value = datetime.datetime.now()
    ws.range("K15").value = "inventory"
    ws.range("C20:C22").clear_contents()  # clear cells with Information's table
    ws.range("C23").value = "Ready"

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

def xw_func():
    """
    Comments:
    NOT USED
        # A Function in xlwings (we can create own function there
        # def xw_func() DO NOT EXIST. I have created it just to make it easier to collapse
    """
    # @xw.func
    # def hello(name):
    #     return f"Hi {name}!"
    #
    # if __name__ == "__main__":
    #     xw.Book("hitmidkit.xlsm").set_mock_caller()
    #     main()

def getExcelData():
    wb = xw.Book.caller()
    ws = wb.sheets[0]
    dep = ws["C4"].value
    startWeek = ws["T3"].value
    endWeek = ws["T4"].value
    path = ws["C5"].value + "\\"
    return dep, startWeek,endWeek, path