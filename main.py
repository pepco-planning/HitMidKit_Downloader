import pandas as pd
import daxDownloader as td
import daxQueries as daxQ
import functions as f
import xlwings as xw
import datetime


dax_query_list = [daxQ.grading,daxQ.sku_plu,daxQ.md,daxQ.plu_available,
                     daxQ.promo_reg,daxQ.promo_tv,daxQ.prh_data,daxQ.perf_dep,daxQ.pcal,daxQ.prh]

def main():
    wb = xw.Book.caller()
    ws = wb.sheets[0]
    dep = ws["C4"].value
    startWeek = ws["Q3"].value
    endWeek = ws["Q4"].value
    path = ws["C5"].value + "\\"
    ws.range("G3:H20").clear_contents() # clear cells with times
    ws.range("K3:L20").clear_contents() # clear cells with info about the files

    dax_query_list = [daxQ.grading, daxQ.sku_plu, daxQ.md, daxQ.plu_available,
                      daxQ.promo_reg, daxQ.promo_tv, daxQ.prh_data, daxQ.perf_dep, daxQ.pcal, daxQ.prh]

    for index, query in enumerate(dax_query_list):
        ws.range("B20").value = f"Chosen Departament: {dep}\n " \
                                f"Downloading table: {query.__name__}"

        ws.range((index + 3, 7)).value = datetime.datetime.now()
        f.runDaxQuery(startWeek,endWeek,dep,query,path)
        ws.range((index + 3, 8)).value = datetime.datetime.now()
        ws.range((index + 3, 11)).value = query.__name__


    # Inventory
    weeksList = []
    wList = pd.read_csv("c:\Mariusz\MyProjects\HitMidKit\input data\pcal.csv")
    for week in wList.iloc[:, 5]:
        weeksList.append(week)
    del wList

    ws.range("B20").value = f"Chosen Departament: {dep}\n " \
                            f"Downloading table: inventory"
    ws.range("G15").value = datetime.datetime.now()
    col = ['SKU PLU', 'SKU Colour', 'STR Number', 'Pl Year', 'Pl Week', 'SalesU', 'SalesV', 'SalesM', 'CSOHU', 'CSOHV']
    inventory_df = pd.DataFrame()
    for week in weeksList:
        ws.range("B20").value = f"Chosen Departament: {dep}\n " \
                                f"Downloading table: inventory\n" \
                                f"Downloading week: {week}"
        df = td.dataFrameFromTabular(daxQ.inventory(startWeek,endWeek,week,dep))
        inventory_df = inventory_df.append(df)

    inventory_df.columns = col
    inventory_df.to_csv(path + "inventory.csv", index=False)

    ws.range("H15").value = datetime.datetime.now()
    ws.range("K15").value = "inventory"
    ws.range("B20").value = "Data has been saved"