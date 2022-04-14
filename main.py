import HitMidKit_Functions as f
import HitMidKit_Model as m
import xlwings as xw
import datetime

# import os
# from sys import path
# path.append(os.getcwd() + '\\dll')
# from pyadomd import Pyadomd

"""
function for "HitMidKit_model.xlsm"
"""
def refreshBI():
    """
    Function used via xlsm model. Once you click a button tn the model, the chosen files will be downloaded
    :return:
    """
    dep, startWeek, endWeek, path, fileName = f.getExcelData()      # get required parameters
    f.runDaxQueries(startWeek,endWeek,dep,f.dax_query_list,path)    # run DAX queries + download the chosen files and save them on PC
    f.getCalcParam(path, fileName)                                  # get parameters from "CalcPar" sheet and save it as .csv

def model1():
    PATH_DF4 = r"c:\Mariusz\MyProjects\HitMidKit_Downloader\input data\Analysis"

    wb = xw.Book.caller()
    ws = wb.sheets['notes']

    ws.range("C9").value = "Parameters"
    ws.range("C10").value = "Not Ready"
    ws.range("C11").value = datetime.datetime.now()

    df_final, HIERARCHY = m.modelPart1()
    df_final.to_excel(PATH_DF4 + f"\HMK_" + str(HIERARCHY) + "_0414.xlsx", index=False)
    #ws.range("C12").value = shape
    ws.range("D11").value = datetime.datetime.now()
    ws.range("C10").value = "Ready"