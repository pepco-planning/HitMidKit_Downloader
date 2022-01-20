import functions as f
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