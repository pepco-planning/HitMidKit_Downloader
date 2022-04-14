import HitMidKit_Functions as f

def downloadData():
    """
    Short description:
        .exe file loading data in the background

    loadParameters:
        - open excel file with parameters and get the required data. Close the file

    loadInventory:
        - run the query and save the outputs on the chosen path
        - create .txt file informing is the data ready or not

    :return:
    """

    dep, startWeek, endWeek, minPar, path, hierIdx, reportType = f.loadParameters()

    if reportType == "Only Inventory":
        f.loadInventory(startWeek,endWeek,minPar,dep,path,hierIdx)
    elif reportType == "All Reports (excl. Inventory)":
        f.runDaxQueriesExe(startWeek,endWeek,dep,f.dax_query_list,path,hierIdx)
    else:
        f.runDaxQueriesExe(startWeek,endWeek,dep,f.dax_query_list,path,hierIdx)
        f.loadInventory(startWeek,endWeek,minPar,dep,path,hierIdx)

downloadData()
