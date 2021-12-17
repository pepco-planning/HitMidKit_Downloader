import functions as f

def refreshBI():
    dep, startWeek, endWeek, path = f.getExcelData()

    f.runDaxQueries(startWeek,endWeek,dep,f.dax_query_list,path)
    #f.runDaxInventory(startWeek,endWeek,dep,path)


