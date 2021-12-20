import functions as f

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

    dep, startWeek, endWeek, minPar, path, hierIdx = f.loadParameters()
    f.loadInventory(startWeek,endWeek,minPar,dep,path,hierIdx)

downloadData()
