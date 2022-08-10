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

    MODEL, MODEL_PATH = f.getPath()
    PARAMETERS = f.getParametersNew(MODEL_PATH + "\Parameters.csv")
    #dep, startWeek, endWeek, minPar, path, hierIdx, reportType = f.loadParameters()
    f.loadInventory(PARAMETERS['Week From'],
                    PARAMETERS['Week To'],
                    PARAMETERS['Minimum Sales Units'],
                    PARAMETERS['Departament'],
                    PARAMETERS['Path Inventory'],
                    PARAMETERS['Hierarchy'])


downloadData()
