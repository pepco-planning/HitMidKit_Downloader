############################################################################################
import HitMidKit_Functions as func
import pandas as pd
import numpy as np
import sys

############################################################################################
def refreshKPI():
    # Variables
    debug = 0 # 1: Yes, 0: No (required for Jupyter Notebook analysis)
    MODEL, MODEL_PATH = func.getPath()
    PARAMETERS = func.getParametersNew(MODEL_PATH + "\Parameters.csv")
    if debug == 1:
        MODEL = 'c:\Mariusz\MyProjects\HitMidKit_Downloader\workbook\Model\JupyterModel\Test_202303\excel\HitMidKit_2023_Q2_AW22-ClothingTest.xlsm'
    else:
        pass
    ############################################################################################
    # Reading Actual KPIs
    df_ItemSummary = pd.read_excel(MODEL, sheet_name="ItemSummary")
    df_ItemSummary = df_ItemSummary.loc[:, ~df_ItemSummary.columns.str.contains('^Unnamed')]
    if df_ItemSummary.shape[0] == 0:
        sys.exit()

    df_KpiOverwrite = pd.read_excel(MODEL, sheet_name="KPIsOverwrite", skiprows=1)  # , header = None
    if df_KpiOverwrite.shape[0] == 0:
        sys.exit()
    ############################################################################################
    # Refreshing KPIs
    a, b = func.refreshKPI(df_KpiOverwrite, int(PARAMETERS['Hierarchy'][2]))

    df = df_ItemSummary.copy()
    df['MerchGroup'] = np.where(df['SKU Merch Type'] == 'Y', 'Y', 'NonY')
    df = pd.merge(df, b, on='MerchGroup', how='left')
    df = pd.merge(df, a, on='Class', how='left')
    ############################################################################################
    # Re-Calculations Final Scores
    if PARAMETERS['Hierarchy'][2] == '2':
        df['Sell-Through Score'] = np.where(df['Sell-Through in Period'] >= df['ST Tier 1'], 1,
                                            np.where(df['Sell-Through in Period'] >= df['ST Tier 2'], 0.5, 0))

        df['ROS Score'] = np.where(df['ROS Value FINAL'] >= df['ROS_V_Tier1'], 1.5,
                                   np.where(df['ROS Value FINAL'] >= df['ROS_V_Tier2'], 1.0,
                                            np.where(df['ROS Value FINAL'] >= df['ROS_V_Tier3'], 0.5, 0)))

        df['ROS Margin Score'] = np.where(df['ROS Margin Value'] >= df['ROS_M_Tier1'], 1.5,
                                          np.where(df['ROS Margin Value'] >= df['ROS_M_Tier2'], 1.0,
                                                   np.where(df['ROS Margin Value'] >= df['ROS_M_Tier3'], 0.5, 0)))

        df['TOTAL SCORE'] = df['ROS Score'] + df['ROS Margin Score'] + np.where(df['SKU Merch Type'] == "Y", 0,
                                                                                df['Sell-Through Score'])
        df['HIT / KIT / MID'] = np.where((df['SKU Merch Type'] == "Y") & (df['TOTAL SCORE'] >= 2.5), "01_HIT",
                                         np.where(df['TOTAL SCORE'] >= 3, "01_HIT",
                                                  np.where(df['TOTAL SCORE'] >= 1.5, "02_MID", "03_KIT")))
        df = df.iloc[:, :-9]
    else:
        df['Sell-Through Score'] = np.where(df['Sell-Through in Period'] >= df['ST Tier 1'], 4,
                                            np.where(df['Sell-Through in Period'] >= df['ST Tier 2'], 2, 0))

        df['ROS Score'] = np.where(df['ROS Value FINAL'] >= df['ROS_V_Tier1'], 4,
                                   np.where(df['ROS Value FINAL'] >= df['ROS_V_Tier2'], 2, 0))

        df['ROS Margin Score'] = np.where(df['ROS Margin Value'] >= df['ROS_M_Tier1'], 2,
                                          np.where(df['ROS Margin Value'] >= df['ROS_M_Tier2'], 1, 0))

        df['TOTAL SCORE'] = df['ROS Score'] + df['ROS Margin Score'] + df['Sell-Through Score']
        df['HIT / KIT / MID'] = np.where(df['TOTAL SCORE'] >= 7, "01_HIT",
                                         np.where(df['TOTAL SCORE'] >= 4, "02_MID", "03_KIT"))

        df = df.iloc[:, :-7]
    ############################################################################################
    # Saving
    df.to_excel(MODEL_PATH + f"\Summary_{PARAMETERS['Hierarchy']}.xlsx", index=False)

refreshKPI()