############################################################################################
import HitMidKit_Functions as func
import pandas as pd
import numpy as np

############################################################################################
def refreshKPI():

    # Opening and transforming data
    MODEL, MODEL_PATH = func.getPath()
    PARAMETERS  = func.getParameters(MODEL)

    df_ItemSummary = pd.read_excel(MODEL, sheet_name="ItemSummary")
    df_ItemSummary = df_ItemSummary.loc[:, ~df_ItemSummary.columns.str.contains('^Unnamed')]

    df_KpiOverwrite = pd.read_excel(MODEL, sheet_name="KPIsOverwrite", skiprows=1)
    df_KpiOverwrite = df_KpiOverwrite.loc[:, ~df_KpiOverwrite.columns.str.contains('^Unnamed')]

    a, b = func.refreshKPI(df_KpiOverwrite, int(PARAMETERS['Hierarchy'][2]))

    df = df_ItemSummary.copy()
    df['MerchGroup'] = np.where(df['SKU Merch Type']=='Y','Y','NonY')
    df = pd.merge(df,b,on='MerchGroup',how='left')
    df = pd.merge(df,a,on='Sub Department',how='left')

    # NonClothing:
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
        df['Sell-Through Score'] = np.where(df['Sell-Through in Period'] >= df['ST Tier 1'], 1,
                                            np.where(df['Sell-Through in Period'] >= df['ST Tier 2'], 0.5, 0))

        df['ROS Score'] = np.where(df['ROS Value FINAL'] >= df['ROS_V_Tier1'], 1.0,
                                   np.where(df['ROS Value FINAL'] >= df['ROS_V_Tier2'], 0.5, 0))

        df['MD Score'] = np.where(df['MD_SLS'] >= df['MD Tier 1'], 1,
                                  np.where(df['MD_SLS'] >= df['MD Tier 2'], 0.5, 0))

        df['TOTAL SCORE'] = df['ROS Score'] + df['MD Score'] + df['Sell-Through Score']

        df['HIT / KIT / MID'] = np.where(df['TOTAL SCORE'] > 2, "01_HIT",
                                         np.where(df['TOTAL SCORE'] > 1, "02_MID", "03_KIT"))

        df = df.iloc[:,:-7]

    df.to_excel(MODEL_PATH + f"\Summary_{PARAMETERS['Hierarchy']}.xlsx",index=False)

refreshKPI()