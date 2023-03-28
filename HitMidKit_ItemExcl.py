############################################################################################
import HitMidKit_Functions as func
import HitMidKit_DAX as dax
import pandas as pd
import numpy as np
import sys


def refreshItemExcl():

    # Opening and transforming data
    MODEL, MODEL_PATH = func.getPath()
    PARAMETERS = func.getParametersNew(MODEL_PATH + "\Parameters.csv")

    df_db = pd.read_csv(MODEL_PATH + f"\Database_{PARAMETERS['Hierarchy']}.csv")
    if df_db.shape[0] == 0:
        sys.exit()

    df_ie = pd.read_excel(MODEL, sheet_name="ItemExclusion", usecols=['Option', 'Item Exclusion', 'Multi Overwirte'])
    if df_ie.shape[0] == 0:
        sys.exit()

    df_Sku_plu = func.dataFrameFromTabular(dax.sku_plu(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament'],PARAMETERS['Minimum Sales Units']))
    df_Perf = func.dataFrameFromTabular(dax.perf_dep(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament']))

    table_list = [df_Perf,df_Sku_plu]
    for table in range(len(table_list)):
        cols = func.changeColumnName(table_list[table].columns)
        table_list[table].columns = cols

    df_Sku_plu = func.createOption(df_Sku_plu)
    df_Sku_plu['SKU Store Grade'] = df_Sku_plu['SKU Store Grade'].apply(func.replace_grade)
    df_Sku_plu['SKU Store Grade'] = df_Sku_plu['SKU Store Grade'].astype(int)
    df_Sku_plu['SKU Store Grade'].unique()

    # Overwrite 'Item Exclusion', 'Multi Overwirte'
    df8 = pd.merge(df_db, df_ie[['Option', 'Item Exclusion', 'Multi Overwirte']], on='Option', how='left')
    df8['ItemExcl'] = np.where(df8['Item Exclusion'] == 1, 1, df8['ItemExcl'])
    df8['Multi'] = np.where(df8['Multi Overwirte'] == 0, 0, df8['Multi'])
    df8.drop(columns={'Item Exclusion', 'Multi Overwirte'}, inplace=True)

    # Leave just required columns in a table
    cols_inv = ['Option', 'SKU Store Grade', 'Sales Units', 'Sales Value','Sales Units in Period',
                'Sales Units in Period Promo','Sales Units in Period Regular','Sales Value in Period',
                'Sales Value in Period Promo', 'Sales Value in Period Regular','Stock Units in Period',
                'GradeStores', 'Week Start', 'Week End', 'Promo Start', 'Promo End','% Total Sales Value',
                '% Total Sales Units', 'ROS_ProdGrade', 'ROS_ProdGradeV','Promo Adjusted ROS','Promo Adjusted ROS V',
                'Final ROS_U', 'Final ROS_V','Sell-Through', 'Sell-Through In Period', 'Avg Sell-Through In Period',
                'Multi','ItemExcl','NonPromo ROS', 'Promo ROS']

    cols_sku = ['SKU Class','SKU Name','SKU Subclass','SKU Merch Group', 'SKU Merch Season Type',
                'SKU Merch To Season', 'SKU Merch Type', 'SKU PPL Initial Retail Price','Sales Margin','Option', 'VAT']

    df_final = pd.merge(df8[cols_inv],df_Sku_plu[cols_sku],on='Option', how='inner')
    df_final['ROS Margin Value'] = df_final['Final ROS_V'] / df_final['VAT'] * df_final['Sales Margin']

    # # Scoring A
    df_perf, ST_Tier_1, ST_Tier_2 = func.calcPerf(df_Perf,PARAMETERS['Hierarchy'])
    score_cols = ['Option', 'SKU Merch Type', 'ItemExcl', 'Sales Value', 'Final ROS_V', 'ROS Margin Value',
                  'Avg Sell-Through In Period', 'SKU Class', 'SKU Subclass']
    df_score1 = df_final[score_cols]

    if PARAMETERS['Hierarchy'][2] == '2':  # division 2: Non Clothing
        df_score1 = pd.merge(df_score1, df_perf[['SKU Class', 'Tier1_GM', 'Tier2_GM', 'Tier3_GM']], on='SKU Class', how='inner')
    else:
        df_score1 = pd.merge(df_score1, df_perf[['SKU Class', 'Tier1_CL', 'Tier2_CL']], on='SKU Class', how='inner')

    df_score1['MerchGroup'] = np.where(df_score1['SKU Merch Type'] != 'Y', "NonY", "Y")
    df_score1['ItemCountDep'] = df_score1[df_score1.ItemExcl == 0]['Option'].groupby(df_score1['SKU Class']).transform('count')
    df_score1['ItemCountMer'] = df_score1[df_score1.ItemExcl == 0]['Option'].groupby(df_score1['MerchGroup']).transform('count')

    ## Scoring B
    ros_list = ['Final ROS_V', 'ROS Value', 'ROS Score']
    margin_list = ['ROS Margin Value', 'Margin Value', 'Margin Score']

    df_ros, df_ros_kpi = func.calcScoringROS(ros_list[0], ros_list[1], ros_list[2], df_score1, PARAMETERS['Hierarchy'])
    df_margin, df_margin_kpi = func.calcScoringROS(margin_list[0], margin_list[1], margin_list[2], df_score1, PARAMETERS['Hierarchy'])
    df_st, df_st_kpi = func.calcScoringSellThru(df_score1, ST_Tier_1, ST_Tier_2, PARAMETERS['Hierarchy'])

    # In Clothing Dep Margin has smaler weight 2 and 1 points instead of 4 and 2 (50% weight)
    if PARAMETERS['Hierarchy'][2] == '1':
        df_margin['Margin Score'] = df_margin['Margin Score'] * 0.5
    else:
        pass

    df_score2 = pd.merge(df_score1, df_ros, on='Option', how='left')
    df_score2 = pd.merge(df_score2, df_st, on='Option', how='left')
    df_score2 = pd.merge(df_score2, df_margin, on='Option', how='left')
    #########################################################################################################################################
    # KPI Overwrite
    df_margin_kpi.rename(columns={'ROS_V_Tier1': 'ROS_M_Tier1', 'ROS_V_Tier2': 'ROS_M_Tier2', 'ROS_V_Tier3': 'ROS_M_Tier3'},inplace=True)
    df_KPI_1 = pd.merge(df_ros_kpi, df_margin_kpi, on='SKU Class', how='inner')
    df_KPI_2 = df_st_kpi
    del df_ros_kpi, df_margin_kpi, df_st_kpi

    if PARAMETERS['Hierarchy'][2] == '2':
        df_score2['TOTAL SCORE'] = df_score2['ROS Score'] + df_score2['Margin Score'] + np.where(df_score2['SKU Merch Type'] == "Y", 0, df_score2['ST Score'])
        df_score2['HIT / KIT / MID'] = np.where((df_score2['SKU Merch Type'] == "Y") & (df_score2['TOTAL SCORE'] >= 2.5), "01_HIT",
                                                np.where(df_score2['TOTAL SCORE'] >= 3, "01_HIT",
                                                         np.where(df_score2['TOTAL SCORE'] >= 1.5, "02_MID", "03_KIT")))
        # df_final columns
        cols = ['Option', 'ROS Score', 'ST Score', 'Margin Score', 'TOTAL SCORE', 'HIT / KIT / MID']

        # KPI Overwrite: tu dodaj cały df_KPI_1 (ROS/Ros Margin) + ST
        df_KPI_1[['ROS_V_Tier1 MAN', 'ROS_V_Tier2 MAN', 'ROS_V_Tier3 MAN',
                  'ROS_M_Tier1 MAN', 'ROS_M_Tier2 MAN', 'ROS_M_Tier3 MAN']] = np.nan

        df_KPI_2 = df_KPI_2[['MerchGroup', 'ST Tier 1 Calc', 'ST Tier 2 Calc']]
        df_KPI_2.rename(columns={'ST Tier 1 Calc': 'ST Tier 1', 'ST Tier 2 Calc': 'ST Tier 2'}, inplace=True)
        df_KPI_2[['ST Tier 1 MAN', 'ST Tier 2 MAN']] = np.nan
    else:
        df_score2['TOTAL SCORE'] = df_score2['ROS Score'] + df_score2['Margin Score'] + df_score2['ST Score']
        df_score2['HIT / KIT / MID'] = np.where(df_score2['TOTAL SCORE'] >= 7, "01_HIT",
                                                np.where(df_score2['TOTAL SCORE'] >= 4, "02_MID", "03_KIT"))

        # df_final columns
        cols = ['Option', 'ROS Score', 'ST Score', 'Margin Score', 'TOTAL SCORE', 'HIT / KIT / MID']

        # KPI Overwrite: tu dodaj cały df_KPI_1 (ROS/Ros Margin) + ST
        df_KPI_1[['ROS_V_Tier1 MAN', 'ROS_V_Tier2 MAN',
                  'ROS_M_Tier1 MAN', 'ROS_M_Tier2 MAN']] = np.nan

        df_KPI_2 = df_KPI_2[['MerchGroup', 'ST Tier 1 Calc', 'ST Tier 2 Calc']]
        df_KPI_2.rename(columns={'ST Tier 1 Calc': 'ST Tier 1', 'ST Tier 2 Calc': 'ST Tier 2'}, inplace=True)
        df_KPI_2[['ST Tier 1 MAN', 'ST Tier 2 MAN']] = np.nan

    df_final = pd.merge(df_final, df_score2[cols], on='Option', how='inner')
    #########################################################################################################################################
    # Change the name in the code above in a free time
    final_cols = ['Class', 'Subclass', 'Option', 'SKU Name', 'Multi', 'SKU Merch Group', 'SKU Merch Season Type',
                  'SKU Merch To Season','SKU Merch Type', 'Initial Price', 'Grade', 'Sales Margin %', 'Promo Start', 'Promo End',
                  'Week Start', 'Week End', 'Sales Value', 'Sales Units','Sales Value in Period', 'Sales Units in Period',
                  '% Total Sales Value', '% Total Sales Units','Promo Sales Value in Period', 'Promo Sales Units in Period',
                  'RAW ROS', 'NonPromo ROS','Promo ROS', 'Promo Adjusted ROS', 'ROS FINAL', 'RAW ROS Value', 'Promo Adjusted ROS Value',
                  'ROS Value FINAL', 'ROS Score', 'Closing Stock Units', 'Sell-Through in Period', 'Sell-Through Score',
                  'ROS Margin Value', 'ROS Margin Score', 'TOTAL SCORE', 'HIT / KIT / MID', 'ItemExcl']

    df_final.rename(columns={'SKU Class': 'Class', 'SKU Subclass': 'Subclass', 'Final ROS_U': 'ROS FINAL',
                             'Final ROS_V': 'ROS Value FINAL', 'SKU PPL Initial Retail Price': 'Initial Price',
                             'SKU Store Grade': 'Grade', 'Sales Margin': 'Sales Margin %',
                             'Avg Sell-Through In Period': 'Sell-Through in Period',
                             'Sales Value in Period Promo': 'Promo Sales Value in Period', 'ROS_ProdGrade': 'RAW ROS',
                             'Sales Units in Period Promo': 'Promo Sales Units in Period',
                             'Margin Score': 'ROS Margin Score',
                             'ROS_ProdGradeV': 'RAW ROS Value', 'Promo Adjusted ROS V': 'Promo Adjusted ROS Value',
                             'Stock Units in Period': 'Closing Stock Units', 'ST Score': 'Sell-Through Score'},
                    inplace=True)

    df_final = df_final[final_cols]

    df_final.to_excel(MODEL_PATH + f"\Summary_{PARAMETERS['Hierarchy']}.xlsx",index=False)
    df_KPI_1.to_csv(MODEL_PATH + f"\KPI_{PARAMETERS['Hierarchy']}.csv", index=False)
    df_KPI_2.to_csv(MODEL_PATH + f"\KPI_{PARAMETERS['Hierarchy']}.csv", mode='a', index=False)

############################################################################################
refreshItemExcl()