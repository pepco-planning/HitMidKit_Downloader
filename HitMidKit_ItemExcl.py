############################################################################################
import HitMidKit_Functions as func
import HitMidKit_DAX as dax
import pandas as pd
import numpy as np

############################################################################################
def refreshItemExcl():

    # Opening and transforming data
    MODEL, MODEL_PATH = func.getPath()
    PARAMETERS  = func.getParameters(MODEL)

    df_db = pd.read_csv(MODEL_PATH + f"\Database_{PARAMETERS['Hierarchy']}.csv")
    df_ie = pd.read_excel(MODEL, sheet_name="ItemExclusion", usecols=['Option', 'Item Exclusion', 'Multi Overwirte'])
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
                'Markdown Value','MD_ratio','Multi','ItemExcl','NonPromo ROS', 'Promo ROS']

    cols_sku = ['SKU Sub Department','SKU Name','SKU Category','SKU Merch Group', 'SKU Merch Season Type',
                'SKU Merch To Season', 'SKU Merch Type', 'SKU PPL Initial Retail Price','Sales Margin','Option', 'VAT']

    df_final = pd.merge(df8[cols_inv],df_Sku_plu[cols_sku],on='Option', how='inner')
    df_final['ROS Margin Value'] = df_final['Final ROS_V'] / df_final['VAT'] * df_final['Sales Margin']

    # # Scoring A
    df_perf, ST_Tier_1, ST_Tier_2, MD_Tier1, MD_Tier2 = func.calcPerf(df_Perf,PARAMETERS['Hierarchy'])

    score_cols = ['Option','SKU Merch Type','ItemExcl','Sales Value','Final ROS_V', 'ROS Margin Value', 'Avg Sell-Through In Period',
                  'Markdown Value','SKU Sub Department', 'SKU Category']
    df_score1 = df_final[score_cols]

    if PARAMETERS['Hierarchy'][2] == '2': # division 2: Non Clothing
        df_score1 = pd.merge(df_score1, df_perf[['SKU Sub Department','Tier1','Tier2','Tier3']], on='SKU Sub Department', how='inner')
    else:
        df_score1 = pd.merge(df_score1, df_perf[['SKU Sub Department','SlsTier1','SlsTier2']], on='SKU Sub Department', how='inner')

    df_score1['MerchGroup'] = np.where(df_score1['SKU Merch Type']!='Y',"NonY","Y")
    df_score1['ItemCountDep'] = df_score1[df_score1.ItemExcl==0]['Option'].groupby(df_score1['SKU Sub Department']).transform('count')
    df_score1['ItemCountMer'] = df_score1[df_score1.ItemExcl==0]['Option'].groupby(df_score1['MerchGroup']).transform('count')

    # ## Scoring B
    ros_list = ['Final ROS_V','ROS Value','ROS Score']
    margin_list = ['ROS Margin Value','Margin Value','Margin Score']

    df_ros = func.calcScoringROS(ros_list[0],ros_list[1],ros_list[2],df_score1,PARAMETERS['Hierarchy'])
    df_margin = func.calcScoringROS(margin_list[0],margin_list[1],margin_list[2],df_score1,PARAMETERS['Hierarchy'])
    df_st = func.calcScoringSellThru(df_score1,ST_Tier_1,ST_Tier_2)
    df_md = func.calcScoringMD(df_score1,MD_Tier1,MD_Tier2)
    df_score2 = pd.merge(df_score1,df_ros,on='Option',how='left')
    df_score2 = pd.merge(df_score2,df_st,on='Option',how='left')
    df_score2 = pd.merge(df_score2,df_margin,on='Option',how='left')
    df_score2 = pd.merge(df_score2,df_md,on='Option',how='left')

    df_score2['MD_SLS'].fillna(0,inplace=True)
    df_score2['MD Score'] = np.where(df_score2['MD_SLS']==0,1,df_score2['MD Score'])

    if PARAMETERS['Hierarchy'][2] == '2':
        df_score2['TOTAL SCORE'] = df_score2['ROS Score'] + df_score2['Margin Score'] + np.where(df_score2['SKU Merch Type']=="Y",0,df_score2['ST Score'])

        df_score2['HIT / KIT / MID'] = np.where((df_score2['SKU Merch Type']=="Y")&(df_score2['TOTAL SCORE']>=2.5),"01_HIT",
                                           np.where(df_score2['TOTAL SCORE']>=3,"01_HIT",
                                                   np.where(df_score2['TOTAL SCORE']>=1.5,"02_MID","03_KIT")))
    else:
        df_score2['TOTAL SCORE'] = df_score2['ROS Score'] + df_score2['MD Score'] + df_score2['ST Score']

        df_score2['HIT / KIT / MID'] = np.where(df_score2['TOTAL SCORE']>2,"01_HIT",
                                                np.where(df_score2['TOTAL SCORE']>1,"02_MID","03_KIT"))
    if PARAMETERS['Hierarchy'][2] == '2':
        cols = ['Option','ROS Score','ST Score','Margin Score','TOTAL SCORE','HIT / KIT / MID']
    else:
        cols = ['Option','ROS Score','ST Score','MD Score','MD_SLS','TOTAL SCORE','HIT / KIT / MID']

    df_final = pd.merge(df_final, df_score2[cols], on='Option', how='inner')

    # Change the name in the code above in a free time
    df_final.rename(columns={'SKU Sub Department':'Sub Department','SKU Category':'Category','Final ROS_U':'ROS FINAL',
                             'Final ROS_V':'ROS Value FINAL','SKU PPL Initial Retail Price':'Initial Price',
                             'SKU Store Grade':'Grade','Sales Margin':'Sales Margin %','Avg Sell-Through In Period':'Sell-Through in Period',
                             'Sales Value in Period Promo':'Promo Sales Value in Period','ROS_ProdGrade':'RAW ROS (excludes MD)',
                             'Sales Units in Period Promo':'Promo Sales Units in Period', 'Margin Score':'ROS Margin Score',
                             'ROS_ProdGradeV':'RAW ROS Value (excludes MD)','Promo Adjusted ROS V':'Promo Adjusted ROS Value',
                             'Stock Units in Period':'Closing Stock Units','ST Score':'Sell-Through Score','MD_SLS':'MD % SLS',
                            'Markdown Value':'MD Retail'},
               inplace = True)
    hardline_cols = ['Sub Department','Category','Option','SKU Name','Multi','SKU Merch Group', 'SKU Merch Season Type', 'SKU Merch To Season',
                     'SKU Merch Type','Initial Price','Grade','Sales Margin %','Promo Start','Promo End','Week Start','Week End','Sales Value','Sales Units',
                     'Sales Value in Period','Sales Units in Period','% Total Sales Value','% Total Sales Units',
                     'Promo Sales Value in Period','Promo Sales Units in Period','RAW ROS (excludes MD)','NonPromo ROS',
                     'Promo ROS','Promo Adjusted ROS','ROS FINAL','RAW ROS Value (excludes MD)','Promo Adjusted ROS Value',
                     'ROS Value FINAL','ROS Score','Closing Stock Units','Sell-Through in Period','Sell-Through Score',
                     'ROS Margin Value','ROS Margin Score','TOTAL SCORE','HIT / KIT / MID','ItemExcl']

    clothing_cols = ['Sub Department','Category','Option','SKU Name','Multi','SKU Merch Group', 'SKU Merch Season Type', 'SKU Merch To Season',
                     'SKU Merch Type','Initial Price','Grade','Sales Margin %','Promo Start','Promo End','Week Start','Week End','Sales Value','Sales Units',
                     'Sales Value in Period','Sales Units in Period','% Total Sales Value','% Total Sales Units',
                     'Promo Sales Value in Period','Promo Sales Units in Period','RAW ROS (excludes MD)','NonPromo ROS',
                     'Promo ROS','Promo Adjusted ROS','ROS FINAL','RAW ROS Value (excludes MD)','Promo Adjusted ROS Value',
                     'ROS Value FINAL','ROS Score','Closing Stock Units','Sell-Through in Period','Sell-Through Score',
                     'MD Retail','MD % SLS','MD Score','TOTAL SCORE','HIT / KIT / MID','ItemExcl']
    if PARAMETERS['Hierarchy'][2] == '2':
        df_final = df_final[hardline_cols]
    else:
        df_final = df_final[clothing_cols]

    df_final.to_excel(MODEL_PATH + f"\Summary_{PARAMETERS['Hierarchy']}.xlsx",index=False)
    #df4.to_csv(MODEL_PATH + f"\Database_{PARAMETERS['Hierarchy']}.csv", index=False)
    df8.to_csv(MODEL_PATH + f"\Database_{PARAMETERS['Hierarchy']}.csv", index=False)

############################################################################################
refreshItemExcl()