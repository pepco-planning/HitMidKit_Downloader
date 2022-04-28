############################################################################################
import HitMidKit_Functions as func
import HitMidKit_DAX as dax
import pandas as pd
import numpy as np

############################################################################################
def calculateHitMidKit():
    STATUS_TOTAL_NO = 9 # this value need tp be fix
    MODEL, MODEL_PATH = func.getPath()

    PARAMETERS  = func.getParameters(MODEL)

    func.showStatusModel(MODEL_PATH, 1, PARAMETERS['Departament'], PARAMETERS['Hierarchy'], 'Downloading Inventory',1,STATUS_TOTAL_NO)
    df_Inventory = pd.read_csv(PARAMETERS['Path Inventory'] + "\\" + f"{PARAMETERS['Hierarchy']}_HitMidKit.csv")

    ############################################################################################
    # variable for DAX
    func.showStatusModel(MODEL_PATH, 1, PARAMETERS['Departament'], PARAMETERS['Hierarchy'], 'Downloading Reports',2,STATUS_TOTAL_NO)
    df_Grading = func.dataFrameFromTabular(dax.grading(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament']))
    df_Markdown = func.dataFrameFromTabular(dax.md(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament'],PARAMETERS['Minimum Sales Units']))
    df_Pcal = func.dataFrameFromTabular(dax.pcal(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament']))
    df_Perf = func.dataFrameFromTabular(dax.perf_dep(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament']))
    df_Plu_available = func.dataFrameFromTabular(dax.plu_available(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament'],PARAMETERS['Minimum Sales Units']))
    df_Prh = func.dataFrameFromTabular(dax.prh(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament']))
    df_Prh_data = func.dataFrameFromTabular(dax.prh_data(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament']))
    df_Promo_reg = func.dataFrameFromTabular(dax.promo_reg(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament'],PARAMETERS['Minimum Sales Units']))
    df_Promo_tv = func.dataFrameFromTabular(dax.promo_tv(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament'],PARAMETERS['Minimum Sales Units']))
    df_Sku_plu = func.dataFrameFromTabular(dax.sku_plu(PARAMETERS['Week From'],PARAMETERS['Week To'],PARAMETERS['Departament'],PARAMETERS['Minimum Sales Units']))

    ############################################################################################
    # Data Transformation
    # Changing column names. Remove Characters such as '[', ']'
    func.showStatusModel(MODEL_PATH, 1, PARAMETERS['Departament'], PARAMETERS['Hierarchy'],
                         'Data Transformation', 3, STATUS_TOTAL_NO)
    table_list = [df_Grading,df_Markdown,df_Pcal,df_Perf,df_Plu_available,df_Prh,df_Prh_data,df_Promo_reg,df_Promo_tv,df_Sku_plu]
    for table in range(len(table_list)):
        cols = func.changeColumnName(table_list[table].columns)
        table_list[table].columns = cols

    ############################################################################################
    # Adjust the Plu Available table
    df_Plu_available = pd.melt(df_Plu_available, id_vars=['SKU PLU', 'SKU Colour'],
                            var_name='xtemp', value_name='Available')
    df_Plu_available['Available'] = np.where(df_Plu_available['Available']=='No',0,1)
    df_Plu_available['STR Company'] = df_Plu_available['xtemp'].str[4:7]
    df_Plu_available.drop(columns=(['xtemp']),axis=1,inplace=True)
    df_Inventory['Multi'] = np.where((df_Inventory['SKU Colour'].str[:5]=='Multi'),1,0)
    df_Inventory['STR Number'] = df_Inventory['STR Number'].astype(int)
    df_Inventory['Pl Year'] = df_Inventory['Pl Year'].astype(int)
    df_Pcal.rename(columns={'PCAL_WEEK_KEY': 'Wk_Key'}, inplace=True)
    df_Promo_reg.rename(columns={'PCAL_WEEK_KEY': 'Wk_Key'}, inplace=True)

    # Create 'Option' column
    df_Promo_reg = func.createOption(df_Promo_reg)
    df_Sku_plu = func.createOption(df_Sku_plu)
    df_Plu_available = func.createOption(df_Plu_available)
    df_Inventory = func.createOption(df_Inventory)
    df_Promo_tv = func.createOption(df_Promo_tv)
    df_Markdown = func.createOption(df_Markdown)

    # Sometimes we have duplicates in SKU PLU (different initial price/margin). Keep the newest ones
    df_Sku_plu.drop(df_Sku_plu[df_Sku_plu.Option.duplicated(keep='first')].index, inplace = True)

    # Replace values to get just integer values: 0 if not [1,2,3]
    df_Sku_plu['SKU Store Grade'] = df_Sku_plu['SKU Store Grade'].apply(func.replace_grade)
    df_Sku_plu['SKU Store Grade'] = df_Sku_plu['SKU Store Grade'].astype(int)
    df_Sku_plu['SKU Store Grade'].unique()
    WeekMin = df_Pcal['Wk_Key'].min()
    WeekMax = df_Pcal['Wk_Key'].max()

    ############################################################################################
    # CALCULATIONS
    # Add weeks and gradings
    df_inv = func.addWeekKey(func.addCountry(df_Inventory, df_Grading, float(PARAMETERS['Grade2']), float(PARAMETERS['Grade3']),PARAMETERS['Grading Type']), df_Pcal)

    # Add "Plu Available" to the main table
    df_plu = df_Plu_available[['STR Company', 'Option', 'Available']].drop_duplicates()
    df_inv2 = pd.merge(df_inv, df_plu, on=['Option','STR Company'], how='left')
    df_inv2['Available'].fillna(0,inplace=True)
    df_inv2['Available'] = df_inv2['Available'].astype(int)

    # Add "Period Type" column (Promo, Markdown, Regular)
    df_inv3 = func.showPromo(df_inv2, df_Promo_reg, df_Sku_plu)

    # Calculate InStock and MinInStock *1
    df_inv3['InStock'] = np.where(((df_inv3['SalesU']>0)|(df_inv3['CSOHU']>int(PARAMETERS['MinStrStk']))),1,0)
    df_inv3['MinInStock'] = np.where(((df_inv3['SalesU']>0)|(df_inv3['CSOHU']>0)),1,0)

    # Fill Store Grade with zeros (NaN to 0) and set the data type into integer
    df_inv3['StoreGrade'].fillna(0,inplace=True)
    df_inv3['StoreGrade'] = df_inv3['StoreGrade'].astype(int)

    # Calculate amount of stores per grade for each PLU
    df_gs = func.GradeStores(df_Plu_available,df_Sku_plu,float(PARAMETERS['Grade2']), float(PARAMETERS['Grade3']),
                             df_Grading,PARAMETERS['Grading Type']) .groupby('Option')['GradeStores'].sum().reset_index()

    # Add Grade Stores into the main table
    df4 = pd.merge(df_inv3,df_gs,on='Option',how='left')
    df4['MinStores'] = (df4['GradeStores'] * float(PARAMETERS['StrCount'])).round(2) # not less than 60% of total stores
    df4['MinStoresX'] = (df4['GradeStores'] * float(PARAMETERS['WeekExcl'])).round(2) # not less than 35% of total stores

    # Add 'SKU Store Grade' to the main table
    df_temp = df_Sku_plu[['Option','SKU Store Grade']]
    df4 = pd.merge(df4, df_temp, on='Option', how='left')
    del df_temp

    # Add Item Exclusion column (0 as default which means do not exclude the PLU)
    df4['ItemExcl'] = 0
    df4 = func.weeksCalc(df4, df_Sku_plu,WeekMin,WeekMax,
                         int(PARAMETERS['E Merch Group Duration']),int(PARAMETERS['W Merch Group Duration']),int(PARAMETERS['S Merch Group Duration']))
    # df4.to_csv(PATH_DF4 + f"\df4_"+str(HIERARCHY)+"_0414.csv",index=False)

########################
    """SECOND PART OF THE MODEL - thing how to save the first part to avoid running it every time..."""
    # Aggregation - Option level
    func.showStatusModel(MODEL_PATH, 1, PARAMETERS['Departament'], PARAMETERS['Hierarchy'],
                         'Data Aggregation', 4, STATUS_TOTAL_NO)
    df6 = func.InventoryAggregation(df4, df_Sku_plu, WeekMin, WeekMax,
                                    int(PARAMETERS['E Merch Group Duration']),int(PARAMETERS['W Merch Group Duration']),int(PARAMETERS['S Merch Group Duration']))

    func.showStatusModel(MODEL_PATH, 1, PARAMETERS['Departament'], PARAMETERS['Hierarchy'],
                         'ROS Calculations', 5, STATUS_TOTAL_NO)
    avgROS, avgROS_Ratio = func.grade_multiEquivalentU(df4, df_Sku_plu, int(PARAMETERS['MinTotSls']))
    avgROSV, avgROS_RatioV = func.grade_multiEquivalentV(df4, df_Sku_plu, int(PARAMETERS['MinTotSls']))

    df_sellThru = func.avgSellThru(df4,df_Sku_plu,WeekMin,WeekMax,
                                   int(PARAMETERS['E Merch Group Duration']),int(PARAMETERS['W Merch Group Duration']),int(PARAMETERS['S Merch Group Duration']))
    df6 = pd.merge(df6, df_sellThru, on=['Option','SKU Store Grade'], how='left')

    # Finalization ROS *2
    df6['Final ROS_U'] = np.where(df6['SKU Store Grade']==2,
                                   df6['Promo Adjusted ROS'] / avgROS_RatioV['2'][0],
                                   np.where(df6['SKU Store Grade']==3,
                                           df6['Promo Adjusted ROS'] / avgROS_RatioV['3'][0],
                                           df6['Promo Adjusted ROS']))
    df6['Final ROS_V'] = np.where(df6['SKU Store Grade']==2,
                                   df6['Promo Adjusted ROS V'] / avgROS_RatioV['2'][0],
                                   np.where(df6['SKU Store Grade']==3,
                                           df6['Promo Adjusted ROS V'] / avgROS_RatioV['3'][0],
                                           df6['Promo Adjusted ROS V']))

    # Excluding Promo TV
    df_Promo_tv['Promo_tv'] = 1
    df_Promo_tv = df_Promo_tv[['Option','Promo_tv']].drop_duplicates()
    df7 = pd.merge(df6, df_Promo_tv, on='Option', how='left')
    df7['Promo_tv'].fillna(0,inplace=True)
    df7['Promo_tv'] = df7['Promo_tv'].astype(int)

    # Sell-Through
    df7['Sell-Through In Period'] = df7['Sales Units in Period'] / (df7['Sales Units in Period'] + df7['Stock Units in Period'])
    df7['Sell-Through'] = df7['Sales Units'] / (df7['Sales Units'] + df7['Stock Units in Period'])

    # Markdowns
    df_Markdown['Markdown Value'] = -df_Markdown['Markdown Retail Report']
    df8 = pd.merge(df7, df_Markdown[['Option','Markdown Value']], on='Option', how='left')
    df8['Markdown Value'].fillna(0,inplace=True)
    df8['Markdown Value'] = df8['Markdown Value'].astype(int)
    df8['MD_ratio'] = np.where(df8['Markdown Value']>0,
                             np.where(df8['Sales Value']>0,
                                     df8['Markdown Value'] / df8['Sales Value'],
                                     1),0)

    # Item Exclusion
    df_effectiveWeeks = func.effectiveWeeks(df4)
    df8 = pd.merge(df8,df_effectiveWeeks,on='Option',how='left')
    df8['Effective Weeks'].fillna(0,inplace=True)
    df8['Effective Weeks'] = df8['Effective Weeks'].astype(int)
    df8['Weeks In Period'] = (df8['Week End'] - df8['Week Start']) + 1
    df8['ItemExcl'] = np.where(((df8['Week Start']==0)&(df8['Week End']==0)),2,
                               np.where(df8['Promo_tv']==1,3,
                                        np.where(df8['Effective Weeks']<=(0.4*df8['Weeks In Period']),4,0)))

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
    func.showStatusModel(MODEL_PATH, 1, PARAMETERS['Departament'], PARAMETERS['Hierarchy'],
                         'Scoring Calculations', 6, STATUS_TOTAL_NO)
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
    func.showStatusModel(MODEL_PATH, 1, PARAMETERS['Departament'], PARAMETERS['Hierarchy'],
                         'Finalizing Data', 7, STATUS_TOTAL_NO)
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

    func.showStatusModel(MODEL_PATH, 1, PARAMETERS['Departament'], PARAMETERS['Hierarchy'],
                         'Saving Data', 8, STATUS_TOTAL_NO)

    df_final.to_excel(MODEL_PATH + f"\Summary_{PARAMETERS['Hierarchy']}.xlsx",index=False)
    #df4.to_csv(MODEL_PATH + f"\Database_{PARAMETERS['Hierarchy']}.csv", index=False)
    df8.to_csv(MODEL_PATH + f"\Database_{PARAMETERS['Hierarchy']}.csv", index=False)
    #df_Sku_plu.to_csv(MODEL_PATH + f"\Sku_plu_{PARAMETERS['Hierarchy']}.csv", index=False)
    func.showStatusModel(MODEL_PATH, 0, dep=None, hier=None, status=None,value=STATUS_TOTAL_NO,value_total=STATUS_TOTAL_NO)

############################################################################################
calculateHitMidKit()

############################################################################################
# QUESTIONS
"""

**Questions**
Should the stores with no grades be taken into consideration?
- Should we fill NaN values in StoreGrade column?
- How to treat stores with Grade=0. Should we calculate for the only "Total Sales", "InStock/MinInStock"? <p>Currently, InStock does not calculate for Grade = 0 but the MinInStock does. Is it proper approach?

df_final[(df_final['Sub Department']=='e Baskets and Boxes')&(df_final.ItemExcl==0)]['Option'].count()
df_final.to_excel(f"Analysis/HMK_"+str(chosenHierarchy)+"_0413.xlsx",index=False)
___
**Question**

Do we need units in promo_reg file? Why the values there are not equal with Inventory (Likely answer: "promo" shows sales only for promo products while in "Inventory" table just a part of the sales shows promo because the first day of promotion starts on Thursday)?
**Final Questions**</p>
- <p>W Data Model mamy błędne SKU Merch Group (note2/Hit Mid Kit_Analysis.xlsx).</p>
- <p>Po co pokazywać wartości w ItemSummary dla produków takich jak "276636 Multicolor 5", czyli takich które nie mają określonego weekfrom/end i mają puste wiersze</p>
- <p>Czy trzeba brać pod uwagę sklepy, których nie ma w pliku Grading.csv (note4/Hit Mid Kit_Analysis.xlsx)?</p>
- <p>Po co brać pod uwagę PLU bez sprzedaży (brak w Inventory?)</p>
    ['326220 Multicolor 5',
    '327847 Multicolor 5',
    '325537 Multicolor 5',
    '331142 Multicolor 5',
    '314443 Multicolor 5',
    '325966 Multicolor 5',
    '331658 Multicolor 5',
    '327329 Multicolor 5',
    '312499 Royal Blue',
    '312500 Black',
    '312502 Black',
    '327085 Multicolor 5',
    '280429 Multicolor 5']
- Sell Through if regular to 0
- ItemExclusion: dodaj tę funkcję do df4 bo ratio/ros/promofactor liczą się dla ItemExcl=0
- Jest opcja zmiany rosa (w data model "x") - dodaj to
---
# Next Steps
- <p>df6 = ItemExclusion(df6, Sku_plu)<p>There will be two buttons in the model (calculate/refresh parameters). If we click refresh then "ItemExclusion" table will be overwritten and the scoring will be re-calculated. Note that the first runn need to be done before we will be able to refresh the calculations. Suggestions: calculate the tables before share the files with planners</p>
- <p>In "aggregateInv" I use left join because if Merch Type = Y then no matter if InStock >= MinStores. Double check it</p>
- <p>Item Exclusion</p>
    Wrzuć tabelę do excela a planiści manualnie mogą zmieniać 3 kolumny po czym przeliczą się wartości ponownie.
    IF(LOOKUPVALUE(ItemExcl[Multi],ItemExcl[Option],'SKU PLU'[Option])="X",
    IF(LEFT('SKU PLU'[SKU Colour],5)="Multi",1,0),
    0)
- Time Lap niech zależy oid ilości linii w Inventory per ratio dla każdego z zapytanie: saveStatus
"""

"""
*1 
InStock [In-Stock Stores] potrzebny do policzenia week start, 
InStockX [In-Stock Stores in Period] do wykluczenia z wybranych tyg najgorszych sklepów

*2
Tu poprawić należy avgROS_RatioV dla units na avgROS_Ratio. W Excel Modelu jest błąd.
SKU PLU: ROS_AllStoreEquivalentu_PrAdj - MIN(Grades[RatioV]) zamień na MIN(Grades[Ratio])


"""