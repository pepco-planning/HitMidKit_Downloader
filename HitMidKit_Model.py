############################################################################################
# Import Files
import os
# from sys import path
# path.append(os.getcwd() + '\\dll')
# from pyadomd import Pyadomd

from HitMidKit_Functions import *
from daxQueries import *

def modelPart1():
    Parameters = getParameters()
    # Create Variables
    HIERARCHY = Parameters[Parameters['Variables Header'] == 'Hierarchy']['Chosen Variables'].item()
    DEPARTMENT = Parameters[Parameters['Variables Header'] == 'Departament']['Chosen Variables'].item()
    GRADING_TYPE = Parameters[Parameters['Variables Header'] == 'Grading Type']['Chosen Variables'].item()
    START_WEEK = int(Parameters[Parameters['Variables Header'] == 'Week From']['Chosen Variables'])
    END_WEEK = int(Parameters[Parameters['Variables Header'] == 'Week To']['Chosen Variables'])
    MIN_PAR = int(Parameters[Parameters['Variables Header'] == 'Minimum Sales Units']['Chosen Variables'])
    MIN_TOT_SLS = int(Parameters[Parameters['Variables Header'] == 'MinTotSls']['Chosen Variables'])
    MIN_STR_STK = int(Parameters[Parameters['Variables Header'] == 'MinStrStk']['Chosen Variables'])
    STR_COUNT = float(Parameters[Parameters['Variables Header'] == 'StrCount']['Chosen Variables'])
    WEEK_EXCL = float(Parameters[Parameters['Variables Header'] == 'WeekExcl']['Chosen Variables'])
    E_DURATION = int(Parameters[Parameters['Variables Header'] == 'E Merch Group Duration']['Chosen Variables'])
    W_DURATION = int(Parameters[Parameters['Variables Header'] == 'W Merch Group Duration']['Chosen Variables'])
    S_DURATION = int(Parameters[Parameters['Variables Header'] == 'S Merch Group Duration']['Chosen Variables'])

    PATH_INVENTORY = r"c:\Mariusz\MyProjects\HitMidKit_Downloader\input data\Analysis\Inventory"
    df_Inventory = pd.read_csv(PATH_INVENTORY + '\\' + f'{HIERARCHY}_HitMidKit.csv')

    ############################################################################################
    # variable for DAX
    df_Grading = dataFrameFromTabular(grading(START_WEEK,END_WEEK,DEPARTMENT))
    df_Markdown = dataFrameFromTabular(md(START_WEEK,END_WEEK,DEPARTMENT,MIN_PAR))
    df_Pcal = dataFrameFromTabular(pcal(START_WEEK,END_WEEK,DEPARTMENT))
    df_Perf = dataFrameFromTabular(perf_dep(START_WEEK,END_WEEK,DEPARTMENT))
    df_Plu_available = dataFrameFromTabular(plu_available(START_WEEK,END_WEEK,DEPARTMENT,MIN_PAR))
    df_Prh = dataFrameFromTabular(prh(START_WEEK,END_WEEK,DEPARTMENT))
    df_Prh_data = dataFrameFromTabular(prh_data(START_WEEK,END_WEEK,DEPARTMENT))
    df_Promo_reg = dataFrameFromTabular(promo_reg(START_WEEK,END_WEEK,DEPARTMENT,MIN_PAR))
    df_Promo_tv = dataFrameFromTabular(promo_tv(START_WEEK,END_WEEK,DEPARTMENT,MIN_PAR))
    df_Sku_plu = dataFrameFromTabular(sku_plu(START_WEEK,END_WEEK,DEPARTMENT,MIN_PAR))

    ############################################################################################
    # Data Transformation
    # Changing column names. Remove Characters such as '[', ']'
    table_list = [df_Grading,df_Markdown,df_Pcal,df_Perf,df_Plu_available,df_Prh,df_Prh_data,df_Promo_reg,df_Promo_tv,df_Sku_plu]
    for table in range(len(table_list)):
        cols = changeColumnName(table_list[table].columns)
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
    df_Promo_reg = createOption(df_Promo_reg)
    df_Sku_plu = createOption(df_Sku_plu)
    df_Plu_available = createOption(df_Plu_available)
    df_Inventory = createOption(df_Inventory)
    df_Promo_tv = createOption(df_Promo_tv)
    df_Markdown = createOption(df_Markdown)

    # Sometimes we have duplicates in SKU PLU (different initial price/margin). Keep the newest ones
    df_Sku_plu.drop(df_Sku_plu[df_Sku_plu.Option.duplicated(keep='first')].index, inplace = True)

    # Replace values to get just integer values: 0 if not [1,2,3]
    df_Sku_plu['SKU Store Grade'] = df_Sku_plu['SKU Store Grade'].apply(replace_grade)
    df_Sku_plu['SKU Store Grade'] = df_Sku_plu['SKU Store Grade'].astype(int)
    df_Sku_plu['SKU Store Grade'].unique()
    WeekMin = df_Pcal['Wk_Key'].min()
    WeekMax = df_Pcal['Wk_Key'].max()

    ############################################################################################
    # CALCULATIONS
    # Add weeks and gradings
    df_inv = addWeekKey(addCountry(df_Inventory, df_Grading, Parameters, GRADING_TYPE), df_Pcal)

    # Add "Plu Available" to the main table
    df_plu = df_Plu_available[['STR Company', 'Option', 'Available']].drop_duplicates()
    df_inv2 = pd.merge(df_inv, df_plu, on=['Option','STR Company'], how='left')
    df_inv2['Available'].fillna(0,inplace=True)
    df_inv2['Available'] = df_inv2['Available'].astype(int)

    # Add "Period Type" column (Promo, Markdown, Regular)
    df_inv3 = showPromo(df_inv2, df_Promo_reg, df_Sku_plu)

    # Calculate InStock and MinInStock *1
    df_inv3['InStock'] = np.where(((df_inv3['SalesU']>0)|(df_inv3['CSOHU']>MIN_STR_STK)),1,0)
    df_inv3['MinInStock'] = np.where(((df_inv3['SalesU']>0)|(df_inv3['CSOHU']>0)),1,0)

    # Fill Store Grade with zeros (NaN to 0) and set the data type into integer
    df_inv3['StoreGrade'].fillna(0,inplace=True)
    df_inv3['StoreGrade'] = df_inv3['StoreGrade'].astype(int)

    # Calculate amount of stores per grade for each PLU
    df_gs = GradeStores(df_Plu_available,df_Sku_plu,Parameters,df_Grading,GRADING_TYPE) .groupby('Option')['GradeStores'].sum().reset_index()

    # Add Grade Stores into the main table
    df4 = pd.merge(df_inv3,df_gs,on='Option',how='left')
    df4['MinStores'] = (df4['GradeStores'] * STR_COUNT).round(2) # not less than 60% of total stores
    df4['MinStoresX'] = (df4['GradeStores'] * WEEK_EXCL).round(2) # not less than 35% of total stores

    # Add 'SKU Store Grade' to the main table
    df_temp = df_Sku_plu[['Option','SKU Store Grade']]
    df4 = pd.merge(df4, df_temp, on='Option', how='left')
    del df_temp

    # Add Item Exclusion column (0 as default which means do not exclude the PLU)
    df4['ItemExcl'] = 0
    df4 = weeksCalc(df4, df_Sku_plu,WeekMin,WeekMax,E_DURATION,W_DURATION,S_DURATION)
    # df4.to_csv(PATH_DF4 + f"\df4_"+str(HIERARCHY)+"_0414.csv",index=False)

    # Aggregation - Option level
    df6 = InventoryAggregation(df4, df_Sku_plu, WeekMin, WeekMax,E_DURATION,W_DURATION,S_DURATION)

    avgROS, avgROS_Ratio = grade_multiEquivalentU(df4, df_Sku_plu, MIN_TOT_SLS)
    avgROSV, avgROS_RatioV = grade_multiEquivalentV(df4, df_Sku_plu, MIN_TOT_SLS)

    df_sellThru = avgSellThru(df4,df_Sku_plu,WeekMin,WeekMax,E_DURATION,W_DURATION,S_DURATION)
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
    df_effectiveWeeks = effectiveWeeks(df4)
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
                'Markdown Value','MD_ratio','Multi','ItemExcl','NonPromo ROS', 'Promo ROS'] #, 'ItemExcl'

    cols_sku = ['SKU Sub Department','SKU Category','SKU Merch Type','SKU Name',
                'SKU PPL Initial Retail Price','Sales Margin','Option', 'VAT']

    df_final = pd.merge(df8[cols_inv],df_Sku_plu[cols_sku],on='Option', how='inner')
    df_final['ROS Margin Value'] = df_final['Final ROS_V'] / df_final['VAT'] * df_final['Sales Margin']

    # # Scoring A
    df_perf, ST_Tier_1, ST_Tier_2, MD_Tier1, MD_Tier2 = calcPerf(df_Perf,HIERARCHY)

    score_cols = ['Option','SKU Merch Type','ItemExcl','Sales Value','Final ROS_V', 'ROS Margin Value', 'Avg Sell-Through In Period',
                  'Markdown Value','SKU Sub Department', 'SKU Category']
    df_score1 = df_final[score_cols]


    if HIERARCHY[2] == '2': # division 2: Non Clothing
        df_score1 = pd.merge(df_score1, df_perf[['SKU Sub Department','Tier1','Tier2','Tier3']], on='SKU Sub Department', how='inner')
    else:
        df_score1 = pd.merge(df_score1, df_perf[['SKU Sub Department','SlsTier1','SlsTier2']], on='SKU Sub Department', how='inner')

    df_score1['MerchGroup'] = np.where(df_score1['SKU Merch Type']!='Y',"NonY","Y")
    df_score1['ItemCountDep'] = df_score1[df_score1.ItemExcl==0]['Option'].groupby(df_score1['SKU Sub Department']).transform('count')
    df_score1['ItemCountMer'] = df_score1[df_score1.ItemExcl==0]['Option'].groupby(df_score1['MerchGroup']).transform('count')

    # ## Scoring B
    ros_list = ['Final ROS_V','ROS Value','ROS Score']
    margin_list = ['ROS Margin Value','Margin Value','Margin Score']

    df_ros = calcScoringROS(ros_list[0],ros_list[1],ros_list[2],df_score1,HIERARCHY)
    df_margin = calcScoringROS(margin_list[0],margin_list[1],margin_list[2],df_score1,HIERARCHY)
    df_st = calcScoringSellThru(df_score1,ST_Tier_1,ST_Tier_2)
    df_md = calcScoringMD(df_score1,MD_Tier1,MD_Tier2)
    df_score2 = pd.merge(df_score1,df_ros,on='Option',how='left')
    df_score2 = pd.merge(df_score2,df_st,on='Option',how='left')
    df_score2 = pd.merge(df_score2,df_margin,on='Option',how='left')
    df_score2 = pd.merge(df_score2,df_md,on='Option',how='left')

    df_score2['MD_SLS'].fillna(0,inplace=True)
    df_score2['MD Score'] = np.where(df_score2['MD_SLS']==0,1,df_score2['MD Score'])

    if HIERARCHY[2] == '2':
        df_score2['TOTAL SCORE'] = df_score2['ROS Score'] + df_score2['Margin Score'] + np.where(df_score2['SKU Merch Type']=="Y",0,df_score2['ST Score'])

        df_score2['HIT / KIT / MID'] = np.where((df_score2['SKU Merch Type']=="Y")&(df_score2['TOTAL SCORE']>=2.5),"01_HIT",
                                           np.where(df_score2['TOTAL SCORE']>=3,"01_HIT",
                                                   np.where(df_score2['TOTAL SCORE']>=1.5,"02_MID","03_KIT")))
    else:
        df_score2['TOTAL SCORE'] = df_score2['ROS Score'] + df_score2['MD Score'] + df_score2['ST Score']

        df_score2['HIT / KIT / MID'] = np.where(df_score2['TOTAL SCORE']>2,"01_HIT",
                                                np.where(df_score2['TOTAL SCORE']>1,"02_MID","03_KIT"))
    if HIERARCHY[2] == '2':
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
    hardline_cols = ['Sub Department','Category','Option','SKU Name','Multi','SKU Merch Type','Initial Price','Grade',
                         'Sales Margin %','Promo Start','Promo End','Week Start','Week End','Sales Value','Sales Units',
                         'Sales Value in Period','Sales Units in Period','% Total Sales Value','% Total Sales Units',
                         'Promo Sales Value in Period','Promo Sales Units in Period','RAW ROS (excludes MD)','NonPromo ROS',
                         'Promo ROS','Promo Adjusted ROS','ROS FINAL','RAW ROS Value (excludes MD)','Promo Adjusted ROS Value',
                         'ROS Value FINAL','ROS Score','ROS Margin Value','ROS Margin Score','Closing Stock Units',
                         'Sell-Through in Period','Sell-Through Score','TOTAL SCORE','HIT / KIT / MID','ItemExcl']
    clothing_cols = ['Sub Department','Category','Option','SKU Name','Multi','Initial Price','Grade',
                         'Sales Margin %','Promo Start','Promo End','Week Start','Week End','Sales Value','Sales Units',
                         'Sales Value in Period','Sales Units in Period','% Total Sales Value','% Total Sales Units',
                         'Promo Sales Value in Period','Promo Sales Units in Period','RAW ROS (excludes MD)','NonPromo ROS',
                         'Promo ROS','Promo Adjusted ROS','ROS FINAL','RAW ROS Value (excludes MD)','Promo Adjusted ROS Value',
                         'ROS Value FINAL','ROS Score','Closing Stock Units','Sell-Through in Period','Sell-Through Score',
                         'MD Retail','MD % SLS','MD Score','TOTAL SCORE','HIT / KIT / MID','ItemExcl']
    if HIERARCHY[2] == '2':
        df_final = df_final[hardline_cols]
    else:
        df_final = df_final[clothing_cols]

    return df_final, HIERARCHY
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

"""

"""
*1 
InStock [In-Stock Stores] potrzebny do policzenia week start, 
InStockX [In-Stock Stores in Period] do wykluczenia z wybranych tyg najgorszych sklepów

*2
Tu poprawić należy avgROS_RatioV dla units na avgROS_Ratio. W Excel Modelu jest błąd.
SKU PLU: ROS_AllStoreEquivalentu_PrAdj - MIN(Grades[RatioV]) zamień na MIN(Grades[Ratio])


"""