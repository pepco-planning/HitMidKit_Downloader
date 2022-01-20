import xlwings as xw
import pandas as pd
import numpy as np

# Below tables show how the DataModel in PowerPivot looks like
def ExclusionType():
    # no additional cols
    pass

def calcGrading(df, ans, g2, g3):
    """
    Columns:
        The columns from the file: 'Grading.csv' +
        RankX:      Sales ranking based on 'DepSalesV' sorted DESC (1,2,3...)
        Cumm%:      IF([GradeType] = "Count",Grading[RankX]/[StoreCount],
                    SUMX(FILTER(Grading,Grading[RankX]<=EARLIER(Grading[RankX])),[DepSalesV])/[SalesTotal])
        Grade:      MAXX(FILTER(Grades,Grades[Cut-Off (Count / Value)]>=EARLIER([Cumm%])),Grades[Grade])

    :return:
    """
    column_names = ['STR Number', 'STR Company', 'DepSalesV']
    df.columns = column_names
    df = df.sort_values('DepSalesV', ascending=False)
    df['RankX'] = np.arange(df.shape[0]) + 1

    storeCount = len(df['STR Number'].unique())
    df['Cumm%'] = np.where(ans == "Count", df['RankX'] / storeCount,
                           df['DepSalesV'].cumsum() / df['DepSalesV'].sum())

    df['Grade'] = np.where(df['Cumm%'] < g3, 3, np.where(df['Cumm%'] < g2, 2, 1))
    return df

def Grades(path):
    """
    Grades from 'CalcPar' sheet (1,2,3)

    Columns:
        Cut-Off:            'CalcPar' sheet values
        StoreGradeCount:    CALCULATE(COUNT(Grading[Store STR]),
                            Grading[Grade]>=EARLIER(Grades[Grade]),ALL(Grades))
        AvgRosGrade:        AVERAGEX(
                            ADDCOLUMNS(
                                FILTER(
                                    ADDCOLUMNS(
                                        FILTER('SKU PLU',AND('SKU PLU'[ItemExcl]=0,'SKU PLU'[Grade]="1")),
                                        "SalesTemp",
                                        CALCULATE([Sales Units],ALL(Grades))),
                                    [SalesTemp]>=[MinTotSls]),
                                "ROS",
                                [ROS_Cumm]),
                            [ROS])
        Ratio:              Grades[AvgRosGrade]/CALCULATE(MIN([AvgRosGrade]),ALL(Grades),Grades[Grade]=1)
        AvgRosGradeV:       AVERAGEX(
                            ADDCOLUMNS(
                                FILTER(
                                    ADDCOLUMNS(
                                        FILTER('SKU PLU',AND('SKU PLU'[ItemExcl]=0,'SKU PLU'[Grade]="1")),
                                        "SalesTemp",
                                        CALCULATE([Sales Units],ALL(Grades))),
                                    [SalesTemp]>=[MinTotSls]),
                                "ROS",
                                [ROS_CummV_PrAdj]),
                            [ROS])
        RatioV:             Grades[AvgRosGradeV]/CALCULATE(MIN([AvgRosGradeV]),ALL(Grades),Grades[Grade]=1)
    :return:
    """
    fileName = path + 'parameters.csv'
    df = pd.read_csv(fileName)

    pass

def Other():
    """
    Contains the first table from the excel - 'CalPar' sheet (8 rows):

    :return:
    """

def PCAL():
    """
    Pcal.csv columns plus:
        idx:    COUNTX(FILTER(PCAL,PCAL[Wk_Key]<=EARLIER(PCAL[Wk_Key])),PCAL[Wk_Key])
    :return:
    """

def PRH():
    # no additional cols
    pass

def SKU_PLU():
    """
    Columns from Sku_plu.csv file plus:
        WeekFrom:           if(
                            IF(
                                AND([SKU Merch Group]="Y",[SKU Merch Season]="A"),
                                [WeekMin],
                                [WeekStart])>[WeekMax],0,IF(
                                AND([SKU Merch Group]="Y",[SKU Merch Season]="A"),
                                [WeekMin],
                                [WeekStart]))
        WeekTo:             var WKFROM = LOOKUPVALUE(PCAL[idx],PCAL[Wk_Key],'SKU PLU'[WeekFrom])
                            var WKMAX = MAX(PCAL[idx])
                            return
                            IF([WeekFrom] = 0 , 0 ,LOOKUPVALUE(
                            PCAL[Wk_Key],PCAL[idx],
                            MIN(SWITCH([SKU Merch Group],
                            "Y",WKMAX,
                            "E",WKFROM+[E_Dur]-1,
                            "W",WKFROM+[W_Dur]-1,
                            "S",WKFROM+[S_Dur]-1),WKMAX)))
        Option:             'SKU PLU'[SKU PLU]&" "&'SKU PLU'[SKU Colour]
        Multi:              IF(LOOKUPVALUE(ItemExcl[Multi],ItemExcl[Option],'SKU PLU'[Option])="X",
                            IF(LEFT('SKU PLU'[SKU Colour],5)="Multi",1,0),
                            0)
        ROS_FINAL:          IF([Multi]=0,[ROS_AllStoreEquivalentu_PrAdj],[ROS_AllStoreEquivalentu_PrAdj]/[MultiFactor])
        ROS_VALUE_FINAL:    IF([Multi]=0,[ROS_AllStoreEquivalentV_PrAdj],[ROS_AllStoreEquivalentV_PrAdj]/[MultiFactor])
        Sell-Through FINAL: [Avg Sell-Through In Period]
        ROS_MARGIN_FINAL:   [ROS_VALUE_FINAL]/[VAT]*[Sales Margin]
        MD % SLS            IF([MD Retail]>0,IF([Sales Value]>0,[MD Retail]/[Sales Value],1),0)
        ItemExcl:           IF(LOOKUPVALUE(ItemExcl[Item EXCLUSION],ItemExcl[Option],'SKU PLU'[Option])="1",1,
                            IF(AND('SKU PLU'[WeekFrom]=0,'SKU PLU'[WeekTo]=0),2,
                            IF(LOOKUPVALUE(PromoTV[PromoType],PromoTV[Option],[Option])="TV",3,
                            IF([Effective Weeks]<=0.4*'SKU PLU'[Weeks In Period],4,0))))
        ROS_Score Value:    IF('SKU PLU'[ROS_VALUE_FINAL]>=LOOKUPVALUE(SubDep[ROS_Final_Tier1],SubDep[Sub Department],[Sub Department]),1.5,
                            IF('SKU PLU'[ROS_VALUE_FINAL]>=LOOKUPVALUE(SubDep[ROS_Final_Tier2],SubDep[Sub Department],[Sub Department]),1,
                            IF('SKU PLU'[ROS_VALUE_FINAL]>=LOOKUPVALUE(SubDep[ROS_Final_Tier3],SubDep[Sub Department],[Sub Department]),0.5,
                            0)))
        Sell-Through Score: IF([Sell-Through FINAL]>=LOOKUPVALUE(KPI_Overwrite[ST Tier 1 Final],KPI_Overwrite[Merch Group],[MG Group]),1,if([Sell-Through FINAL]>=LOOKUPVALUE(KPI_Overwrite[ST Tier 2 Final],KPI_Overwrite[Merch Group],[MG Group]),0.5,0))
        Margin_Score:       IF('SKU PLU'[ROS_MARGIN_FINAL]>=LOOKUPVALUE(SubDep[ROSM_Final_Tier1],SubDep[Sub Department],[Sub Department]),1.5,
                            IF('SKU PLU'[ROS_MARGIN_FINAL]>=LOOKUPVALUE(SubDep[ROSM_Final_Tier2],SubDep[Sub Department],[Sub Department]),1,
                            IF('SKU PLU'[ROS_MARGIN_FINAL]>=LOOKUPVALUE(SubDep[ROSM_Final_Tier3],SubDep[Sub Department],[Sub Department]),0.5,
0)))    TOTAL SCORE V:      max(IF('SKU PLU'[SKU Merch Type] = "Y",'SKU PLU'[ROS_Score Value]+'SKU PLU'[Margin_Score],
                            'SKU PLU'[ROS_Score Value]+'SKU PLU'[Sell-Through Score]+'SKU PLU'[Margin_Score]),
                            LOOKUPVALUE('PRH data'[MinMID],'PRH data'[SKU PLU],[SKU PLU],'PRH data'[SKU Colour],[SKU Colour]))
        HIT / MID / KIT V:  IF([SKU Merch Type] = "Y",
                            IF('SKU PLU'[TOTAL SCORE V]>2,"01_HIT",
                            IF('SKU PLU'[TOTAL SCORE V]>1,"02_MID","03_KIT")),
                            IF('SKU PLU'[TOTAL SCORE V]>=3,"01_HIT",
                            IF('SKU PLU'[TOTAL SCORE V]>=1.5,"02_MID","03_KIT")))
        GradeFINAL:         IF(LOOKUPVALUE(ItemExcl[Grade Overwrite],ItemExcl[Option],[Option])="X",'SKU PLU'[Grade]*1,LOOKUPVALUE(ItemExcl[Grade Overwrite],ItemExcl[Option],[Option])*1)
        Weeks In Period:    LOOKUPVALUE(PCAL[idx],PCAL[Wk_Key],'SKU PLU'[WeekTo])
                            -LOOKUPVALUE(PCAL[IDX],PCAL[Wk_Key],'SKU PLU'[WeekFrom])+1
        ROS Rank:           RANKX(FILTER('SKU PLU','SKU PLU'[ItemExcl]=0),'SKU PLU'[ROS_VALUE_FINAL],,DESC)
        ST Rank:            RANKX(FILTER('SKU PLU',AND('SKU PLU'[ItemExcl]=0,[SKU Merch Season]=EARLIER([SKU Merch Season]))),[Sell-Through FINAL],,DESC)
        MD Rank:            RANKX(FILTER('SKU PLU',AND('SKU PLU'[ItemExcl]=0,[SKU Merch Season]=EARLIER([SKU Merch Season]))),[MD % SLS],,ASC)
        Sum:                'SKU PLU'[ROS Rank]+'SKU PLU'[ST Rank]+'SKU PLU'[MD Rank]
        Final Rank:         RANKX(FILTER('SKU PLU','SKU PLU'[ItemExcl]=0),[sum],,ASC)+(3-'SKU PLU'[TOTAL SCORE V])/0.5*10000
        MG Group            IF([SKU Merch Group] = "Y","Y","NonY")
        :

    :return:
    """

def Inventory():
    """
    Columns from the Inventory files (for all hier) plus:

    InStock:                IF(	OR(Inventory[SalesU]>0,Inventory[CSOHU]>[MinStrStk]),1,0)
    Wk_Key:                 LOOKUPVALUE(PCAL[Wk_Key],PCAL[Pl Year],[Pl Year],PCAL[Pl Week],[Pl Week])
    Option:                 Inventory[SKU PLU]&" "&Inventory[SKU Colour]
    MinInStock:             IF(OR(Inventory[SalesU]>0,Inventory[CSOHU]>0),1,0)
    STR Company:            LOOKUPVALUE(Grading[STR Company],Grading[Store STR],Inventory[STR Number])
    PeriodType:             IF(MAXX(
                            CALCULATETABLE(PromoReg,
                            PromoReg[Option]=EARLIER([Option]),
                            PromoReg[PCAL_WEEK_KEY]=EARLIER([Wk_Key]),
                            PromoReg[STR Company]=EARLIER([STR Company])),PromoReg[PCAL_WEEK_KEY])
                            >1,"Promo",
                            IF(IF(Inventory[CSOHV]/Inventory[CSOHU]=BLANK(),Inventory[SalesV]/Inventory[SalesU],Inventory[CSOHV]/Inventory[CSOHU])>=LOOKUPVALUE('SKU PLU'[SKU PPL Initial Price],'SKU PLU'[Option],[Option])*0.85,"Regular",
                            "Markdown"))
    SalesUPromoAdj:         IF(Inventory[PeriodType]="Promo",[Salesu]/[PromoFactor],[Salesu])
    SalesVPromoAdj:         IF(Inventory[PeriodType]="Promo",[SalesV]/[PromoFactor],[SalesV])


    :return:
    """

def MD():
    """
    One more column:
        Option:             [SKU PLU]&" "&[SKU Colour]

    :return:
    """

def ItemExcl():
    # no additional cols
    pass

def SubDep():
    """
    Additional columns:
        ItemCount:          CALCULATE(DISTINCTCOUNT('SKU PLU'[Option]),'SKU PLU'[Sub Department]=EARLIER([Sub Department]),'SKU PLU'[ItemExcl]=0)
        ROS_V_Tier1:        var Tier1Cut = LOOKUPVALUE(PerfDep[Sls Tier 1 %],PerfDep[Sub Department],SubDep[Sub Department]) return
                            var Tier1 = ROUNDUP(SubDep[ItemCount]*Tier1Cut,0)
                            VAR TempTable = CALCULATETABLE('SKU PLU','SKU PLU'[Sub Department]=EARLIER([Sub Department]),'SKU PLU'[ItemExcl]=0)

                            return
                            MINX(FILTER(ADDCOLUMNS(TempTable,"Temp",RANKX(TempTable,[ROS_VALUE_FINAL],,DESC)),[Temp]<=Tier1),[ROS_VALUE_FINAL])
        ROS_V_Tier2:        =var Tier2Cut = LOOKUPVALUE(PerfDep[Sls Tier 2 %],PerfDep[Sub Department],SubDep[Sub Department]) return
                            var Tier2 = ROUNDUP(SubDep[ItemCount]*Tier2Cut,0)
                            VAR TempTable = CALCULATETABLE('SKU PLU','SKU PLU'[Sub Department]=EARLIER([Sub Department]),'SKU PLU'[ItemExcl]=0)

                            return
                            MINX(FILTER(ADDCOLUMNS(TempTable,"Temp",RANKX(TempTable,[ROS_VALUE_FINAL],,DESC)),[Temp]<=Tier2),[ROS_VALUE_FINAL])
        ROS_Man_Tier1:      SUBSTITUTE(LOOKUPVALUE(ROS_overwrite[ROS Tier 1 Man],ROS_overwrite[Sub Department],[Sub Department]),",",".")
        ROS_Man_Tier2:      SUBSTITUTE(LOOKUPVALUE(ROS_overwrite[ROS Tier 2 Man],ROS_overwrite[Sub Department],[Sub Department]),",",".")
        ROS_Final_Tier1:    IFERROR(SubDep[ROS_Man_Tier1]*1,SubDep[ROS_V_Tier1])
        ROS_Final_Tier2:    IFERROR(SubDep[ROS_Man_Tier2]*1,SubDep[ROS_V_Tier2])
    :return:
    """

def PLU_Available():
    """
    Additional columns:
        StoresCount:        LOOKUPVALUE(Grades[Cut-Off (Count / Value)],Grades[Grade],
                            LOOKUPVALUE('SKU PLU'[GradeFINAL],'SKU PLU'[Option],[Option]))*CALCULATE(DISTINCTCOUNT(Grading[Store STR]),Grading[STR Company]=EARLIER([STR Company]))
        Option:             'SKU PLU Avail'[SKU PLU]&" "&'SKU PLU Avail'[SKU Colour]

    :return:
    """

def PromoReg():
    """
    Additional columns:
        Option:             PromoReg[SKU PLU]&" "&PromoReg[SKU Colour]

    :return:
    """

def PromoTV():
    """
    Additional columns:
        Option:             PromoTV[SKU PLU]&" "&PromoTV[SKU Colour]

    :return:
    """

def ROS_Overwrite():
    # no additional cols
    pass

def KPI_Overwrite():
    """
    Additional columns:
        ST Tier 1 Calc:     VAR Tier1=[ST Tier 1 %]
                            VAR MG = KPI_Overwrite[Merch Group] RETURN
                            VAR CALCTABLE = CALCULATETABLE('SKU PLU','SKU PLU'[MG Group]=MG,'SKU PLU'[ItemExcl]=0)
                            RETURN
                            VAR MINRANK =COUNTX(CALCTABLE,[Option])*Tier1
                            RETURN
                            MINX(FILTER(ADDCOLUMNS(CALCTABLE,"RANK",RANKX(CALCTABLE,[Sell-Through FINAL],,DESC)),[RANK]<=MINRANK),[Sell-Through FINAL])
        ST Tier 2 Calc:     VAR Tier2=[ST Tier 2 %]
                            var MG = KPI_Overwrite[Merch Group]
                            RETURN
                            VAR CALCTABLE = CALCULATETABLE('SKU PLU','SKU PLU'[MG Group]=MG,'SKU PLU'[ItemExcl]=0)
                            RETURN
                            VAR MINRANK =COUNTX(CALCTABLE,[Option])*Tier2
                            RETURN
                            MINX(FILTER(ADDCOLUMNS(CALCTABLE,"RANK",RANKX(CALCTABLE,[Sell-Through FINAL],,DESC)),[RANK]<=MINRANK),[Sell-Through FINAL])
        MD Tier 1 Calc:     VAR Tier1Cut = [MD Tier 1 %]
                            VAR MerchSeason=KPI_Overwrite[Merch Group] RETURN
                            CALCULATE(([MD Retail]/[Sales Value])*Tier1Cut,
                            ALL('SKU PLU'),'SKU PLU'[ItemExcl]=0,'SKU PLU'[MG Group]=MerchSeason)
        MD Tier 2 Calc:     VAR Tier2Cut = [MD Tier 2 %]
                            VAR MerchSeason=KPI_Overwrite[Merch Group] RETURN
                            CALCULATE(([MD Retail]/[Sales Value])*Tier2Cut,
                            ALL('SKU PLU'),'SKU PLU'[ItemExcl]=0,'SKU PLU'[MG Group]=MerchSeason)
        ST Tier 1 Final:    IFERROR(VALUE(sUBSTITUTE(KPI_Overwrite[ST Tier 1 Man],",","."))*1,KPI_Overwrite[ST Tier 1 Calc])
        ST Tier 2 Final:    IFERROR(VALUE(SUBSTITUTE(KPI_Overwrite[ST Tier 2 Man],",","."))*1,KPI_Overwrite[ST Tier 2 Calc])
        MD Tier 1 Final:    IFERROR(VALUE(SUBSTITUTE(KPI_Overwrite[MD Tier 1 Man],",","."))*1,KPI_Overwrite[MD Tier 1 Calc])
        MD Tier 2 Final:    IFERROR(VALUE(SUBSTITUTE(KPI_Overwrite[MD Tier 2 Man],",","."))*1,KPI_Overwrite[MD Tier 2 Calc])

    :return:
    """

def PRH():
    """
    One more column
        HKM:                LOOKUPVALUE('SKU PLU'[HIT / MID / KIT V],'SKU PLU'[SKU PLU],[SKU PLU],'SKU PLU'[SKU Colour],[SKU Colour],[ItemExcl],0)
    :return:
    """

def PerfDep():
    """
    Additional columns:
        Sales Perf:         PerfDep[Sales Retail ACT]/PerfDep[Sales Retail AP]-1
        Sls Tier 1:         IF(PerfDep[Sales Perf]>0.05,0.4,IF(PerfDep[Sales Perf]>-0.05,0.3,0.2))
        Sls Tier 2 %:       IF(PerfDep[Sales Perf]>0.05,0.8,IF(PerfDep[Sales Perf]>-0.05,0.65,0.5))
    :return:
    """