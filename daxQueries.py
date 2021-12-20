def inventory(startWeek,endWeek,currentWeek,minPar,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>= """ + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            var pcal2 = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                        'Planning Calendar PCAL'[PCAL_WEEK_KEY]=""" + str(currentWeek) + """)
            
            var 
            StartWk =
                CALCULATE(
                    min('Planning Calendar PCAL'[Issue Date]),
                    pcal)
            
            var
            EndWk = 
                CALCULATE(
                    max('Planning Calendar PCAL'[Issue Date]),
                    pcal)
            
            var
            str2 = 
                
                filter(
                    SUMMARIZECOLUMNS(
                        'Stores STR'[STR Number],
                    filter(
                        all('Stores STR'[STR Number],'Stores STR'[STR Opening Date],'Stores STR'[STR Closure Date]),
                        'Stores STR'[STR Opening Date] <= StartWk
                        &&
                        ('Stores STR'[STR Closure Date] >= EndWk
                        ||
                        'Stores STR'[STR Closure Date] = DATE(1900,1,1)
                        )),pcal,
                        "Sls",[Sales Retail Report dsale]), [Sls] > 0)
            
            var SKU = 
                filter(
                all(
                    'Products SKU'[SKU Department],
                    'Products SKU'[SKU Merch Group],
                    'Products SKU'[SKU Merch Season Type],
                    'Products SKU'[SKU Merch To Season],
                    'Products SKU'[SKU Sourcing]),
                    'Products SKU'[SKU Department] = """ + dep_name + """
                    --"",
                    --"",
                    --""
                )
            
            var SKUfin = 
                filter(
                    SUMMARIZECOLUMNS(
                    'Products SKU'[sku plu],
                        'Products SKU'[SKU Colour],SKU,"slstemp",CALCULATE([Sales Qty dsale],pcal)),[slstemp]>""" + str(minPar) + """)
            
            EVALUATE
            SUMMARIZECOLUMNS(
                'Products SKU'[SKU PLU],
                'Products SKU'[SKU Colour],
                'Stores STR'[STR Number],
                'Planning Calendar PCAL'[Pl Year],
                'Planning Calendar PCAL'[Pl Week],
                skufin,
                pcal,
                pcal2,
                str2,
                "SalesU",[Sales Qty dsale],
                "SalesV",[Sales Retail Report dsale],
                "SalesM",[Sales Margin Value Estimated dsale],
                "CSOHU",[Closing Stock on Hand Qty dcssd],
                "CSOHV",[Closing Stock on Hand Retail Report dcssd]
            )
    """

def grading(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>= """ + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            var 
            StartWk =
                CALCULATE(
                    min('Planning Calendar PCAL'[Issue Date]),
                    pcal)
            
            var
            EndWk = 
                CALCULATE(
                    max('Planning Calendar PCAL'[Issue Date]),
                    pcal)
            
            var
            str2 = 
                
                filter(
                    SUMMARIZECOLUMNS(
                        'Stores STR'[STR Number],
                    filter(
                        all('Stores STR'[STR Number],'Stores STR'[STR Opening Date],'Stores STR'[STR Closure Date]),
                        'Stores STR'[STR Opening Date] <= StartWk
                        &&
                        ('Stores STR'[STR Closure Date] >= EndWk
                        ||
                        'Stores STR'[STR Closure Date] = DATE(1900,1,1)
                        )),pcal,
                    "Sls",
                    [Sales Retail Report dsale]),
                    [Sls] > 0)
            
            var SKU = 
                filter(
                values(
                    'Products SKU'[SKU Department]),
                    'Products SKU'[SKU Department] = """ + dep_name + """
                )
                
            EVALUATE
            SUMMARIZECOLUMNS (
                'Stores STR'[STR Number],
                'Stores STR'[STR Company],
                pcal,
                str2,
                sku,
                "DepSalesV", [Sales Retail Report dsale]
            )
    """

def pcal(startWeek,endWeek,dep):

    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>=""" + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            EVALUATE
            SUMMARIZECOLUMNS (
                'Planning Calendar PCAL'[Pl Year],
                'Planning Calendar PCAL'[Pl Season],
                'Planning Calendar PCAL'[Pl Quarter],
                'Planning Calendar PCAL'[Pl Month],
                'Planning Calendar PCAL'[Pl Week],
                'Planning Calendar PCAL'[PCAL_WEEK_KEY],
                pcal
            )
            
            ORDER BY 
            'Planning Calendar PCAL'[PCAL_WEEK_KEY]
    """

def prh(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var PRH =
                Filter(
                    all('Product Hierarchy PRH'[PRH Department],'Product Hierarchy PRH'[PRH Category]),
                    'Product Hierarchy PRH'[PRH Department] in {""" + dep_name + """} 
                    &&
                    'Product Hierarchy PRH'[PRH Category] <> BLANK())
                    
            EVALUATE
            SUMMARIZECOLUMNS(
                'Product Hierarchy PRH'[PRH PEPCO],
                'Product Hierarchy PRH'[PRH Division],
                'Product Hierarchy PRH'[PRH Group],
                'Product Hierarchy PRH'[PRH Department],
                'Product Hierarchy PRH'[PRH Sub Department],
                'Product Hierarchy PRH'[PRH Category],
                PRH
            )
            ORDER BY
                'Product Hierarchy PRH'[PRH PEPCO],
                'Product Hierarchy PRH'[PRH Division],
                'Product Hierarchy PRH'[PRH Group],
                'Product Hierarchy PRH'[PRH Department],
                'Product Hierarchy PRH'[PRH Sub Department],
                'Product Hierarchy PRH'[PRH Category]
            """

def sku_plu(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'

    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>=""" + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            var SKU = 
                filter(
                all(
                    'Products SKU'[SKU Department],
                    'Products SKU'[SKU Merch Group],
                    'Products SKU'[SKU Merch Season Type],
                    'Products SKU'[SKU Merch To Season],
                    'Products SKU'[SKU Sourcing]),
                    'Products SKU'[SKU Department] = """ + dep_name + """
                    --"&MG&"
                    --"&MST&"
                    --"&SeasonTo&"
                )
            
            var SKUfin = 
                filter(
                    SUMMARIZECOLUMNS(
                    'Products SKU'[sku plu],
                        'Products SKU'[SKU Colour],SKU,"slstemp",CALCULATE([Sales Qty dsale],pcal)),[slstemp]>2000)
            
            EVALUATE
            SUMMARIZECOLUMNS (
                'Products SKU'[SKU PLU],
                'Products SKU'[SKU Colour],
                'Products SKU'[SKU Name],
                'Products SKU'[SKU Store Grade],
                'Products SKU'[SKU Category],
                'Products SKU'[SKU Sub Department],
                'Products SKU'[SKU Price Group],
                'Products SKU'[SKU Style Type],
                'Products SKU'[SKU PPL Initial Retail Price],
                'Products SKU'[SKU Merch Season Type],
                'Products SKU'[SKU Merch Group],
                'Products SKU'[SKU Merch Type],
                'Products SKU'[SKU Merch To Season],
                pcal,
                SKUfin,
                "Sales Margin", [Sales Margin Estimated dsale],
                "VAT", [Sales Retail Report dsale] / [Sales Net Retail Report dsale]
            )
    """

def md(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>= """ + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            
            var SKU = 
                filter(
                all(
                    'Products SKU'[SKU Department],
                    'Products SKU'[SKU Merch Group],
                    'Products SKU'[SKU Merch Season Type],
                    'Products SKU'[SKU Merch To Season],
                    'Products SKU'[SKU Sourcing]),
                    'Products SKU'[SKU Department] = """ + dep_name + """
                    --"&MG&"
                    --"&MST&"
                    --"&SeasonTo&"
                )
            
            
            var SKUfin = 
                filter(
                    SUMMARIZECOLUMNS(
                    'Products SKU'[sku plu],
                        'Products SKU'[SKU Colour],SKU,"slstemp",CALCULATE([Sales Qty dsale],pcal)),[slstemp]>2000)
            
            
            EVALUATE
            SUMMARIZECOLUMNS (
                'Products SKU'[SKU PLU],
                'Products SKU'[SKU Colour],
                pcal,
                SKUfin,
                "Markdown Retail Report", [Markdown HQ Retail Report dprch]
            )
    """

def plu_available(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>=""" + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            var SKU = 
                filter(
                all(
                    'Products SKU'[SKU Department],
                    'Products SKU'[SKU Merch Group],
                    'Products SKU'[SKU Merch Season Type],
                    'Products SKU'[SKU Merch To Season],
                    'Products SKU'[SKU Sourcing]),
                    'Products SKU'[SKU Department] = """ + dep_name + """
                    --"&MG&"
                    --"&MST&"
                    --"&SeasonTo&"
                )
            
            var SKUfin = 
                filter(
                    SUMMARIZECOLUMNS(
                    'Products SKU'[sku plu],
                        'Products SKU'[SKU Colour],SKU,"slstemp",CALCULATE([Sales Qty dsale],pcal)),[slstemp]>2000)
            
            EVALUATE
            SUMMARIZECOLUMNS(
                'Products SKU'[SKU PLU],
                'Products SKU'[SKU Colour],
                'Products SKU'[SKU PPL Available],
                'Products SKU'[SKU PSK Available],
                'Products SKU'[SKU PCZ Available],
                'Products SKU'[SKU PHU Available],
                'Products SKU'[SKU PRO Available],
                'Products SKU'[SKU PHR Available],
                'Products SKU'[SKU PSI Available],
                'Products SKU'[SKU PLT Available],
                'Products SKU'[SKU PLV Available],
                'Products SKU'[SKU PEE Available],
                'Products SKU'[SKU PBG Available],
                'Products SKU'[SKU PIT Available],
                'Products SKU'[SKU PRS Available],
                SKUfin)
    """

def promo_reg(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>=""" + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            
            var SKU = 
                filter(
                all(
                    'Products SKU'[SKU Department],
                    'Products SKU'[SKU Merch Group],
                    'Products SKU'[SKU Merch Season Type],
                    'Products SKU'[SKU Merch To Season],
                    'Products SKU'[SKU Sourcing]),
                    'Products SKU'[SKU Department] = """ + dep_name + """
                    --"&MG&"
                    --"&MST&"
                    --"&SeasonTo&"
                )
            
            
            var SKUfin = 
                filter(
                    SUMMARIZECOLUMNS(
                    'Products SKU'[sku plu],
                        'Products SKU'[SKU Colour],SKU,"slstemp",CALCULATE([Sales Qty dsale],pcal)),[slstemp]>2000)
            
            EVALUATE
            SUMMARIZECOLUMNS (
                'Products SKU'[SKU PLU],
                'Products SKU'[SKU Colour],
                'Stores STR'[STR Company],
                'Promo Codes PCD'[PCD Promo Code],
                'Planning Calendar PCAL'[PCAL_WEEK_KEY],
                pcal,
                SKUfin,
                "promo", CALCULATE (
                    [Sales Qty dsale],
                    'Promo Codes PCD',
                    'Promo Codes PCD'[PCD Promo Type] IN { "HSP", "P", "PW"}
                )
            )
    """

def promo_tv(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>=""" + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            
            var SKU = 
                filter(
                all(
                    'Products SKU'[SKU Department],
                    'Products SKU'[SKU Merch Group],
                    'Products SKU'[SKU Merch Season Type],
                    'Products SKU'[SKU Merch To Season],
                    'Products SKU'[SKU Sourcing]),
                    'Products SKU'[SKU Department] = """ + dep_name + """
                    --"&MG&"
                    --"&MST&"
                    --"&SeasonTo&"
                )
            
            
            var SKUfin = 
                filter(
                    SUMMARIZECOLUMNS(
                    'Products SKU'[sku plu],
                        'Products SKU'[SKU Colour],SKU,"slstemp",CALCULATE([Sales Qty dsale],pcal)),[slstemp]>2000)
            
            
            EVALUATE
            SUMMARIZECOLUMNS (
                'Products SKU'[SKU PLU],
                'Products SKU'[SKU Colour],
                pcal,
                SKUfin,
                "promo", CALCULATE (
                    [Sales Qty dsale],
                    'Promo Codes PCD',
                    'Promo Codes PCD'[PCD Promo Type] IN {"TV"}
                )
            )
    """

def prh_data(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>=""" + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
            var SKU = 
                filter(
                all(
                    'Products SKU'[SKU Department],
                    'Products SKU'[SKU Merch Group],
                    'Products SKU'[SKU Merch Season Type],
                    'Products SKU'[SKU Merch To Season],
                    'Products SKU'[SKU Sourcing]),
                    'Products SKU'[SKU Department] = """ + dep_name + """
                    --"&MG&"
                    --"&MST&"
                    --"&SeasonTo&"
                )	
            
            EVALUATE
            SUMMARIZECOLUMNS (
                'Products SKU'[SKU PLU],
                'Products SKU'[SKU Colour],
                'Products SKU'[SKU Category],
                pcal, SKU,
                "Sales V",[Sales Retail Report dsale],
                "Sales U",[Sales Qty dsale],
                "Markdown V",[Markdown Total Retail Report dprch],
                "Sales Margin V",[Sales Margin Value Estimated dsale]
            )
    """

def perf_dep(startWeek,endWeek,dep):
    dep_name = '"' + dep + '"'
    return """
            DEFINE
            var pcal = 
                filter(
                    values('Planning Calendar PCAL'[PCAL_WEEK_KEY]),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]>=""" + str(startWeek) + """
                    &&
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY]<=""" + str(endWeek) + """)
            
                VAR prh =
                    filter(values (
                        'Product Hierarchy PRH'[PRH Department]),
                        'Product Hierarchy PRH'[PRH Department] = """ + dep_name + """
                    )
            EVALUATE
            SUMMARIZECOLUMNS(
                'Product Hierarchy PRH'[PRH Sub Department],
                pcal,prh,
                "Sales Retail ACT", [Sales Retail Report dsale],
                "Markdown ACT",	[Markdown Total Retail Report dprch],
                "Sales Retail AP", [Sales Retail Report AP mmfp],
                "Markdown AP", [Markdown Total Retail Report dprch],
                "Inflows ACT", [Total Inflows Retail Report dinta],
                "Intake AP",[Intake Retail Report AP mmfp]
                )
    """