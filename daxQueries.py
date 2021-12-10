def weeks(startEndWeeks):
    daxStartWeek = startEndWeeks[0][1:5] + startEndWeeks[0][6:8]
    daxEndWeek = startEndWeeks[1][1:5] + startEndWeeks[1][6:8]

    return """
            EVALUATE
            SUMMARIZECOLUMNS (
                'Planning Calendar PCAL'[PCAL_WEEK_KEY],
                FILTER (
                    VALUES ( 'Planning Calendar PCAL'[PCAL_WEEK_KEY] ),
                    'Planning Calendar PCAL'[PCAL_WEEK_KEY] >= """ + str(daxStartWeek) + """
                        && 'Planning Calendar PCAL'[PCAL_WEEK_KEY] <= """ + str(daxEndWeek) + """
                )
            )
    """

def prh(dep):
    z = str(dep)
    print(z)
    return """
            DEFINE
            var PRH =
                Filter(
                    all('Product Hierarchy PRH'[PRH Department],'Product Hierarchy PRH'[PRH Category]),
                    'Product Hierarchy PRH'[PRH Department] in {"""" & z  """"} 
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