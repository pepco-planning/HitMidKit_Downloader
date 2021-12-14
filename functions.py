import daxDownloader as td

def runDaxQueries(startWeek,endWeek,dep,dax_query_list,path):
    """
    NOT USED!
    :param startWeek:
    :param endWeek:
    :param dep:
    :param dax_query_list:
    :param path:
    :return:
    """
    for query in dax_query_list:
        df = td.dataFrameFromTabular(query(startWeek, endWeek, dep))
        q_name = query.__name__ + '.csv'
        df.to_csv(path + q_name, index=False)

def runDaxQuery(startWeek,endWeek,dep,dax_query,path):
    df = td.dataFrameFromTabular(dax_query(startWeek, endWeek, dep))
    q_name = dax_query.__name__ + '.csv'
    df.to_csv(path + q_name, index=False)

departments = ['a Baby Girls Outerwear', 'b Baby Boys Outerwear', 'c Baby Girls Basics', 'd Baby Boys Basics',
               'e Younger Girls Outerwear', 'f Younger Boys Outerwear', 'g Older Girls Outerwear',
               'h Older Boys Outerwear', 'i Ladies Outerwear', 'j Mens Outerwear', 'a Girls Basics',
               'b Boys Basics', 'c Ladies Basics', 'd Mens Basics', 'e Kids Nightwear', 'f Adult Nightwear',
               'a Accessories', 'b Baby Accessories', 'a Kids Footwear', 'b Adult Footwear', 'a Kitchenware Cooking',
               'b Dining Room', 'c Bathroom', 'd Living Room', 'e Home Deco', 'f Home Tex', 'g Utility', 'a Toys',
               'a Gardening', 'b Tourism', 'c Luggage', 'd Easter', 'e Gravelights', 'f Christmas Decorations', 'g Halloween',
               'a House Electrics', 'b Stationery', 'c Festive Accessories', 'd Pet'
]

def weeksCalculation(startWeek, endWeek, maxWeekNo):
    """
    Comments:
    W raporcie pcal znajdują się wszystkie tygodnie. Funkcja nieużywana
    :param startWeek:
    :param endWeek:
    :param maxWeekNo: number of weeks in the year from startWeek variable
    :return:
    """

    list_of_weeks = []
    if int(str(startWeek)[-2:]) > int(str(endWeek)[-2:]):
        weekNumbers = maxWeekNo - int(str(startWeek)[-2:]) + int(str(endWeek)[-2:])
    else:
        weekNumbers = endWeek - startWeek

    weekRatio = 0
    for weeks in range(0,weekNumbers + 1):
        week = startWeek + weeks - weekRatio

        if int(str(week)[-2:]) == 52:
            year = int(str(week)[:4]) + 1
            startWeek = int(str(year) + "01")
            weekRatio = weeks + 1

        list_of_weeks.append(week)
        #print(week)
    return list_of_weeks

def xw_func():
    #A Function in xlwings (we can create own function there
    # def xw_func() DO NOT EXIST. I have created it just to make it easier to collapse
    @xw.func
    def hello(name):
        return f"Hi {name}!"

    if __name__ == "__main__":
        xw.Book("hitmidkit.xlsm").set_mock_caller()
        main()

def new():
    pass