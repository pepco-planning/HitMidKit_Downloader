import daxDownloader as td
import daxQueries as daxQ
import pandas as pd

# dep = '"' + "e Gravelights" + '"'
departments = ['a Baby Girls Outerwear', 'b Baby Boys Outerwear', 'c Baby Girls Basics', 'd Baby Boys Basics',
               'e Younger Girls Outerwear', 'f Younger Boys Outerwear', 'g Older Girls Outerwear',
               'h Older Boys Outerwear', 'i Ladies Outerwear', 'j Mens Outerwear', 'a Girls Basics',
               'b Boys Basics', 'c Ladies Basics', 'd Mens Basics', 'e Kids Nightwear', 'f Adult Nightwear',
               'a Accessories', 'b Baby Accessories', 'a Kids Footwear', 'b Adult Footwear', 'a Kitchenware Cooking',
               'b Dining Room', 'c Bathroom', 'd Living Room', 'e Home Deco', 'f Home Tex', 'g Utility', 'a Toys',
               'a Gardening', 'b Tourism', 'c Luggage', 'd Easter', 'e Gravelights', 'f Christmas Decorations', 'g Halloween',
               'a House Electrics', 'b Stationery', 'c Festive Accessories', 'd Pet'
]

dep = "a Toys"
maxWeekNo = 52
startWeek = 202219
endWeek = 202223

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

    print(week)





#df =  td.dataFrameFromTabular(daxQ.sku_plu(startWeek,endWeek,dep))


# print(df.loc[0])
# print("---------")
# print(df)


