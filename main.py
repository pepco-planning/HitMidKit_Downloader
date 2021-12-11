import daxQueries as daxQ
import functions as f

dep = "e Gravelights"
startWeek = 202219
endWeek = 202223
path = 'c:\Mariusz\MyProjects\HitMidKit\input data\\'

#list_of_weeks = f.weeksCalculation(startWeek, endWeek, maxWeekNo=52) # w raporcie pcal znajdują się wszystkie tygodnie. Wywal tę funkcję wtedy
dax_query_list = [daxQ.grading,daxQ.sku_plu,daxQ.md,daxQ.plu_available,
                     daxQ.promo_reg,daxQ.promo_tv,daxQ.prh_data,daxQ.perf_dep,daxQ.pcal,daxQ.prh] #inventory

f.runDaxQueries(startWeek,endWeek,dep,dax_query_list,path)