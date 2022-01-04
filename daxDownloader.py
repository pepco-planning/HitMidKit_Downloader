# w path.append jest istotna ścieżka
# muszą się w niej znajdować 2 pliki:
# 1. Microsoft.AnalysisServices.AdomdClient.dll
# 2. Microsoft.AnalysisServices.dll
# Plików szukaj w C:\Windows\assembly\GAC_MSIL\Microsoft.AnalysisServices.(nazwa pliku)\

import pandas as pd
import os
from sys import path
# path.append(r"dll")
# path.append("c:\Mariusz\MyProjects\HitMidKit_Downloader\dll")
path.append(os.getcwd()+'\\dll')
from pyadomd import Pyadomd


def dataFrameFromTabular(query):
    connStr = "Provider=MSOLAP;Data Source=LB-P-WE-AS;Catalog=PEPCODW"

    conn = Pyadomd(connStr)
    conn.open()
    cursor = conn.cursor()
    cursor.execute(query)
    col_names = [i[0] for i in cursor.description]
    cursor.arraysize = 5000
    df = pd.DataFrame(cursor.fetchall(), columns=col_names)
    conn.close()

    return df