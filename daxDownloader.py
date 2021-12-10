# w path.append jest istotna ścieżka
# muszą się w niej znajdować 2 pliki:
# 1. Microsoft.AnalysisServices.AdomdClient.dll
# 2. Microsoft.AnalysisServices.dll
# Plików szukaj w C:\Windows\assembly\GAC_MSIL\Microsoft.AnalysisServices.(nazwa pliku)\

import pandas as pd
from sys import path
path.append(r"dll")
from pyadomd import Pyadomd


def dataFrameFromTabular(query):
    connStr = "Provider=MSOLAP;Data Source=LB-P-WE-AS;Catalog=PEPCODW"

    conn = Pyadomd(connStr)
    conn.open()
    cursor = conn.cursor()
    cursor.execute(query)
    cursor.arraysize = 5000
    df = pd.DataFrame(cursor.fetchall())
    conn.close()

    return df