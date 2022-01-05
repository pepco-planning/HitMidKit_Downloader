import pandas as pd
import sqlite3
from sqlalchemy import create_engine

path = r"c:\Mariusz\MyProjects\HitMidKit_Downloader\input data\a Toys\\"
fileInv = 'inventory.csv'
df_inventory = pd.read_csv(path + fileInv)

conn = sqlite3.connect('hmk.db')
engine = create_engine('sqlite:///hmk.db', echo=False)


df.to_sql('inventory',con=engine, if_exists='append', index=False)


def createDb():
    conn = sqlite3.connect('hmk.db')
    c = conn.cursor()
    c.execute("""CREATE TABLE inventory (
                'SKU PLU' integer,
                'SKU Colour' text,
                'STR Number' integer,
                'Pl Year' integer,
                'Pl Week' text,
                'SalesU' real,
                'SalesV' real,
                'SalesM' real,
                'CSOHU' real,
                'CSOHV' real
                )""")
    conn.commit()
    conn.close()

def selectDb():
    conn = sqlite3.connect('hmk.db')
    c = conn.cursor()
    c.execute("""SELECT [SKU PLU], [SalesV] FROM inventory""")
    xxx = c.fetchall()[:3]

    conn.commit()
    conn.close()

pd.read_sql('inventory',engine).to_excel('bleble.xlsx',index=False)

def drop_db():
    c = conn.cursor()
    c.execute("""DROP TABLE inventory
                """)
    conn.commit()

col = ['SKU PLU', 'SKU Colour', 'STR Number', 'Pl Year', 'Pl Week', 'SalesU',
       'SalesV', 'SalesM', 'CSOHU', 'CSOHV']

"""
zipfile.ZipFile:
"""
### Open the files

import pandas as pd
import glob
import zipfile
from pathlib import Path
import os

dir_a = r"\\10.2.5.140\zasoby\Planowanie\PERSONAL FOLDERS\Mariusz Borycki\HitMidKit_DataBase\PCAL 2022 Q2 AW21"
dir_b = r"\\10.2.5.140\zasoby\Planowanie\PERSONAL FOLDERS\Mariusz Borycki\HitMidKit_DataBase\PCAL 2022 Q2 AW21_W26"


def save_compressed_df(df, dirPath, fileName):
    """Save a Pandas dataframe as a zipped .csv file.

    Parameters
    ----------
    df : pandas.core.frame.DataFrame
        Input dataframe.
    dirPath : str or pathlib.PosixPath
        Parent directory of the zipped file.
    fileName : str
        File name without extension.
    """

    dirPath = Path(dirPath)
    path_zip = dirPath / f'{fileName}.zip'
    txt = df.to_csv(index=False, sep='\t', chunksize=10000)  # chunksize=1000
    with zipfile.ZipFile(path_zip, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f'{fileName}', txt)


files_a = glob.glob(dir_a + "/*.csv")
files_b = glob.glob(dir_b + "/*.csv")

# # temporary: start from 1.1.3.1
# files_a = files_a[33:]
# files_b = files_b[33:]

for filename_a in files_a:
    for filename_b in files_b:
        f_name_csv_a = filename_a.split("\\")[-1]
        f_name_csv_b = filename_b.split("\\")[-1]
        if f_name_csv_b == f_name_csv_a:
            df_a = pd.read_csv(filename_a)
            df_b = pd.read_csv(filename_b)
            df = pd.concat([df_a, df_b], ignore_index=True)
            dirPath = r"C:\Users\mborycki\Desktop\Python Projects\Database verification for HitMidKit"
            save_compressed_df(df, dirPath, f_name_csv_b)
            del df, df_a, df_b

### Missing files

path_a = r"\\10.2.5.140\zasoby\Planowanie\PERSONAL FOLDERS\Mariusz Borycki\HitMidKit_DataBase\PCAL 2022 Q2 AW21_W26"
path_b = r"\\10.2.5.140\zasoby\Planowanie\PERSONAL FOLDERS\Mariusz Borycki\HitMidKit_DataBase\PCAL 2022 Q2 AW21"
missing = ['1.1.3.1_HitMidKit.csv',
           '1.2.1.5_HitMidKit.csv',
           '1.2.1.6_HitMidKit.csv',
           '1.2.2.1_HitMidKit.csv',
           '1.2.3.6_HitMidKit.csv',
           '1.2.3.7_HitMidKit.csv',
           '1.2.4.1_HitMidKit.csv',
           '1.2.4.2_HitMidKit.csv']

p_a = list()
p_b = list()

for p in missing:
    new_dir = path_a + "\\" + p
    p_a.append(new_dir)

    new_dir_b = path_b + "\\" + p
    p_b.append(new_dir_b)

for filename_a in p_a:
    for filename_b in p_b:
        f_name_csv_a = filename_a.split("\\")[-1]
        f_name_csv_b = filename_b.split("\\")[-1]
        if f_name_csv_b == f_name_csv_a:
            df_a = pd.read_csv(filename_a)
            df_b = pd.read_csv(filename_b)
            df = pd.concat([df_a, df_b], ignore_index=True)
            dirPath = r"C:\Users\mborycki\Desktop\Python Projects\Database verification for HitMidKit"
            save_compressed_df(df, dirPath, f_name_csv_b)
            del df, df_a, df_b

### Comparing zip + csv

# zip

files_a = glob.glob(dir_a + "/*.csv")
files_b = glob.glob(dir_b + "/*.csv")
files_zip = glob.glob(os.getcwd() + "/*.zip")

for filename_a in files_a:
    for filename_b in files_b:
        f_name_csv_a = filename_a.split("\\")[-1]
        f_name_csv_b = filename_b.split("\\")[-1]
        if f_name_csv_b == f_name_csv_a:
            df_a = pd.read_csv(filename_a)
            df_b = pd.read_csv(filename_b)
            df = pd.concat([df_a, df_b], ignore_index=True)

            f_name = filename_a.split('\\')[-1]
            print(f"DataFrame: {f_name}. Shape: {df.shape}")
            print("----------\n")

for filename in files_zip:
    df = pd.read_csv(filename)
    f_name = filename.split('\\')[-1]
    print(f"DataFrame: {f_name}. Shape: {df.shape}")
    print("----------\n")
#     print(filename)

