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