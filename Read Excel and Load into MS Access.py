import pyodbc
import pandas as pd
import urllib
from sqlalchemy import create_engine
import openpyxl
import time
import os
import glob

print('Opening Excel...')
list_of_files = glob.glob('Pathofexcelsheetdownloadfolder/*.xlsx') # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getctime)
df= pd.read_excel(latest_file)

print('Opening Access...')

conn_str = (r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
            r'DBQ=PathofMSAccess.accdb;')
connurl = f'access+pyodbc:///?odbc_connect={urllib.parse.quote_plus(conn_str)}'
acc_engine = create_engine(connurl)

print('Writing to Access..')

df.to_sql('NameofimporttableinMSAccess', acc_engine, if_exists='replace')


with acc_engine.begin() as cn:
   sql = """INSERT INTO MainTableinMSAccess ( field1, field2, ...)
            SELECT  I.field1, I.field2, ...
            FROM importtable I
            WHERE NOT EXISTS 
                (SELECT 1 FROM MainTable A
                 WHERE I.JionField = A.JoinField)"""

   cn.execute(sql)

print('Write Complete.')
