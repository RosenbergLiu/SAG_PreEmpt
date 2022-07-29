from asyncio.windows_events import NULL
import psycopg2 as db_connect
import os
from datetime import datetime
import pandas as pd
import openpyxl



connection = db_connect.connect(
    host="172.20.190.105",
    user="postgres",
    password="731797",
    database="asko",
    port='5432'
    )
connection.autocommit =True
cursor = connection.cursor()        

def insertRelease(sap,date,part,quantity,DN):
    if DN ==None:
        delivery_note = "null"
    else:
        delivery_note = DN
    query = f'''
    INSERT INTO 
        public.RELEASE (sap, release_date, part, quantity, delivery_note) 
    VALUES 
        ({sap},'{date}',{part},{quantity},{delivery_note}) 
    ON CONFLICT (sap,part) 
    DO UPDATE 
    SET 
        sap=EXCLUDED.sap, 
        release_date=EXCLUDED.release_date, 
        part = EXCLUDED.part, 
        delivery_note = EXCLUDED.delivery_note
    ;'''
    cursor.execute(query)
    connection.commit

def deleteRelease(sap,part):
    query = f'''
    DELETE FROM public.RELEASE 
    WHERE sap = {sap} AND part = {part}
    ;'''
    cursor.execute(query)
    connection.commit
def importToRelease(TXT):
    raw= pd.read_csv(TXT,delim_whitespace=True)
    target = r'C:/Users/roshan.liu/Scripts/DATA/SAP_OUTPUT/dbprovicional.csv'
    raw.iloc[:,:1].to_csv(target,header=False)
    df=pd.read_csv(target)
    rowcount=len(df.index)
    for row in range(0,rowcount-1):
        v = df.loc[row]['Item']
        if v >4000000000:
            r=row+1
            while df.loc[r]['Item'] < 4000000000:
                if df.loc[r]['GI'] !='PC':
                    
                    DATE= TXT.split('/')[-1][0:10]
                    DN =None
                    SAP=v
                    QTY=df.loc[r]['BUn'].split(',')[0]
                    PART = df.loc[r]['GI']
                    print(SAP,DATE,PART,QTY,DN)
                    insertRelease(SAP,DATE,PART,QTY,DN)
                else:
                    DATE= TXT.split('/')[-1][0:10]
                    DN = df.loc[r]['Date']
                    SAP=v
                    QTY=0
                    PART=df.loc[r]['Created']
                    print(SAP,DATE,PART,QTY,DN)
                    insertRelease(SAP,DATE,PART,QTY,DN)
                r=r+1
                if r>rowcount-1:
                    break

importToRelease(r'C:/Users/roshan.liu/Scripts/DATA/SAP_OUTPUT/2022-07-22-preempt.txt')