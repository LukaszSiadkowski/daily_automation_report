import psycopg2, os, time, sys
import pandas as pd
from datetime import datetime, timedelta, date
import numpy as np
from pathlib import Path
import win32com.client as win32

con = psycopg2.connect(
    dbname='x',
    host='x',
    port='x',
    user='x',
    password='xxx')


class Query(str):

    #just runs the query
    def run(self, cursor):
        cursor.execute(self)
        return self

    #just saves the query to a list
    def append(self, query_list):
        query_list.append(self)
        return self

    #runs and saves query to a list
    def run_append(self, cursor, query_list):
        cursor.execute(self)
        query_list.append(self)
        return self

    #runs the query and stores the output in a table
    def run_store(self, cursor):
        time_init= time.time()
        cursor.execute(self)
        df = pd.DataFrame.from_records(
            iter(cursor), columns=[x[0] for x in cursor.description])
        print('execution time (minutes) =', round((time.time()-time_init)/60, 1))
        return df

    #run, store and append
    def run_store_append(self, cursor, query_list):
        cursor.execute(self)
        df = pd.DataFrame.from_records(
            iter(cursor), columns=[x[0] for x in cursor.description])
        query_list.append(self)
        return df



def query_to_df(query, cursor):
    return Query(query).run_store(cursor=cursor)


def inspect(df, cols):
    if len(cols) == 1 or type(cols) == str:
        return df[cols].value_counts()
    elif len(cols)>1:
        return df[cols].head()

    
try:
    xl = client.gencache.EnsureDispatch('Excel.Application')
except AttributeError:
    import os, re, sys, shutil
    MODULE_LIST = [m.__name__ for m in sys.modules.values()]
    for module in MODULE_LIST:
        if re.match(r'win32com\.gen_py\..+', module):
            del sys.modules[module]
    shutil.rmtree(os.path.join(os.environ.get('LOCALAPPDATA'), 'Temp', 'gen_py'))
    from win32com import client
    xl = client.gencache.EnsureDispatch('Excel.Application')
    
    
    
cursor = cur = con.cursor()
d = date.today() - timedelta(days=1)
dt = d.strftime("%d/%m/%y")

timeframe = "current_date -1"


germany = "'x'"
romania = "'x'"
poland = "'x'"
france = "'x'"
czech = "'x'"
italy = "'x'"
gbr = "'x'" 
deu_s = "'x'"
fra_s = "'x'"
ita_s = "'x' "
pol_t = "'x' "
cze_t = "'x'"

countries = [germany, romania, poland, france, czech, italy, gbr, deu_s, fra_s, ita_s, pol_t, cze_t]


def save_location(country):
    if country == germany:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == romania:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == poland:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == france:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == czech:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == italy:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == gbr:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)    
    if country == deu_s:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == fra_s:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == ita_s:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == pol_t:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    if country == cze_t:
        filepath = Path(f'C:/xxxx/report for {d}.csv')
        filepath.parent.mkdir(parents=True, exist_ok=True)
        master_df.to_csv(filepath)
    
        

        
def recipent(country):
    if country == germany:
        mail.To = 'x@x.x'
        mail.CC = 'x@x.x'
    if country == romania:
        mail.To = 'x@x.x'
    if country == poland:
        mail.To = 'x@x.x'
    if country == france:
        mail.To = 'x@x.x'
    if country == czech:
        mail.To = 'x@x.x'
    if country == italy:
        mail.To = 'x@x.x'
    if country == gbr:
        mail.To = 'x@x.x'
    if country == deu_s:
        mail.To = 'x@x.x'
    if country == fra_s:
        mail.To = 'x@x.x'
    if country == ita_s:
        mail.To = 'x@x.x'
    if country == pol_t:
        mail.To = 'x@x.x'
    if country == cze_t:
        mail.To = 'x@x.x'
        
def subject(country):
    if country == germany:
        mail.Subject = f'Report for {dt}'
    if country == romania:
        mail.Subject = f'Report for {dt}'
    if country == poland:
        mail.Subject = f'Report for {dt}'
    if country == france:
        mail.Subject = f'Report for {dt}'
    if country == czech:
        mail.Subject = f'Report for {dt}'
    if country == italy:
        mail.Subject = f'Report for {dt}'
    if country == gbr:
        mail.Subject = f'Report for {dt}'
    if country == deu_s:
        mail.Subject = f'Report for {dt}'
    if country == fra_s:
        mail.Subject = f'Report for {dt}'
    if country == ita_s:
        mail.Subject = f'Report for {dt}'
    if country == pol_t:
        mail.Subject = f'Report for {dt}'
    if country == cze_t:
        mail.Subject = f'Report for {dt}'


        
         
for country in countries:
    if (country == germany) or (country == romania) or (country == france) or (country == italy) or (country == gbr):
        employer_nm = "<> 'S'"
    elif (country == poland) or (country == czech):
        employer_nm = "<> 'T'"
    elif (country == deu_s) or (country == fra_s) or (country == ita_s):
        employer_nm = "= 'S'"
    elif (country == pol_t) or (country == cze_t):
        employer_nm = "= 'T'"
        
    query = f"""
    with
    pano as (
    SELECT x ,
    x ,
    x ,
    nullif(x,0) as x ,
    nullif(x,0) as x ,
    nullif((x - x),0) as x ,
    nullif(x,0) as x ,
    nullif(x,0) as x ,
    nullif((x / x),0) as x ,
    nullif(x,0) as x ,
    nullif((x / x),0) as x ,
    nullif((x / x),0) as x ,
    nullif(x,0) as x ,
    nullif(x,0) as x ,
    convert(varchar(5), (x/x) * 100) + ' %' as x ,
    nullif((x/x),0) as x ,
    nullif(x, 0) as x
    FROM x.x 
    where snapshot_date = {timeframe}
    and x  = 'A'
    and employer_nm {employer_nm}
    and supp_country IN ({country})
    and x = 'x'
    and x IN ('x', 'x', 'x Office')),
    contact_number as (
    select x ,
    count(x) as x
    from x.x
    where supp_cntry_cd IN ({country})
    and contact_dt = {timeframe}
    and x < '60'
    group by x )
    SELECT x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x ,
    x
    FROM pano
    LEFT JOIN contact_number
    ON panorama.x = contact_number.x
    ORDER BY x"""
    
    query_list = list()
    df1 = Query(query).run_store_append(cursor=cursor, query_list=query_list)
    cols_to_be_changed = ['x',
                          'x',
                          'x',
                          'x',
                          'x',
                          'x',
                          'x']
    for col in cols_to_be_changed:
        df1[col] = pd.to_numeric(df1[col])
        df1[col] = pd.to_datetime(df1[col], unit='s').dt.floor('S')
        df1[col] = df1[col].dt.time
        
        
        
        
        
    query = f"""
        WITH
        pano as (x),
        a as (x),
        b as (x),
        c as (x),
        d as (x),
        e as (x),
        f as (x),
        g as (x)
        SELECT x ,
        x ,
        x ,
        x ,
        x ,
        x ,
        x ,
        x ,
        x ,
        x ,
        x ,
        nullif((isnull(x, 0) + isnull(x, 0) + isnull(x, 0) + isnull(x, 0) + isnull(x, 0) + isnull(x, 0)),0) as x,
        x/nullif(x,0) as x
        FROM pano
        LEFT JOIN a
        ON pano.x = a.a
        LEFT JOIN b
        ON pano.x = b.s
        LEFT JOIN c
        ON pano.x = c.d
        LEFT JOIN d
        ON pano.x = d.f
        LEFT JOIN e
        ON pano.x = e.g
        LEFT JOIN f
        ON pano.x = f.h
        LEFT JOIN g
        ON pano.x = g.j
        ORDER BY x"""
        
    query_list = list()
        
    df2 = Query(query).run_store_append(cursor=cursor, query_list=query_list)
    cols_to_be_changed = ['x',
                              'x',
                              'x',
                              'x']

    for col in cols_to_be_changed:
            df2[col] = pd.to_numeric(df2[col])
            df2[col] = pd.to_datetime(df2[col], unit='s').dt.floor('S')
            df2[col] = df2[col].dt.time
            
            
    query = f"""with 
            pano as (x,
            convert(varchar(5), ((coalesce(x,0)+coalesce(x,0)))/nullif((coalesce(x,0)+coalesce(x,0)+coalesce(x,0)),0)*100) + ' %' as x
            FROM x.x ees
            where snapshot_date = {timeframe}
            and x  = 'A'
            and employer_nm {employer_nm}
            and supp_country IN ({country})
            and x = 'x'
            and x IN ('x', 'x', 'x Office')),
            u as (
            SELECT x as x,
            nullif(ROUND(NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            AS real_x,
            nullif(NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            NVL(SUM(x),0) +
            AS planned_x,
            convert(varchar(5), (real_x/planned_x) * 100) + ' %' as x
            FROM x.x
            WHERE snapshot_date = {timeframe}
            AND supp_country IN ({country})
            AND x ='x'
            GROUP BY x)
            SELECT
                x ,
                x ,
                x ,
                x,
                x,
                x,
                x,
                x,
                x,
                x,
                x,
                x,
                x,
                x,
                x,
                x,
                real_x,
                planned_x                
                FROM panorama
                    LEFT JOIN u
                            ON pano.x = u.x
                            ORDER BY x"""
        
    query_list = list()
        
    df3 = Query(query).run_store_append(cursor=cursor, query_list=query_list)
       
        
    cols_to_be_changed = ['x',
                              'x',
                              'x',
                              'x']
        
    for col in cols_to_be_changed:
            df3[col] = pd.to_numeric(df3[col])
            df3[col] = pd.to_datetime(df3[col], unit='s').dt.floor('S')
            df3[col] = df3[col].dt.time
            
    master_df = pd.concat([df1, df2, df3], axis=1)
    master_df.to_csv('master_source.csv', index=False, encoding='utf-8')
    
    d = date.today() - timedelta(days=1)
    dt = d.strftime("%d/%m/%y")

    save_location(country)

    for x in range(1, 4):
            os.system(f'start "excel" "C:\\xxxxx\\table_{x}.xlsm"')
            time.sleep(10)
            
    
    outlook = win32.gencache.EnsureDispatch('Outlook.Application') 
    mail = outlook.CreateItem(0)
               
    recipent(country)
    subject(country)
    
    table1 = open(r'C:\xxxxxxx\table_1.htm').read()
    table2 = open(r'C:\xxxxxxx\table_2.htm').read()
    table3 = open(r'C:\xxxxxxxxxx\table_3.htm').read()

    mail.HTMLBody = f"""\
        <html>
        <head></head>
        <body>
        
        Hello team,<br><br> 
        Below are metrics for previous day<br><br>
        
        <b>Call Metrics:</b><br><br>
        
        {table1}<br><br>
        
        Reference:<br>

        
        <b>Back office metrics:</b><br><br>
        
        {table2}<br><br>
        
        Reference:<br>

        
        <b>Agent statistics:</b><br><br>
        
        {table3}<br>
        
        Reference:<br>

        If you will have any questions please contact me<br><br>
        
        Kind regards,<br>
        Lukasz<br>
        </body>
        </html>
        """
    mail.Send()
