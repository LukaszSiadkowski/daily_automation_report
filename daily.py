import psycopg2
import pandas as pd
from datetime import datetime, timedelta
import numpy as np
import os
import time


con = psycopg2.connect(
    dbname='xxx',
    host='xxxxx',
    port='xxxx',
    user='xxx',
    password='xxxx')


class Query(str):

    
    def run(self, cursor):
        cursor.execute(self)
        return self

   
    def append(self, query_list):
        query_list.append(self)
        return self

    
    def run_append(self, cursor, query_list):
        cursor.execute(self)
        query_list.append(self)
        return self

    
    def run_store(self, cursor):
        time_init= time.time()
        cursor.execute(self)
        df = pd.DataFrame.from_records(
            iter(cursor), columns=[x[0] for x in cursor.description])
        print('execution time (minutes) =', round((time.time()-time_init)/60, 1))
        return df

    
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

cursor = cur = con.cursor()

query = """SELECT loc_cntry_cd ,
empl_login ,
reports_to_login ,
x ,
x ,
coalesce (x, 0) as x  ,
.....
FROM xx 
where snapshot_date = current_date -1
and x
and x
and x  
order by reports_to_login """

query_list = list()

df = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)  #run_store

df['x'] = df['x'].fillna(0)

df.insert(loc=7, column='x', value=df['x']-df['x'])
df.insert(loc=15, column='OCC', value=(df['x']+df['x'])/(df['x']+df['x']+df['x']))

df.replace(0, np.nan, inplace=True)
df.insert(loc=9, column='AHT', value=df['x']/df['x'])
df.insert(loc=16, column='SLA60', value=df['x']/df['x'])

cols_to_be_numeric = ['AHT', 'x', ........ ]
for col in cols_to_be_numeric:
     df[col] = pd.to_numeric(df[col])

df['delete'] = 0

cols_to_be_datetime = ['x', 'x', ......]
for col in cols_to_be_datetime:
    df[col] = pd.to_datetime(df[col], unit='s').dt.floor('S')

cols_to_be_time = ['x', 'x', ......]
for col in cols_to_be_time:
     df[col] = df[col].dt.time

df = df.replace((df.loc[0][20]), '')
df = df.replace(0, '')

df = df.drop(['x', 'x', ......] , axis=1)


cursor = cur = con.cursor()

query = """select xxxx ,
associate_login ,
case_count
from xxxxx
where (xxxx or xxxx)
and date_trunc('day', closed_ts) = current_date -1
and xxxx <> 'Phone' """

query_list = list()

pano = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list) 


df_pivot = pano.pivot_table(index=["x"], aggfunc='count', fill_value=0)
df_pivot = df_pivot.drop('x', axis=1)
df_pivot['x'] = df_pivot.index
df_pivot.columns = ['x', 'x']

first_join = pd.merge(df, df_pivot, on='x', how='left')

cursor = cur = con.cursor()


query = """select xxxxxxx ,
submitted_by ,
case_count
from xxxxxxx 
where date_trunc('day', create_date) = current_date -1
and (submitted_by=xxxxxxx)"""


query_list = list()

tickets_1 = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)

tickets_1.columns=['xxxxx', 'empl_login', 'case_count']


query = """select xxxxxx ,
resolved_by_login ,
case_count
from xxxxx 
where date_trunc('day', resolved_date) = current_date -1
and (resolved_by_login=xxxxxxx or .......)"""

query_list = list()

tickets_2 = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)

tickets_2.columns=['xxxxx', 'empl_login', 'case_count']

tickets_together = pd.concat([tickets_1, tickets_2])
tickets_together.drop_duplicates(subset =["xxxxxx", "empl_login"], keep = 'last', inplace = True)


tickets_together = tickets_together.drop('xxxxxxxx', axis=1)
df_pivot_2 = tickets_together.pivot_table(index=["empl_login"], aggfunc='count', fill_value=0)
df_pivot_2['empl_login'] = df_pivot_2.index
if df_pivot_2.empty == True:
    df_pivot_2.insert(0,'Tickets Resolved or created', 0)
else:
    df_pivot_2.columns=['Tickets Resolved or created', 'empl_login']
df_pivot_2.reset_index(drop = True, inplace = True)




second_join = pd.merge(first_join, df_pivot_2, on='empl_login', how='left')

second_join.insert(loc=16, column='empl_login ', value=df['empl_login'])
second_join.insert(loc=17, column='xxxx ', value=df['xxxxx'])
second_join['SLA60'] = pd.to_numeric(second_join['SLA60'], errors='coerce').fillna(0).map("{:.2%}".format)
second_join = second_join.replace('0.00%', '')
second_join['OCC'] = pd.to_numeric(second_join['OCC'], errors='coerce').fillna(0).map("{:.2%}".format)
second_join = second_join.replace('0.00%', '')


second_join.to_csv('main_team.csv', index=False, encoding='utf-8')



cursor = cur = con.cursor()


query = """SELECT loc_cntry_cd ,
empl_login ,
reports_to_login ,
local_log_in ,
coalesce (xxxx, 0) as xxxxx  ,
coalesce (xxxxx, 0) as xxxxxx ,
.....
FROM xxxxx.xxxxxx
where snapshot_date = current_date -1
and xxxxxx
and xxxxxxxxx
and (xxxxxxxx or xxxxxx or xxxxxxxxxxxx)  
order by reports_to_login """

query_list = list()

df = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)  #run_store

df['x'] = df['x'].fillna(0)

df.insert(loc=7, column='x', value=df['x']-df['x'])
df.insert(loc=15, column='OCC', value=(df['x']+df['x'])/(df['x']+df['x']+df['x']))

df.replace(0, np.nan, inplace=True)
df.insert(loc=9, column='x', value=df['x']/df['x'])
df.insert(loc=16, column='SLA60', value=df['x']/df['x'])

cols_to_be_numeric = ['xx', 'lunch', 'xx', ......]
for col in cols_to_be_numeric:
     df[col] = pd.to_numeric(df[col])

df['delete'] = 0

cols_to_be_datetime = ['AHT', 'x', 'x', ......]
for col in cols_to_be_datetime:
    df[col] = pd.to_datetime(df[col], unit='s').dt.floor('S')

cols_to_be_time = ['x', 'delete', ......]
for col in cols_to_be_time:
     df[col] = df[col].dt.time

df = df.replace((df.loc[0][18]), '')
df = df.replace(0, '')

df = df.drop(['x', 'x', 'delete', 'x', 'x'] , axis=1)


cursor = cur = con.cursor()

query = """select case_num ,
x ,
case_count
from xxxxxx
where xxxxxxx
and date_trunc('day', closed_ts) = current_date -1
and xxxxx <> 'Phone' """


query_list = list()

pano = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)


df_pivot = pano.pivot_table(index=["x"], aggfunc='count', fill_value=0)
df_pivot = df_pivot.drop('case_num', axis=1)
df_pivot['empl_login'] = df_pivot.index
df_pivot.columns = ['xxxxxxx', 'empl_login']

first_join = pd.merge(df, df_pivot, on='empl_login', how='left')


cursor = cur = con.cursor()

query = """select xxxxxxxx ,
submitted_by ,
case_count
from xxxxxxxx
where date_trunc('day', create_date) = current_date -1
and (submitted_by='xxxxxx' or xxxxxxxx)
order by xxxxxxx"""

query_list = list()

tickets_1 = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)

tickets_1.columns=['xxxxxxxx', 'empl_login', 'case_count']



query = """select xxxxxxxx ,
resolved_by_login ,
case_count
from xxxxxxx  
where date_trunc('day', resolved_date) = current_date -1
and (resolved_by_login='xxxxxxxx' or xxxxx)
order by xxxxx"""

query_list = list()

tickets_2 = Query(query).run_store_append(cursor=cursor,
                                       query_list=query_list)

tickets_2.columns=['xxxxxx', 'empl_login', 'case_count']



tickets_together = pd.concat([tickets_1, tickets_2])
tickets_together.drop_duplicates(subset =["xxxx", "empl_login"], keep = 'last', inplace = True)
tickets_together = tickets_together.drop('xxxxxxxx', axis=1)


df_pivot_2 = tickets_together.pivot_table(index=["empl_login"], aggfunc='count', fill_value=0)
df_pivot_2['empl_login'] = df_pivot_2.index

if df_pivot_2.empty == True:
    df_pivot_2.insert(0,'Tickets Resolved or created', 0)
else:
    df_pivot_2.columns=['Tickets Resolved or created', 'empl_login']

df_pivot_2.reset_index(drop = True, inplace = True)



second_join = pd.merge(first_join, df_pivot_2, on='empl_login', how='left')

second_join.insert(loc=16, column='empl_login ', value=df['empl_login'])
second_join.insert(loc=17, column='xxxxx ', value=df['xxxxxx'])
second_join['SLA60'] = pd.to_numeric(second_join['SLA60'], errors='coerce').fillna(0).map("{:.2%}".format)
second_join = second_join.replace('0.00%', '')
second_join['OCC'] = pd.to_numeric(second_join['OCC'], errors='coerce').fillna(0).map("{:.2%}".format)
second_join = second_join.replace('0.00%', '')
second_join.to_csv('external_team.csv', index=False, encoding='utf-8')


os.system('start "excel" "xxxxxxxx\\Daily\\market_1.xlsm"')

