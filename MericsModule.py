### preparation
import numpy as np
import matplotlib.pyplot as plt
import pandas as pd
from pandas import DataFrame, Series
from datetime import date

workbook=pd.ExcelFile('C:\\Users\\becintern\\Downloads\\deals-2644782-303.xlsx')
dictionary = {}
for sheet_name in workbook.sheet_names:
 df = workbook.parse(sheet_name)
 dictionary[sheet_name] = df

df['Deal - Deal created'] = pd.to_datetime(df['Deal - Deal created'])
df.index = df['Deal - Deal created'] 
###del df['Deal - Deal created']

### leads trends plot
df['Deal - Deal created'].resample('W').count()
weekly_leads = df['Deal - Deal created'].resample('W').count()
weekly_leads.plot()
df['Deal - Deal created'].resample('M').count()
monthly_leads = df['Deal - Deal created'].resample('M').count()
monthly_leads.plot()

### period setting and use
#begin = pd.Timestamp('2018-06-18 07:00:00')
#end = pd.Timestamp('2018-06-24 23:00:00')
#df[(df.index >= begin)&(df.index <=end)]

###1.Leads last week/month
leads_w = df['Deal - Deal created'].resample('W').count().tail(1).get(0)
print ('Leads_last week =', leads_w)
leads_m = df['Deal - Deal created'].resample('M').count().tail(1).get(0)
print ('Leads_last month =', leads_m)

###2.Conversion ratio last week/month
dfa = df.groupby(['Deal - Status']).get_group('Won')
won_w = dfa['Deal - Deal created'].resample('W').count().tail(1)
won_w = won_w.get(0)
all_w = df['Deal - Deal created'].resample('W').count().tail(1)
all_w = all_w.get(0)
print ('Conversion ratio_last week = '+'{:.1%}'.format(won_w/all_w))

dfa = df.groupby(['Deal - Status']).get_group('Won')
won_m = dfa['Deal - Deal created'].resample('M').count().tail(1)
won_m = won_m.get(0)
all_m = df['Deal - Deal created'].resample('M').count().tail(1)
all_m = all_m.get(0)
print ('Conversion ratio_last month = '+'{:.1%}'.format(won_m/all_m))

###3.Revenue:YTD, last week/month, quarter
dfa = df.groupby(['Deal - Status']).get_group('Won')
rev_w = dfa['Deal - Value'].resample('W').sum().tail(1)
rev_w = rev_w.get(0)
print ('Rev_last week =', rev_w, 'USD')

dfa = df.groupby(['Deal - Status']).get_group('Won')
rev_m = dfa['Deal - Value'].resample('M').sum().tail(1)
rev_m = rev_m.get(0)
print ('Rev_last month =', rev_m, 'USD')

dfa = df.groupby(['Deal - Status']).get_group('Won')
rev_q = dfa['Deal - Value'].resample('Q').sum().tail(1)
rev_q = rev_q.get(0)
print ('Rev_last quarter =', rev_q, 'USD')

dfa = df.groupby(['Deal - Status']).get_group('Won')
### Because the latest pipedrive raw data only contains 2018 period
rev_YTD = dfa['Deal - Value'].sum()
print ('Rev_YTD =', rev_YTD.round(2), 'USD')

### Pattern Detection based on 2013-2018 leads ###
workbook=pd.ExcelFile('C:\\Users\\becintern\\Downloads\\deals-2644782-306.xlsx')
dictionary = {}
for sheet_name in workbook.sheet_names:
 df = workbook.parse(sheet_name)
 dictionary[sheet_name] = df

df['Deal - Deal created'] = pd.to_datetime(df['Deal - Deal created'])
df.index = df['Deal - Deal created'] 
