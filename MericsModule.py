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

###1.Leads last week/month
df['Deal - Deal created'].resample('W').count()
weekly_leads = df['Deal - Deal created'].resample('W').count()
weekly_leads.plot()
df['Deal - Deal created'].resample('M').count()
monthly_leads = df['Deal - Deal created'].resample('M').count()
monthly_leads.plot()

### period setting and use
###begin = pd.Timestamp('2018-06-18 07:00:00')
###end = pd.Timestamp('2018-06-24 23:00:00')
###df[(df.index >= begin)&(df.index <=end)]
