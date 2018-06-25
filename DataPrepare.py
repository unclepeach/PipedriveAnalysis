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
del df['Deal - Deal created']
