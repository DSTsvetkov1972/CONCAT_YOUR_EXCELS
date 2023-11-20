import pandas as pd
import os
df3=pd.DataFrame()
df0 = pd.read_csv('0.csv')
df1 = pd.read_csv('1.csv')
df3=df3._append(df0, ignore_index = True)
df3=df3._append(df0, ignore_index = True)
print(df0)
print(df1)
print(df3)
