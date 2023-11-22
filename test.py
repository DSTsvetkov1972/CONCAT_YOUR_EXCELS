import pandas as pd
import os

df = pd.read_excel(os.path.join('Исходники','Test.xlsx'), sheet_name="Лист2")
print(df)
df = df._append(df)
df.to_csv('aaa.csv', encoding = 'utf-8')
df = pd.read_csv('aaa.csv')
#print(len(df))


print(len(open('RESULT.csv', encoding='utf-8').readlines()))