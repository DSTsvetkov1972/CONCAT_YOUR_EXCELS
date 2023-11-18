import pandas
import os
df = pandas.read_excel(os.path.join(os.getcwd(),'Исходники','Книга_1.xlsx'),header =None, sheet_name="Лист2",nrows=100)
print(df)
