import pandas as pd
import numpy as np
import openpyxl
import os 
import csv
import win32com.client
from tkinter import *
from tkinter import messagebox, ttk
import tkinter as tk
from datetime import datetime
from colorama import init, Fore, Back, Style

import warnings
warnings.filterwarnings('ignore')

VERSION = 2
TITLE = "CONCAT_YOUR_EXCELS v/%s"%(VERSION)

def proceed_type(proceed_type):
    def wrapper(func):
        func.proceed_type = proceed_type
        return func
    return wrapper

def start_finish_time(func):
    def wrapper(*args, **kwargs):
        start = datetime.now()
        print(Fore.YELLOW + '\n', start,'\tОбработка %s стартовала...'%(func.proceed_type) + Fore.WHITE)
        value = func(*args, **kwargs)
        finish = datetime.now()
        print(Fore.YELLOW, finish,'\tОбработка %s завершена!\n'%(func.proceed_type) + Fore.WHITE)
        return value
    return wrapper


@start_finish_time
@proceed_type('"Сканирование листов в книгах"')
def get_sheets():
    """
    Функция получает список листов во всех экселевских книгах в папке Исходники
    и загружает его в файл .sheets.csv
    Затем функция открывает на рабочем столе файл .sheets.xlsm
    """
    source_folder = os.walk('Исходники')
    sheets_list = []
    for i in source_folder:
        folder =i[0]
        files = i[2]
        for file in files:
            if '~' in file:
                continue
            elif file[-5:] in ['.xlsx','.xlsm']:
                print(Fore.WHITE + folder,file)
                xl_path = os.path.join(folder,file)
                try:
                    #print("a")
                    wb = openpyxl.load_workbook(os.path.join(folder,file))
                    sheets = wb.worksheets
                    #print("b")
                except:
                    print(Fore.RED + "Не удалось получить список листов дляЖ\n" + folder,file + Fore.WHITE)
                    continue
                for sheet in sheets:
                    try:
                        #print("c")
                        sheets_list.append([folder,
                                            file,
                                            sheet.title,
                                            sheet.max_row,
                                            sheet.max_column,
                                            'ДА',
                                            sheet.max_row,
                                            sheet.max_column])
                        print(Fore.GREEN + 
                              'В папке: %s в файле: %s на листе: %s строк: %s колонок: %s'%(folder,file, sheet.title,sheet.max_row,sheet.max_column) + 
                              Fore.WHITE)
                    except:
                        print(Fore.RED + "Не удалось получить информацию о листе для\n" + folder,file,sheet + Fore.WHITE)
                        continue
 
    with open('.sheets.csv', 'w', newline='', encoding='utf-8') as sheets_csv:
        writer = csv.writer(sheets_csv, delimiter='\t')
        writer.writerow(['Папка',
                         'Книга',
                         'Лист',
                         'Строк на листе',
                         'Столбцов на листе',
                         'Добавить',
                         'Сколько строк нужно',
                         'Сколько колонок нужно'
                         ])                

    with open('.sheets.csv', 'a', newline='', encoding='utf-8') as sheets_csv:
        writer = csv.writer(sheets_csv, delimiter='\t')
        writer.writerows(sheets_list)

    #os.system('start excel.exe %s'%('.sheets.xlsm'))
    
    fileName = os.path.join(os.getcwd(),'.sheets.xlsm')
    xl = win32com.client.DispatchEx("Excel.Application")
    wb = xl.Workbooks.Open(fileName)
    xl.Visible = True
    wb.RefreshAll()    
"""
if __name__ == '__main__':   
    get_sheets()
"""
def show_sheets():
    fileName = os.path.join(os.getcwd(),'.sheets.xlsm')
    xl = win32com.client.DispatchEx("Excel.Application")
    wb = xl.Workbooks.Open(fileName)
    xl.Visible = True




@start_finish_time
@proceed_type('"Сравниваем листы в файле .sheets.csv и .sheets.xlmm"')
def compare_sheets ():
    df_csv  = pd.read_csv('.sheets.csv', sep = '\t').iloc[:,:5]
    df_xlms = pd.read_excel('.sheets.xlsm', sheet_name='Sheets_in_book', header = 1).iloc[:,:5]
    #print(df_csv)
    #print(df_xlms)
    return df_csv.equals(df_xlms)
#print(compare_sheets ())




#print(fileName)

##xl = win32com.client.DispatchEx("Excel.Application")
#wb = xl.Workbooks.Open(fileName)
#xl.Visible = True

#os.system('start excel.exe .sheets.xlsm')
    
"""
wb.RefreshAll()
wb.Close(True)
xl.Quit()
os.system("taskkill /f /im excel.exe")
"""

#open_sheets_and_settings()  


@start_finish_time
@proceed_type('"Создание списка таблиц"')
def get_tables_from_sheets():
    
    tables_from_sheets_list = []
    sheets_for_processing_df = pd.read_excel('.sheets.xlsm', header = 0)
    sheets_for_processing_df = sheets_for_processing_df[sheets_for_processing_df['Добавить'] == 'ДА']
    for row in sheets_for_processing_df.iterrows():
        try:
            folder        = row[1]['Папка']
            book          = row[1]['Книга']
            sheet         = row[1]['Лист']
            rows_limit    = int(row[1]['Сколько строк нужно'])
            columns_limit = int(row[1]['Сколько колонок нужно'])
            #print("a")
            table         = pd.read_excel(os.path.join(folder,book), sheet_name = sheet, header = None,  nrows = rows_limit, usecols = range(columns_limit))#.iloc[:,:columns_limit]
            #print("b")
            table.insert(0, 'Папка' , folder)
            table.insert(1, 'Книга' , book) 
            table.insert(2, 'Лист'  , sheet)
            table.insert(3, 'Строка в исходнике', table.index.tolist())
            tables_from_sheets_list.append({'Папка'                : folder,
                                            'Книга'                : book,
                                            'Лист'                 : sheet,
                                            'Сколько строк нужно'  : rows_limit,
                                            'Сколько колонок нужно': columns_limit,
                                            'Таблица'              : table
                                            }) 
            print(Fore.GREEN + folder,book,sheet + Fore.WHITE)
        except:
            print(Fore.RED + 'Не удалось получить информацию для:\n' +folder,book,sheet + Fore.WHITE)
    return tables_from_sheets_list
#t = get_tables_from_sheets()
#print(t)



@start_finish_time
@proceed_type('"Создание списка таблиц с ненайдеными заголовками"')
def get_exceptions():
    exceptions_df = pd.read_excel('.headers.xlsx', sheet_name='Exceptions')
    return [list(i[1:]) for i in exceptions_df.itertuples()]

@start_finish_time
@proceed_type('"Получение заголовков"')
def get_headers(tables_from_sheets_list):
    exceptions_list = get_exceptions()
    headers_specifications_df = pd.read_excel('.headers.xlsx', sheet_name = 'Settings')  #print(headers_specifications_df)  

    all_tables_headers_df = pd.DataFrame()
    table_with_not_located_headers = pd.DataFrame()
    
    for table_from_sheets_list in tables_from_sheets_list:
        #print(headers_specifications_df)
        print(Fore.GREEN + 'Папка: {Папка} Книга: {Книга} Лист: {Лист}'.format(**table_from_sheets_list) + Fore.WHITE)
        #print("table_dict['Таблица']\n",table_dict['Таблица'])
        #print('*'*50)
        #break
        #input('aaaa')
        table_headers_df = pd.DataFrame()
        
        for specification in headers_specifications_df.iterrows():
            column_number = specification[1]['Колонка']
            sign          = specification[1]['Признак']
            #print('Колонка: %s Признак: %s'%(column_number,sign))  
            #print('-'*50) 

            # отрабатываем то, что в таблице может быть меньше колонок чем в спецификации на поиск заголовка
            try:          
                header_df = table_from_sheets_list['Таблица'][table_from_sheets_list['Таблица'][column_number] == sign ] 
            except KeyError:
                header_df = pd.DataFrame() 

            #print(header_df)
            #print('-'*50)
            if len(header_df) > 0:
                #print('header_df\n', header_df ) 
                header_list = list(header_df.iloc[0])[:4]
                header_list.append(column_number)
                header_list.append(sign)
                #print('header_list \n', header_list ) 
                header_df.insert(4, 'Найден по колонке' , column_number)
                header_df.insert(5, 'По признаку' , sign)
                header_df.fillna('')
                #table_from_sheets_list['Строка заголовка']  = (header_df['Строка'].iloc[0])
                if header_list not in exceptions_list:
                    table_headers_df = table_headers_df._append(header_df)
                    #print(table_headers_df)
                    table_from_sheets_list['Строка заголовка']  = table_headers_df.index[0]
                    #print(table_from_sheets_list['Строка заголовка'] )
                else:
                    print('Найденный заголовок в списке исключений!')

                #input()
            #print('table_headers_df\n',table_headers_df)
            #print('$'*50)    
        if len(table_headers_df) == 0:
            table_headers_df = pd.DataFrame([{'Папка':table_from_sheets_list['Папка'],
                                              'Книга':table_from_sheets_list['Книга'],
                                              'Лист':table_from_sheets_list['Лист'],
                                              'Строка':'',
                                              'Найден по колонке':'',
                                              'По признаку':''
                                              }])
            table_with_not_located_headers = table_from_sheets_list['Таблица']
            table_from_sheets_list['Строка заголовка']  = None  
        
        table_with_not_located_headers.to_csv('.table_with_not_located_headers.csv', sep = '\t', index= False)         
        all_tables_headers_df = all_tables_headers_df._append(table_headers_df, ignore_index = True)
        
        #print('aaaaaaaaa', int(column_number),sep='\n')                
        #print('bbbbbbbbb',header_df,sep='\n')
        #print('ccccccccc',all_tables_headers_df.fillna(''),sep='\n')
 
        #print('-'*50,'\n'*2)
    #result = True if len(table_with_not_located_headers) == 0 else False
    #print(all_tables_headers_df)
    all_tables_headers_df.to_csv('.headers.csv', sep = '\t', index= False)
    #concat_button = ttk.Button(root, text ="Объединить таблицы", width = 30, command = concat_tables)
    #concat_button.pack(anchor = CENTER, pady = (25,0))
    return tables_from_sheets_list


def check_multiple_headers():
    try:
        headers_df = pd.read_csv('.headers.csv', sep ='\t')
    except pd.errors.EmptyDataError:
        print(Fore.RED + 'Обработка невозможна! Вероятно отсутствуют книги или листы для которых требуется собрать информацию?' + Fore.WHITE)
        return
    headers_df_gropby = headers_df.groupby(['Папка','Книга','Лист']).count().reset_index()
    headers_df_gropby = headers_df_gropby.rename(columns = {'Найден по колонке':'Повторов'})
    headers_df_gropby = headers_df_gropby[headers_df_gropby['Повторов'] > 1]
    headers_df_gropby = headers_df_gropby[['Папка','Книга','Лист']]
    multiple_headers_df = headers_df_gropby.merge(headers_df,
                                                   how='inner',
                                                   on = ['Папка','Книга','Лист'] )
    multiple_headers_df.to_csv('.multiple_headers.csv', sep = '\t', index=None)
    if len(multiple_headers_df) == 0:
        print(Fore.GREEN + 'Случаи когда в одной таблице обнаружены несколько заголовков отсутствуют!' + Fore.WHITE)
        return True
    else:
        print(Fore.RED + 'Есть случаи когда в одной таблице обнаружены несколько заголовков!\nДобавьте неправильные закголвки в таблицу на листе Exceptions!' + Fore.WHITE)
        return False
if __name__ == '__main__':
    print(check_multiple_headers())

def check_no_header(): 
    """
    Возвращает истину если нет таблиц с необнаружеными заголовками
    """   
    try: 
        pd.read_csv('.table_with_not_located_headers.csv', sep = '\t')
        print(Fore.RED + 'Есть таблицы для которых не найдены заголовки!' + Fore.WHITE)
        return False
    except pd.errors.EmptyDataError:
        print(Fore.GREEN + 'Заголовки найдены для всех таблиц!' + Fore.WHITE)        
        return True



    
@start_finish_time
@proceed_type('"Отобразить .headers.xlsx"')
def open_headers_xls(headers_button):
    """
    Функция открывает файл .headers.xlsx
    """
    global wb, tables_from_sheets_list

    headers_button.pack_forget()

    tables_from_sheets_list = get_tables_from_sheets()
    tables_from_sheets_list = get_headers(tables_from_sheets_list)

    try:
        wb.RefreshAll()
    except :
        fileName = os.path.join(os.getcwd(),'.headers.xlsx')
        xl = win32com.client.DispatchEx("Excel.Application")
        wb = xl.Workbooks.Open(fileName)
        xl.Visible = True
        wb.RefreshAll()

    if check_multiple_headers() and check_no_header():
        headers_button.pack(anchor = CENTER, pady = (25,0), padx=(0,0))
        

    return wb, tables_from_sheets_list 
  
#wb = open_headers_xls()
"""
while True:
    input("Пробуем")
    open_headers_xls()
"""


@start_finish_time
@proceed_type('"Объединение таблиц"')
def concat_tables():
    if not check_no_header():
        return
    if not check_multiple_headers():
       return
    global tables_from_sheets_list
    total_table_df = pd.DataFrame()

    for table_dict in tables_from_sheets_list:
        #print(table_dict)
        try: 
            #print('table_dict\n',table_dict )
            header_row = table_dict['Строка заголовка']
            #print('header_row\n',header_row)
            header_list = list(table_dict['Таблица'].iloc[header_row])[4:]
            #print('header_list\n',header_list)
            column_names = ['Папка','Книга','Лист','Строка в исходнике'] + header_list
            #print('column_names\n',column_names)
            #print('Таблица\n',table_dict['Таблица'])
            result_table = table_dict['Таблица'][header_row+1:]
            result_table.columns = column_names
            #print('result\n',result_table)
            #print('-'*100,'\n')
            print(Fore.GREEN + 'Папка: {Папка} Книга: {Книга} Лист: {Лист} - данные извлечены успешно!'.format(**table_dict) + Fore.WHITE)
        except:
            print(Fore.RED + 'Папка: {Папка} Книга: {Книга} Лист: {Лист} - НЕ УДАЛОСЬ ИЗВЛЕЧЬ ДАННЫЕ!'.format(**table_dict) + Fore.WHITE)
            continue
        total_table_df = total_table_df._append(result_table, ignore_index = True)
    #total_table_df = total_table_df.replace('',np.nan)
    total_table_df_columns = total_table_df.columns[4:]
    #print('total_table_df_columns\n',total_table_df_columns)
    total_table_df = total_table_df.dropna(axis=0,subset=total_table_df_columns,how='all')
    #print(total_table_df)
    total_table_df = total_table_df.dropna(axis=1, how = 'all')
    #print(total_table_df)
    total_table_df.to_csv('RESULT.csv', sep ='\t')



        
    #total_table_df._append(table_df, ignore_index = True)
    #total_table_df.to_csv('RESULT.csv', encoding = 'utf-8', sep='\t')   
"""
if __name__ == '__main__':
    t = get_tables_from_sheets()
    tables_from_sheets_list = get_headers(t)
    #print(r)
    concat_tables()
"""
#break
#tabs = get_tables_from_sheets()
#res = get_headers(tabs)
#print('*'*100)
#print('*'*100)
#print('*'*100)
#concat_tables(res)