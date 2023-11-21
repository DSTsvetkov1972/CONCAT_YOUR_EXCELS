import pandas as pd
import numpy as np
import openpyxl
import os 
import csv
import time
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
def get_sheets(show_sheets_button,get_headers_button):
    """
    Функция получает список листов во всех экселевских книгах в папке Исходники
    и загружает его в файл .sheets.csv
    Затем функция открывает на рабочем столе файл .sheets.xlsm
    """
    if os.path.exists(os.path.join(os.getcwd(),'~$.sheets.xlsm')):
        messagebox.showerror(TITLE,'Закройте таблицу .sheets.xlms и повторите попытку!')
        return
    source_folder = os.walk('Исходники')
    sheets_list = []
    for i in source_folder:
        folder =i[0]
        files = i[2]
        for file in files:
            if '~' in file:
                continue
            elif file[-5:] in ['.xlsx','.xlsm']:
                print(Fore.WHITE + 'Папка: %s Книга: %s:'%(folder,file))
                xl_path = os.path.join(folder,file)
                try:
                    #print("a")
                    wb = openpyxl.load_workbook(os.path.join(folder,file))
                    sheets = wb.worksheets
                    #print("b")
                except:
                    print(Fore.RED + "Не удалось получить список листов для:\n" + folder,file + Fore.WHITE)
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
                              'На листе: %s строк: %s колонок: %s'%(sheet.title,sheet.max_row,sheet.max_column) + Fore.WHITE)
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
    try:
        wb_get_sheets.RefreshAll()   
    except:
        fileName = os.path.join(os.getcwd(),'.sheets.xlsm')
        xl_get_sheets = win32com.client.DispatchEx("Excel.Application")
        wb_get_sheets = xl_get_sheets.Workbooks.Open(fileName)
        xl_get_sheets.Visible = True
        wb_get_sheets.RefreshAll()
        wb_get_sheets.Save()
        #wb_get_sheets.SaveAs(Filename=os.path.join(os.getcwd(),'.sheets.xlsm'))

    if sheets_list != []:
        show_sheets_button.pack(anchor = CENTER, pady=(25,0))
        get_headers_button.pack(anchor = CENTER, pady = (25,0), padx=(0,0))

"""
if __name__ == '__main__':   
    get_sheets()
"""
@start_finish_time
@proceed_type('"Открываем файл с листами в книгах"')
def show_sheets():
    #print('dfg')
    if os.path.exists(os.path.join(os.getcwd(),'~$.sheets.xlsm')):
        messagebox.showerror(TITLE,'Закройте таблицу .sheets.xlsm и повторите попытку!')
        return
    else:
        fileName = os.path.join(os.getcwd(),'.sheets.xlsm')
        xl_show_sheets = win32com.client.DispatchEx("Excel.Application")
        xl_show_sheets.Workbooks.Open(fileName)
        xl_show_sheets.Visible = True   


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
def get_tables_from_sheets(tables_from_sheets_dict, sheets_for_processing_list):
    #print('tables_from_sheets_dict_ДО\n',tables_from_sheets_dict.keys())

    # если поменялись отобранные листы, то проверяем какие таблицы нужно загрузить
    sheets_for_processing_df = pd.read_excel(os.path.join(os.getcwd(),'.sheets.xlsm'), header = 0)
    sheets_for_processing_df = sheets_for_processing_df[sheets_for_processing_df['Добавить'] == 'ДА']
    sheets_for_processing_list.clear()
    for row in sheets_for_processing_df.itertuples():
        sheets_for_processing_list.append(row[1:4]+row[7:])
   
    if os.path.exists(os.path.join(os.getcwd(),'.selected_sheets.csv')):
        sheets_for_processing_df_before = pd.read_csv(os.path.join(os.getcwd(),'.selected_sheets.csv'),sep='\t')
    else:
        sheets_for_processing_df.to_csv(os.path.join(os.getcwd(),'.selected_sheets.csv'),sep='\t',index = False)
        sheets_for_processing_df_before = sheets_for_processing_df
    #print('\nsheets_for_processing_df_before\n',sheets_for_processing_df_before)
    #print('\nsheets_for_processing_df\n',       sheets_for_processing_df)
    #print('\nsheets_for_processing_df_before.equals(sheets_for_processing_df)',sheets_for_processing_df_before.equals(sheets_for_processing_df))
    #print('\ntables_from_sheets_dict != ()\n',tables_from_sheets_dict)
    if sheets_for_processing_df_before.equals(sheets_for_processing_df) and tables_from_sheets_dict != {}: 
        print(Fore.GREEN + 'Список листов, которые нужно загрузить не менялся, таблицы уже загружены!' + Fore.WHITE)
    else:        
        for row in sheets_for_processing_df.iterrows():
            #try:
                folder        = row[1]['Папка']
                book          = row[1]['Книга']
                sheet         = row[1]['Лист']
                rows_limit    = int(row[1]['Сколько строк нужно'])
                columns_limit = int(row[1]['Сколько колонок нужно'])
                print(Fore.WHITE + 'Книга: %s Папка: %s Лист: %s'%(folder,book,sheet), end =' ')
                if  (folder,book,sheet,rows_limit,columns_limit) not in tables_from_sheets_dict:
                    try:
                        table = pd.read_excel(os.path.join(folder,book), sheet_name = sheet, header = None,  nrows = rows_limit, usecols = range(columns_limit))#.iloc[:,:columns_limit]
                    except:
                        table = pd.read_excel(os.path.join(folder,book), sheet_name = sheet, header = None,  nrows = rows_limit)#.iloc[:,:columns_limit]    
                    #print("b")
                    table.insert(0, 'Папка' , folder)
                    table.insert(1, 'Книга' , book) 
                    table.insert(2, 'Лист'  , sheet)
                    table.insert(3, 'Строка в исходнике', table.index.tolist())
                    tables_from_sheets_dict[(folder,book,sheet,rows_limit,columns_limit)]={'Таблица': table,'Превью': table.iloc[:30,:30]}
                    print(Fore.GREEN + ' - загрузили из экселя')
                else:
                    print(Fore.GREEN + ' - таблица с листа уже была загружена')
                
            #except:
            #    print(Fore.RED + 'Не удалось получить информацию для:\n' +folder,book,sheet + Fore.WHITE)

        sheets_for_processing_df.to_csv(os.path.join(os.getcwd(),'.selected_sheets.csv'),sep='\t',index= False)
   # print('tables_from_sheets_dict_ПОСЛЕ\n',tables_from_sheets_dict.keys())


#t = get_tables_from_sheets()
#print(t)



#@start_finish_time
#@proceed_type('"Создаём список исключений"')
def get_exceptions():
    exceptions_df = pd.read_excel('.headers.xlsx', sheet_name='Exceptions')
    return [list(i[1:]) for i in exceptions_df.itertuples()]

@start_finish_time
@proceed_type('"Получение заголовков"')
def get_headers(tables_from_sheets_dict, sheets_for_processing_list):

    exceptions_list = get_exceptions()
    get_tables_from_sheets(tables_from_sheets_dict,sheets_for_processing_list)
    headers_specifications_df = pd.read_excel('.headers.xlsx', sheet_name = 'Settings')  #print(headers_specifications_df)  

    all_tables_headers_df = pd.DataFrame()
    table_with_not_located_headers = pd.DataFrame()
    
    for sheet_for_processing_list in sheets_for_processing_list:
        #print(headers_specifications_df)
        print(Fore.WHITE + "Папка: {} Книга: {} Лист: {}".format(*sheet_for_processing_list), end = ' ')
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
                header_df = tables_from_sheets_dict[(sheet_for_processing_list)]['Превью'][tables_from_sheets_dict[(sheet_for_processing_list)]['Превью'][column_number] == sign ] 
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
                    tables_from_sheets_dict[(sheet_for_processing_list)]['Строка заголовка']  = table_headers_df.index[0]
                    #print(table_from_sheets_list['Строка заголовка'] )
                else:
                    print(Fore.RED + ' - есть заголовок из списка исключений' + Fore.WHITE, end = ' ')
                
                #input()
            #print('table_headers_df\n',table_headers_df)
            #print('$'*50)    
        if len(table_headers_df) == 0:
            table_headers_df = pd.DataFrame([{'Папка':sheet_for_processing_list[0],
                                              'Книга':sheet_for_processing_list[1],
                                              'Лист':sheet_for_processing_list[2],
                                              'Строка':'',
                                              'Найден по колонке':'',
                                              'По признаку':''
                                              }])
            table_with_not_located_headers = tables_from_sheets_dict[(sheet_for_processing_list)]['Превью']
            tables_from_sheets_dict[(sheet_for_processing_list)]['Строка заголовка']  = None  
            print(Fore.RED + ' - заголовки не найдены')
        else:
            print(Fore.GREEN + ' - заголовки найдены')
        
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
    #print(tables_from_sheets_dict)




#@start_finish_time
#@proceed_type('"Проверка на наличие случаев, когда для одной таблицы найдено несколько заголовков"')
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

#if __name__ == '__main__':
#    print(check_multiple_headers())

#@start_finish_time
#@proceed_type('"Проверка на наличие случаев, когда для таблицы не найдено ни одного заголовка"')
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

def check_identical_column_names(list_to_check):
    list_to_check = list_to_check
    identical_column_names_list = []
    while list_to_check:
        i = list_to_check[0]
        list_to_check = list_to_check[1:]
        if i in list_to_check and i not in identical_column_names_list:
            identical_column_names_list.append(i)
    return identical_column_names_list

    
#@start_finish_time
#@proceed_type('"Отобразить .headers.xlsx"')
def open_headers_xls(tables_from_sheets_dict,sheets_for_processing_list,concat_tables_button):
    """
    Функция открывает файл .headers.xlsx
    """
    if os.path.exists(os.path.join(os.getcwd(),'~$.headers.xlsx')):
        messagebox.showerror(TITLE,'Закройте таблицу .headers.xlsx и повторите попытку!')
        return
    concat_tables_button.pack_forget()
    get_headers(tables_from_sheets_dict, sheets_for_processing_list)

    try:
        wb.RefreshAll()
    except :
        fileName = os.path.join(os.getcwd(),'.headers.xlsx')
        xl = win32com.client.DispatchEx("Excel.Application")
        wb = xl.Workbooks.Open(fileName)
        xl.Visible = True
        wb.RefreshAll()

    if check_multiple_headers() and check_no_header():
        concat_tables_button.pack(anchor = CENTER, pady = (25,0), padx=(0,0))
        

  
#wb = open_headers_xls()
"""
while True:
    input("Пробуем")
    open_headers_xls()
"""


@start_finish_time
@proceed_type('"Объединение таблиц"')
def concat_tables(tables_from_sheets_dict,sheets_for_processing_list,concat_tables_button):
    sheets_for_processing_df = pd.read_excel(os.path.join(os.getcwd(),'.sheets.xlsm'), header = 0)
    sheets_for_processing_df = sheets_for_processing_df[sheets_for_processing_df['Добавить'] == 'ДА']
    sheets_for_processing_list_actual = []
    sheets_for_processing_list_cant_add = []

    for row in sheets_for_processing_df.itertuples():
        print(row)
        sheets_for_processing_list_actual.append(row[1:4]+row[7:])
    if  sorted(sheets_for_processing_list_actual) != sorted(sheets_for_processing_list):
        concat_tables_button.pack_forget()
        messagebox.showwarning(TITLE, "Изменился список листов таблицы с которых нужно объеденить.\nТребуется пересобрать заголовки!")   
        return
    
    if not check_no_header():
        return
    if not check_multiple_headers():
       return
    
    total_table_df = pd.DataFrame()

    for sheet_for_processing_list in sheets_for_processing_list:
            print(Fore.WHITE + 'Папка: {} Книга: {} Лист: {}'.format(*sheet_for_processing_list), end = ' ')
        #try: 
            header_row = tables_from_sheets_dict[(sheet_for_processing_list)]['Строка заголовка']
            header_list = list(tables_from_sheets_dict[(sheet_for_processing_list)]['Таблица'].iloc[header_row])[4:]
            column_names_raw = ['Папка','Книга','Лист','Строка в исходнике'] + header_list
            i=0
            column_names = []
            for column_name_raw in column_names_raw:
                if pd.isna(column_name_raw):
                    column_name = '_' + str(i)
                    i += 1
                else:
                    column_name = column_name_raw
                column_names.append(column_name)
            identical_column_names = ', '.join(list(map(str,check_identical_column_names(column_names))))
            if identical_column_names:
                sheet_for_processing_list_cant_add = list(sheet_for_processing_list[:3])
                sheet_for_processing_list_cant_add.append(identical_column_names)
                sheets_for_processing_list_cant_add.append(('Папка: {} Книга: {} Лист: {} Колонки {} встречаются более одного раза!'.format(*sheet_for_processing_list_cant_add)))
                print(Fore.RED + ' - таблица не может быть обработана так как названия колонок ',
                      Fore.WHITE + identical_column_names,
                      Fore.RED + ' встречаются более одного раза!' + Fore.WHITE)
                continue
            result_table = tables_from_sheets_dict[(sheet_for_processing_list)]['Таблица'][header_row+1:]
            result_table.columns = column_names
            result_table = result_table.dropna(how='all', axis = 1)
            total_table_df = total_table_df._append(result_table, ignore_index = True)
            print(Fore.GREEN + ' - данные извлечены и добавлены в итоговую таблицу' + Fore.WHITE)
        #except:
        #    sheets_for_processing_list_cant_add.append(sheets_for_processing_list_cant_add)
        #    print(Fore.RED + 'НЕ УДАЛОСЬ ИЗВЛЕЧЬ ДАННЫЕ ПО ПРИЧИНЕ РАНЕЕ НЕ ВСТРЕЧАВШЕЙСЯ! ОБРАТИТЕСЬ К РАЗРАБОТЧИКАМ!')
        #    continue
        #finally:
        #    print(total_table_df)
            
            #print(Fore.CYAN + 'total_table_df\n' + Fore.CYAN,total_table_df)
      
    total_table_df = total_table_df.replace('',np.nan)
    total_table_df_columns = total_table_df.columns[4:]
    total_table_df = total_table_df.dropna(axis=0,subset=total_table_df_columns,how='all')
    #print(total_table_df)
    total_table_df = total_table_df.dropna(axis=1, how = 'all')
    total_table_df = total_table_df.fillna('')
    total_table_df = total_table_df.map(lambda x: str(x).replace(r'\n',r'\\n').replace(chr(10),''))
    #print(total_table_df)
    print(Fore.YELLOW + '', datetime.now(),'\t записываем результат в RESULT.csv' + Fore.WHITE)
 
    total_table_df.to_csv('RESULT.csv', sep ='\t')
    print(Fore.YELLOW + '', datetime.now(),'\t результат записан в RESULT.csv' + Fore.WHITE)

    if len(sheets_for_processing_list_cant_add) > 0:
        print(Fore.RED + 'ВНИМАНИЕ: НЕКОТОРЫЕ ТАБЛИЦЫ НЕ УДАЛОСЬ ОБРАБОТАТЬ!' + Fore.WHITE)
        for sheet_for_processing_list_cant_add in sheets_for_processing_list_cant_add:
            print(Fore.RED + sheet_for_processing_list_cant_add + Fore.WHITE)
            #print(Fore.RED + 'Папка: {} Книга: {} Лист: {}'.format(*sheet_for_processing_list_cant_add) + Fore.WHITE)
        messagebox.showwarning(TITLE, "Таблицы объеденены,\nНО НЕ ВСЕ!\nРезультат записан в RESULT.csv'")   
    else:
        messagebox.showinfo(TITLE, "Таблицы объеденены. Результат записан в RESULT.csv'")


        
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