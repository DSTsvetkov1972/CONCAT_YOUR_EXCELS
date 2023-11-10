import pandas as pd
import datetime
import os
import warnings
from tkinter import messagebox
from colorama import init, Fore, Back, Style

VERSION = 2
message_title = 'CONCAT_YOUR_EXCELS v.%s'%(VERSION)

warnings.filterwarnings("ignore")

#import colorama


init()
print(Style.BRIGHT)

files_to_proceed_df = pd.read_excel('_CONCAT_YOUR_EXCELS (предобработка).xlsm', 'Titles_in_book', header = 1)

print(Fore.CYAN)
print("-----------------------------------------------------------------------------------------")
print("                       0б СЛИВАЕМ ФАЙЛЫ В ОДНУ ТАБЛИЦУ version: %s                       "%VERSION)
print("-----------------------------------------------------------------------------------------")
print(Fore.BLUE + '\nЕсть вопросы к разработчикам?\nПишите: ' + Fore.CYAN + 'TsvetkovDS@trcont.ru\n' + Fore.BLUE  + 'Будем рады Вам помочь!',
      sep = '\n')


print(Fore.YELLOW + '\n',datetime.datetime.now(),'\tОбработка стартовала...\n')

#
amount_df = pd.DataFrame()
for file_to_proceed in files_to_proceed_df.itertuples():


      try:
            file_to_proceed_row          = file_to_proceed[0]
            file_to_proceed_folder       = file_to_proceed[1]
            file_to_proceed_book         = file_to_proceed[2]
            file_to_proceed_sheet        = file_to_proceed[3]  
            file_to_proceed_start_row    = int(file_to_proceed[5])
            
            """
            print(file_to_proceed_row,
                  file_to_proceed_folder,
                  file_to_proceed_book,
                  file_to_proceed_sheet,
                  file_to_proceed_start_row,
                  sep = '\n')
            """     
      
            file_to_proceed_df = pd.read_excel(os.path.join(file_to_proceed_folder,file_to_proceed_book),
                              file_to_proceed_sheet,
                              header=file_to_proceed_start_row
                              )

            #print(file_to_proceed_df)
            #input()
            file_to_proceed_df = file_to_proceed_df.dropna(axis=0, how = 'all').dropna(axis=1, how = 'all')

            #print(file_to_proceed_df)

            #break
            file_to_proceed_df['source_folder'] = file_to_proceed_folder
            file_to_proceed_df['source_book']   = file_to_proceed_book    
            file_to_proceed_df['source_sheet']  = file_to_proceed_sheet
            file_to_proceed_df['Строка в исходнике'] = file_to_proceed_df.index + file_to_proceed_start_row + 2
            print(Fore.GREEN + '%-5s'%file_to_proceed_row,
            '%-40s'%file_to_proceed_book,
            '%-20s'%file_to_proceed_sheet,
            '%7s'%len(file_to_proceed_df),
            sep = '\t')
            amount_df = amount_df._append(file_to_proceed_df, ignore_index=True)


      except:
            print(Fore.WHITE + Back.RED + 
                  'Папка: %s'%(file_to_proceed_folder) + '\t' + 
                  Back.BLACK + Fore.WHITE + Style.BRIGHT)            
            print(Fore.WHITE + Back.RED + 
                  'Книга: %s'%(file_to_proceed_book) + '\t' + 
                  Back.BLACK + Fore.WHITE + Style.BRIGHT)            
            print(Fore.WHITE + Back.RED + 
                  'Лист:  %s'%(file_to_proceed_sheet) + '\t' + 
                  Back.BLACK + Fore.WHITE + Style.BRIGHT)
            print(Fore.WHITE + Back.RED + 
                  'Содержимое невозможно обработать!' + '\t' + 
                  Back.BLACK + Fore.WHITE + Style.BRIGHT)
            
            messagebox.showerror(message_title,
                                 'Содержимое невозможно обработать!'
                                 )
            continue
      amount_df_columns_list = amount_df.columns
      service_columns_list = ['source_folder', 'source_book', 'source_sheet', 'Строка в исходнике']
      output_columns_list = service_columns_list + [i for i in amount_df_columns_list if i not in service_columns_list]
      output_df = amount_df[output_columns_list]
      #print(amount_df)
      #print(output_df)     
    


    #print(file_to_proceed_df)
    #break
print(Fore.YELLOW + '\n',
      datetime.datetime.now(),
      '\tЗаписываем в 0в.ПРЕДОБРАБОТКА из ПИТОНА.csv %s строк...'%len(amount_df)
      )
#amount_df.to_csv('result_%s.csv'%datetime.datetime.now().strftime('%H-%M'), sep='\t', encoding='utf-8', decimal=',')
attempt_status = True

while attempt_status:
      try:
            output_df.to_csv('_CONCAT_YOUR_EXCELS.csv', sep='\t', encoding='utf-8', decimal=',')
            attempt_status = False
      except PermissionError:
            messagebox.showerror(message_title,
                                 'Невозможно запиисать результат в файл _CONCAT_YOUR_EXCELS.csv\nВозможно он открыт в другой программе.\nПопробуйте закрыть его и нажмите ОК, чтобы повторить попытку записи!')      

print(Fore.YELLOW + '\n',
      datetime.datetime.now(),
      '\tГотово!\n\n',
      'Нажмите ввод чтобы выйти из программы...',
      )

input()


"""
print('записываем в эксель...', datetime.datetime.now())
amount_df.to_excel('result_%s.xlsx'%datetime.datetime.now().strftime('%H-%M'))
print('Эксель готово...', datetime.datetime.now())
"""