from tkinter import *
from tkinter import ttk
import functions as fn
import os
#from importlib import reload

from functions import VERSION, TITLE
from colorama import init, Fore, Back, Style

import warnings
warnings.filterwarnings('ignore')


init()
print(Style.BRIGHT)
print(Fore.CYAN)
print("-----------------------------------------------------------------------------------------")
print("                                  %s                      "%(TITLE))
print("-----------------------------------------------------------------------------------------")
print(Fore.BLUE + '\nЕсть вопросы к разработчикам?\nПишите: ' + Fore.CYAN + 'TsvetkovDS@trcont.ru\n' + Fore.BLUE  + 'Будем рады Вам помочь!',
      sep = '\n')

tables_from_sheets_dict = {}
sheets_for_processing_list = []


root = Tk()

root.title(TITLE)
root.geometry("250x250")

#reload_button = ttk.Button(root, text ="Отладчик", width = 30, command = lambda:reload(fn))
#reload_button.pack(anchor = CENTER, pady=(25,0))

get_sheets_button = ttk.Button(root, text ="Собрать список листов", width = 30, command = lambda:fn.get_sheets(show_sheets_button, get_headers_button))
get_sheets_button.pack(anchor = CENTER, pady=(25,0))

show_sheets_button = ttk.Button(root, text ="Просмотреть листы", width = 30, command = fn.show_sheets)
if os.path.exists(os.path.join(os.getcwd(),'.sheets.csv')):
      show_sheets_button.pack(anchor = CENTER, pady=(25,0))

get_headers_button = ttk.Button(root, text ="Найти заголовки", width = 30, command = lambda:fn.open_headers_xls(tables_from_sheets_dict,sheets_for_processing_list,concat_tables_button))
if os.path.exists(os.path.join(os.getcwd(),'.sheets.csv')):
      get_headers_button.pack(anchor = CENTER, pady = (25,0), padx=(0,0))


concat_tables_button = ttk.Button(root, text ="Объединить таблицы", width = 30, command = lambda:fn.concat_tables(tables_from_sheets_dict,sheets_for_processing_list))


root.mainloop()
