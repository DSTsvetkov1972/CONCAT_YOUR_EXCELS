from tkinter import *
from tkinter import ttk
import functions as fn
from service_functions import *
import os
from service_functions import *
from colorama import init, Fore, Back, Style

import warnings
warnings.filterwarnings('ignore')


init()
print(Style.BRIGHT)
print(Fore.MAGENTA)
print("-----------------------------------------------------------------------------------------")
print("                                  %s                      "%(TITLE))
print("-----------------------------------------------------------------------------------------")

if fn.check_files_isnt_open(TITLE):

      
      tables_from_sheets_dict = {}
      sheets_for_processing_list = []


      root = Tk()

      root.title(TITLE)
      root.geometry("330x250")
      root.resizable(height=False, width=False)

      #reload_button = ttk.Button(root, text ="Отладчик", width = 30, command = lambda:reload(fn))
      #reload_button.pack(anchor = CENTER, pady=(25,0))



      if not os.path.exists('Исходники'):
            os.mkdir('Исходники')

      fn.get_sheets()

      show_sheets_button = ttk.Button(root, text ="Просмотреть листы", width = 30, command = fn.get_sheets)
      show_sheets_button.pack(anchor = CENTER, pady=(25,0))

      open_headers_button = ttk.Button(root, text ="Собрать заголовки", width = 30, command = lambda:fn.open_headers_xls(tables_from_sheets_dict,sheets_for_processing_list,concat_tables_button))
      open_headers_button.pack(anchor = CENTER, pady=(25,0))

      concat_tables_button = ttk.Button(root, text ="Объединить таблицы", width = 30, command = lambda:fn.concat_tables(tables_from_sheets_dict, sheets_for_processing_list, concat_tables_button))


      mainmenu = Menu(root)
      root.config(menu=mainmenu)
      mainmenu.add_command(label='Контакты разработчиков', command = lambda: show_message(developers_message, TITLE))
      root.mainloop()
