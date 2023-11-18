from tkinter import *
from tkinter import ttk
from functions import *
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
root = Tk()

root.title(TITLE)
root.geometry("250x250")



get_sheets_button = ttk.Button(root, text ="Собрать список листов", width = 30, command = lambda:get_sheets(show_sheets_button, get_headers_button))
get_sheets_button.pack(anchor = CENTER, pady=(25,0))

show_sheets_button = ttk.Button(root, text ="Просмотреть листы", width = 30, command = show_sheets)
if os.path.exists(os.path.join(os.getcwd(),'.sheets.csv')):
      show_sheets_button.pack(anchor = CENTER, pady=(25,0))

get_headers_button = ttk.Button(root, text ="Найти заголовки", width = 30, command = lambda:get_tables_from_sheets(tables_from_sheets_dict,concat_tables_button))
if os.path.exists(os.path.join(os.getcwd(),'.sheets.csv')):
      get_headers_button.pack(anchor = CENTER, pady = (25,0), padx=(0,0))


concat_tables_button = ttk.Button(root, text ="Объединить таблицы", width = 30, command = concat_tables)


root.mainloop()
