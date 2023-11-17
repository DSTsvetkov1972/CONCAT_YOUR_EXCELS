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



root = Tk()

root.title(TITLE)
root.geometry("250x250")
sheets_button = ttk.Button(root, text ="Собрать информацию о листах", width = 30, command = get_sheets)
sheets_button.pack(anchor = CENTER, pady=(25,0))

sheets_button = ttk.Button(root, text ="Просмотреть информацию листах", width = 30, command = show_sheets)
sheets_button.pack(anchor = CENTER, pady=(25,0))

headers_button = ttk.Button(root, text ="Собрать заголовки", width = 30, command = lambda:open_headers_xls(headers_button))
headers_button.pack(anchor = CENTER, pady = (25,0), padx=(0,0))

headers_button = ttk.Button(root, text ="Объединить таблицы", width = 30, command = concat_tables)

root.mainloop()
