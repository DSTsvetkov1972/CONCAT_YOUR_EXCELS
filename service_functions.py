import tkinter

VERSION = 2
TITLE = "CONCAT_YOUR_EXCELS v/%s"%(VERSION)
developers_message = """
    Приложение работает не так как ожидалось?
    Есть идеи что добавить или улучшить?
    Хотите угостить разработчиков кофе?
    Всегда рады будем с Вами пообщаться!
    Пишите нам на электронку:
    TsvetkovDS@trcont.ru
"""

def keys(event): # Функция чтобы работала вставка из буфера в русской раскладке
    import ctypes
    u = ctypes.windll.LoadLibrary("user32.dll")
    pf = getattr(u, "GetKeyboardLayout")
    if hex(pf(0)) == '0x4190419':
        keyboard_layout = 'ru'
    if hex(pf(0)) == '0x4090409':
        keyboard_layout = 'en'

    if keyboard_layout == 'ru':
        if event.keycode==86:
            event.widget.event_generate("<<Paste>>")
        elif event.keycode==67: 
            event.widget.event_generate("<<Copy>>")    
        elif event.keycode==88: 
            event.widget.event_generate("<<Cut>>")    
        elif event.keycode==65535: 
            event.widget.event_generate("<<Clear>>")
        elif event.keycode==65: 
            event.widget.event_generate("<<SelectAll>>")

def show_message(message, TITLE):
    root = tkinter.Tk()
    root.title(TITLE)
    root.geometry('380x160')


    developers_info_text = tkinter.Text(root,wrap=tkinter.WORD)

    developers_info_text.bind("<Control-KeyPress>", keys)
    developers_info_text.insert('1.0', message)
    developers_info_text.configure(state='disabled')
    developers_info_text.pack()
    root.mainloop() 