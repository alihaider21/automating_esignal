import pandas as pd
import os
import pyperclip
import pywinauto
import time
from pywinauto.keyboard import send_keys

current_directory = os.getcwd()
excel_file_path = os.path.join(current_directory, 'Trade_AI data.xlsx') #here you need to change the file name

excel_file = pd.read_excel(excel_file_path)
column_data = excel_file['Symbol'] #here you need to enter the column name which have stock name
list_data = column_data.tolist()

app = pywinauto.Application().connect(path="eSignal.exe")
e_sginal =  app.window(title="eSignal")
e_sginal.set_focus()
e_sginal.right_click_input()
time.sleep(1)
send_keys("{DOWN 22}")
time.sleep(1)
send_keys("{ENTER 1}")


for i in range(3,len(list_data)):
    pyperclip.copy(list_data[i])
    app = pywinauto.Application().connect(path="eSignal.exe")
    e_sginal =  app.window(best_match="eSignal")
    e_sginal.set_focus()
    time.sleep(1)   
    send_keys("{VK_TAB 1}")
    send_keys('^v')
    send_keys("{ENTER 1}")
    time.sleep(5)
    e_sginal.right_click_input()
    time.sleep(1)
    send_keys("{DOWN 12}")
    send_keys("{ENTER 1}")
    time.sleep(1)
    try:
        app = pywinauto.Application().connect(path="eSignal.exe")
        e_sginal =  app.window(title="Export Data")
        e_sginal.set_focus()
        time.sleep(2)   
        send_keys("{ENTER 1}")
        time.sleep(1)
        send_keys("{ENTER 1}")
        time.sleep(1)
        send_keys("{VK_TAB 16}")
        time.sleep(1)
        send_keys("{VK_SPACE 1}")
    except:
        send_keys("+{ENTER}")



