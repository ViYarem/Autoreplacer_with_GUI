# Script for automatically replacing the Скорочення by replacement word + GUI
# version 1.0. ViYarem
from pynput.keyboard import Key, Listener, Controller
from pynput import keyboard
from tkinter import *
from tkinter import messagebox
import threading
import pyautogui
import pandas as pd
import openpyxl
import os
import xlsxwriter
from ctypes import windll

# Some WindowsOS styles, required for task bar integration
GWL_EXSTYLE = -20
WS_EX_APPWINDOW = 0x00040000
WS_EX_TOOLWINDOW = 0x00000080
# ---------------------------------------------------------
# VARIABLES
# ---------------------------------------------------------
COMBINATION_S = {keyboard.Key.ctrl_l, keyboard.Key.alt_l}
# COMBINATION_E = {keyboard.Key.shift, keyboard.Key.space}
current_s = set()
# current_e = set()
board = Controller()
opt = False
listening = False
typed_keys = []
ex = False
longest_string = 0
z = 0
n = 0
symbols = (u"qwertyuiop[]asdfghjkl;'zxcvbnm,.\\",
           u"йцукенгшщзхїфівапролджєячсмитьбюґ"
           )
tr1 = {ord(a): ord(b) for a, b in zip(*symbols)}
tr2 = {ord(b): ord(a) for a, b in zip(*symbols)}
# ---------------------------------------------------------
# FUNCTIONS


def cyr(i):
    return i.translate(tr1)


def lat(j):
    return j.translate(tr2)
# ---------------------------------------------------------Excel


def fexcel():
    '''check if excel with replacements exists: no-False,yes-create dictionary {Скороченняs:replacement}'''
    global replacements
    if os.path.isfile('Автозаміни.xlsx') == False:
        return False
    else:
        df = pd.read_excel('Автозаміни.xlsx', sheet_name=0)[
            ['Скорочення', 'Повна фраза']].dropna()
        df['Скорочення'] = df['Скорочення'].astype('str')
        df['Повна фраза'] = df['Повна фраза'].astype('str')
        replacements = df.set_index('Скорочення')['Повна фраза'].to_dict()
        return True


def nofile():
    '''create necessary excel with filling example'''
    workbook = xlsxwriter.Workbook('Автозаміни.xlsx', {
                                   'strings_to_numbers':  False, 'number_to_string':  False})
    worksheet = workbook.add_worksheet()
    cf1 = workbook.add_format({'bg_color': 'green'})
    worksheet.set_column(0, 3, 70)
    worksheet.write('A1', 'Скорочення', cf1)
    worksheet.write('B1', 'Повна фраза', cf1)
    worksheet.write('A2', 'т-м')
    worksheet.write('B2', 'Розписання товарно-матеріальних цінностей')
    worksheet.write('A3', 'шр')
    worksheet.write('B3', 'Перенесення змін у штатний розпис')
    worksheet.write('A4', '1')
    worksheet.write('B4', 'Про довезення учнів')
    cf2 = workbook.add_format({'bg_color': 'red'})
    cf3 = workbook.add_format({'bg_color': 'yellow'})
    worksheet.write(
        'C1', 'Не змінюйте клітинки A1 (Скорочення) / B1 (Повна фраза)', cf2)
    worksheet.write(
        'C2', 'Якщо клітинка Ax - не порожня, то клітинка Bx також мусить бути заповнена і навпаки', cf2)
    worksheet.write(
        'C3', 'Використовуйте кирилицю для скорочень', cf2)
    worksheet.write(
        'C4', 'Скорочення повинні складитися з малих літер кирилиці та інших символів', cf2)
    worksheet.write(
        'C5', 'Скорочення не повинні мати пропуски, латиницю (англ. літери) та великі літери', cf2)
    worksheet.write(
        'C6', 'Повна фраза може містити всі символи (в т.ч. латиницю) та пропуски', cf2)
    # worksheet.write(
    # 'C6', 'All questions and suggestions write to Vira Yaremchuk', cf3)
    workbook.close()
    root = Tk()
    root.configure(bg='#1C2833')
    w1 = Label(root, text='\n1) Не змінюйте клітинки A1 (Скорочення) / B1 (Повна фраза)', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w1.pack(side=TOP, anchor=W)
    w2 = Label(root, text='2) Якщо клітинка Ax - не порожня, то клітинка Bx також мусить бути заповнена і навпаки', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w2.pack(side=TOP, anchor=W)
    w3 = Label(root, text='3) Скорочення повинні складитися з малих літер кирилиці та інших символів', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w3.pack(side=TOP, anchor=W)
    w4 = Label(root, text='4) Скорочення не повинні мати пропуски, латиницю (англ. літери) та великі літери', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w4.pack(side=TOP, anchor=W)
    w5 = Label(root, text='5) Повна фраза може містити всі символи (в т.ч. латиницю) та пропуски', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w5.pack(side=TOP, anchor=W)
    messagebox.showerror("Автозаміни.xlsx не знайдений",
                         "Автозаміни.xlsx не знайдений в поточній директорії:(\n\n\n\nСкрипт створив його для Вас самостійно :)\n\nВи можете заповнити файл користуючись правилами в 2-му вікні")

# ---------------------------------------------------------GUI
# mouse's motion styles


def on_enter(e):
    e.widget['background'] = '#F1948A'
    e.widget['foreground'] = '#7FB3D5'


def on_enter2(e):
    e.widget['background'] = '#0479B4'
    e.widget['foreground'] = '#7FB3D5'


def on_leave(e):
    e.widget['background'] = '#212F3D'
    e.widget['foreground'] = "#7FB3D5"


def on_leave2(e):
    e.widget['background'] = '#212F3D'
    e.widget['foreground'] = '#7FB3D5'


def on_enter3(e):
    e.widget['background'] = 'white'
    e.widget['foreground'] = '#7FB3D5'


def on_leave3(e):
    e.widget['background'] = 'white'
    e.widget['foreground'] = '#7FB3D5'

# Main GUI


def menu():
    global opt
    global liskey
    global longest_string
    global ex
    global z
    global listening

    window = Tk()

    def move_window(event):
        '''moving possibility'''
        window.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

    def w_exit():
        '''tkinter window destroy and variable to stop listening of keyboards'''
        global ex
        global listening
        listening = False
        ex = True
        window.quit()

    # for Minimize begin -->

    def minimizeGUI():
        global z
        window.state('withdrawn')
        window.overrideredirect(False)
        window.state('iconic')
        z = 1

    def frameMapped(event=None):
        global z
        window.overrideredirect(True)
        if z == 1:
            set_appwindow(window)
            z = 0

    def set_appwindow(window):
        hwnd = windll.user32.GetParent(window.winfo_id())
        stylew = windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        stylew = stylew & ~WS_EX_TOOLWINDOW
        stylew = stylew | WS_EX_APPWINDOW
        window.wm_withdraw()
        window.after(10, lambda: window.wm_deiconify())

    # <--for Minimize end

    def changeText():
        '''buttons styles and variable to stop/start replacement by keyboard'''
        global opt
        opt = not opt
        if close_button["state"] == "normal":
            close_button["state"] = "disabled"
            close_button['background'] = "white"
            close_button.bind("<Enter>", on_enter3)
            close_button.bind("<Leave>", on_leave3)
            lbl_title["text"] = "АКТИВНИЙ"
            lbl_title["foreground"] = "#83f28f"
            lbl1["text"] = "Закриття можливе після зупинки (натисніть кнопку нижче)"
            button["text"] = "ЗУПИНИТИ"
        else:
            close_button["state"] = "normal"
            close_button['background'] = '#212F3D'
            close_button.bind("<Enter>", on_enter)
            close_button.bind("<Leave>", on_leave2)
            lbl_title["text"] = "НЕ АКТИВНИЙ"
            lbl_title["foreground"] = "#f94449"
            lbl1["text"] = "Автозамінник - зупинений, можете закривати програму"
            button["text"] = "РОЗПОЧАТИ"

    def listbox_copy(event):
        '''copy row from listbox by mouse left key double-clicking'''
        j = 0
        window.clipboard_clear()
        selected = mylist.get(ANCHOR)
        selected_index = mylist.curselection()[0]
        for i in range(len(selected)):
            if selected[i] == '>':
                j = i+2
        window.clipboard_append(selected[j:])
        mylist.selection_clear(selected_index, END)
        for i in range(mylist.size()):
            if i == selected_index:
                mylist.itemconfig(i, bg='#0479B4')
            else:
                mylist.itemconfig(i, bg='#1C2833')
        mylist.itemconfig(selected_index, bg='#1C2833')

    window.wm_attributes('-alpha', 0.85)
    window.attributes('-topmost', 1)
    window.configure(bg='#1C2833')
    window.bind("<Map>", frameMapped)
    window.overrideredirect(True)
    title_bar = Frame(window, bg='#1C2833', relief='raised', bd=0)
    title_bar.pack(fill='x', expand=True)
    title_bar.bind('<B1-Motion>', move_window)
    lbl_title = Label(title_bar, text="АКТИВНИЙ ", font=(
        "Arial", 12), bg='#1C2833', fg="#83f28f")
    lbl_title.pack(side=LEFT)
    close_button = Button(title_bar, text='X', command=w_exit, relief="raised", bg='white', padx=2,
                          pady=2, bd=1, font="bold", fg='#7FB3D5', highlightthickness=0, state="disabled")
    close_button.pack(side=RIGHT)
    close_button.bind("<Enter>", on_enter3)
    close_button.bind("<Leave>", on_leave3)
    min_button = Button(title_bar, text='-', command=minimizeGUI, relief="raised", bg='#212F3D', padx=2,
                        pady=2, bd=1, font="bold", fg='#7FB3D5', highlightthickness=0)
    min_button.pack(side=RIGHT)
    min_button.bind("<Enter>", on_enter2)
    min_button.bind("<Leave>", on_leave2)
    lbl1 = Label(window, text="Закриття можливе після зупинки (натисніть кнопку нижче)", font=(
        "Arial", 9), bg='#1C2833', fg='#AAAAAA')
    lbl1.pack(anchor=E)
    '''lbl2 = Label(window, text="START - активувати, STOP - зупинити", font=(
        "Arial", 11), fg='#7FB3D5', bg='#1C2833')
    lbl2.pack()'''
    button = Button(window, text='ЗУПИНИТИ', font=("Arial Bold", 14), pady=12, bg='#212F3D', fg='#7FB3D5', activebackground="#000000", activeforeground="#FFFFFF",
                    command=changeText)
    button.pack(fill=BOTH, expand=True)
    button.bind("<Enter>", on_enter2)
    button.bind("<Leave>", on_leave2)
    lbl3 = Label(
        window, text="Натисніть комбінацію зліва ctrl+alt->з'явиться знак *->\n->надрукуйте скорочення-->натисніть пробіл", font=("Arial", 10), bg='#1C2833', fg='#AAAAAA')
    lbl3.pack(side=TOP)
    lbl4 = Label(window, text="СПИСОК СКОРОЧЕНЬ ТА АВТОЗАМІН", font=(
        "Arial", 11), fg='#7FB3D5', bg='#1C2833')
    lbl4.pack(side=TOP, anchor=W)
    liskey = []
    lisrR = ['']
    for key in replacements:
        liskey.append(key)
        rowy = key+' -> ' + replacements[key]
        lisrR.append(rowy)
        lisrR.append(' ')
    longest_string = len(max(liskey, key=len))
    myscrollY = Scrollbar(window, orient='vertical', borderwidth=0,)
    myscrollY.pack(side=RIGHT, fill=Y)
    myscrollX = Scrollbar(window, orient='horizontal', borderwidth=0)
    if len(lisrR) < 38:
        height_number = len(lisrR)
    else:
        height_number = 38
    mylist = Listbox(window, height=height_number, width=43, bd=0, relief=FLAT, font=(
        "Arial", 11), fg='#98AFC7', bg='#1C2833', activestyle='none', yscrollcommand=myscrollY.set, xscrollcommand=myscrollX.set)
    for x in range(len(lisrR)):
        mylist.insert(END, ' '+lisrR[x])
    mylist.bind('<Double-Button-1>', listbox_copy)
    mylist.pack(side=TOP, fill=BOTH, expand=True)
    myscrollY.config(command=mylist.yview)
    myscrollX.pack(side=BOTTOM, fill=X)
    myscrollX.config(command=mylist.xview)
    lbl5 = Label(window, text="!Подвійне натискання ЛК миші копіює фразу",
                 font=("Arial", 10), bg='#1C2833', fg='#AAAAAA')
    lbl5.pack(side=TOP, anchor=E)
    window.mainloop()

# ---------------------------------------------------------MAIN
# create shortcut begin --->


def comb_press(key):
    global COMBINATION_S
    # global COMBINATION_E
    global current_s
    # global current_e
    if key in COMBINATION_S:
        current_s.add(key)
        if all(k in current_s for k in COMBINATION_S):
            return 'start'
        # elif all(k in current_e for k in COMBINATION_E):
            # return 'end'


def on_release(key):
    global current_s
    # global current_e
    try:
        current_s.remove(key)
    except KeyError:
        pass
    # try:
        # current_e.remove(key)
    # except KeyError:
        # pass
# <---create shortcut end


def on_press(key):
    '''on_press function for replacement. АКТИВНИЙ by ctrl+`symbol, disable by space key. If combination of letters meets combination in excel it will be removed and changed'''
    global opt
    global typed_keys
    global listening
    global replacements
    global longest_string
    global ex
    global n
    global board
    st = '*'
    if opt == False and ex == False:
        n += 1
        if comb_press(key) == 'start':
            board.press(st)
            n = 0
            typed_keys = []
            listening = True
        if listening:
            if len(typed_keys) <= longest_string+22 or n <= longest_string+22:
                if hasattr(key, 'char'):
                    typed_keys.append(key.char)
                    '''if macro_starter in typed_keys:
                        typed_keys.remove(macro_starter)
                        if st in typed_keys:
                        typed_keys.remove(st)
                        n -= 1'''
                if key == Key.backspace and len(typed_keys) != 0:
                    n -= 1
                    typed_keys.pop()
                if key == Key.backspace and len(typed_keys) == 0:
                    listening = False

            else:
                listening = False

            if key == Key.space:
                # board.press(Key.space)
                listening = False
                # listening = True
                if typed_keys[0] == '*':
                    candidate_shortcut = ""

                    candidate_shortcut = candidate_shortcut.join(
                        typed_keys[1:])

                    if candidate_shortcut != "":

                        if cyr(candidate_shortcut) in replacements.keys():

                            # if candidate_shortcut.isalpha() and replacements[candidate_shortcut].isalpha():
                            pyautogui.press(
                                'backspace', presses=len(candidate_shortcut)+2)

                            for i in (
                                    replacements[cyr(candidate_shortcut)]):
                                board.press(i)
                            '''pyautogui.typewrite(lat(
                                replacements[cyr(candidate_shortcut)]))'''
                            listening = False
                        '''else:
                            pyautogui.press(
                                'backspace', presses=len(candidate_shortcut)+2)
                            pyautogui.press(
                                'shift')
                            pyautogui.typewrite(
                                replacements[candidate_shortcut])
                            listening = False'''

    elif opt == True and ex == False:
        listening = False

    else:
        return False


def dem():
    with Listener(on_press=on_press, on_release=on_release) as listener:
        listener.join()


if __name__ == '__main__':
    if fexcel() == True:
        # ---------------------------------------------------------THREADING
        t1 = threading.Thread(target=menu)
        t1.start()
        # ---------------------------------------------------------LISTENER
        dem()
    else:
        nofile()
