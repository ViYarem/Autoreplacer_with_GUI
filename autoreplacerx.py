from pynput.keyboard import Key, Listener
from tkinter import *
from tkinter import messagebox
from tkinter import ttk
import threading
import pyautogui
import string
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
replacements = {}
opt = True
alphabet = list(string.ascii_lowercase)
macro_starter = '`'
macro_ender = Key.shift
listening = True
typed_keys = []
ex = False
liskey = []
longest_string = 0
z = 0
# ---------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------


def fexcel():
    global replacements
    if os.path.isfile('autoreplacement.xlsx') == False:
        return False
    else:
        df = pd.read_excel('autoreplacement.xlsx', sheet_name=0)
        test_keys = df["Keyword"].values.tolist()
        test_values = df["Replacement"].values.tolist()
        replacements = {test_keys[i]: test_values[i]
                        for i in range(len(test_keys))}
        for k in tuple(replacements.keys()):
            for i in k:
                if i not in alphabet:
                    del replacements[k]
        for w, r in tuple(replacements.items()):
            for a in r:
                if a.isdigit():
                    del replacements[w]
        return True


def nofile():
    workbook = xlsxwriter.Workbook('autoreplacement.xlsx')
    worksheet = workbook.add_worksheet()
    cf1 = workbook.add_format({'bg_color': 'yellow'})
    worksheet.set_column(1, 3, 60)
    worksheet.write('A1', 'Keyword', cf1)
    worksheet.write('B1', 'Replacement', cf1)
    worksheet.write('A2', 'br')
    worksheet.write('B2', 'Best regards,')
    cf2 = workbook.add_format({'bg_color': 'red'})
    worksheet.write(
        'C1', 'A1 (Keyword) and B1 (Replacement) must not be changed', cf2)
    worksheet.write(
        'C2', 'If cell Ax is not empty, Bx must also not be empty, and vice versa', cf2)
    worksheet.write(
        'C3', 'Column A has only a combination of lowercase Latin letters', cf2)
    worksheet.write(
        'C4', 'Column B must not contain numbers in the values', cf2)
    workbook.close()
    root = Tk()
    root.configure(bg='#1C2833')
    w = Label(root, text='\nRules for filling autoreplacement.xlsx:\n\n\n1) Column A has only a combination of lowercase Latin letters\n\n2) Column B must not contain numbers in the values\n\n3) A1 (Keyword) and B1 (Replacement) must not be changed \n\n4) If cell Ax is not empty, Bx must also not be empty, and vice versa\n', font=(
        "Arial Bold", 15), fg='#7FB3D5', bg='#1C2833')
    w.pack(side=TOP, anchor=SW)
    messagebox.showerror("ERROR: autoreplacement.xlsx NOT FOUND",
                         "Necessary autoreplacement.xlsx was not found in directory:(\n\n\n\nScript has created this file for you :)\n\nYou can fill this file using rules in 2nd window\nor replace auto-created file by own")

# ---------------------------------------------------------MENU


def on_enter(e):
    e.widget['background'] = '#F1948A'
    e.widget['foreground'] = "#000"


def on_enter2(e):
    e.widget['background'] = '#28c7e1'
    e.widget['foreground'] = "#000"


def on_leave(e):
    e.widget['background'] = '#212F3D'
    e.widget['foreground'] = "#7FB3D5"


def on_leave2(e):
    e.widget['background'] = '#212F3D'
    e.widget['foreground'] = "#000"


def on_enter3(e):
    e.widget['background'] = 'white'
    e.widget['foreground'] = "grey"


def on_leave3(e):
    e.widget['background'] = 'white'
    e.widget['foreground'] = "grey"


def makeSomething(value):
    global opt
    opt = value


def full_exit(value):
    global ex
    ex = value


def menu():
    global opt
    global liskey
    global longest_string
    global ex
    global z
    window = Tk()

    def move_window(event):
        window.geometry('+{0}+{1}'.format(event.x_root, event.y_root))

    def w_exit():
        global ex
        ex = True
        window.quit()
        # window.destroy()

    def minimizeGUI():
        global z
        window.state('withdrawn')
        window.overrideredirect(False)
        window.state('iconic')
        z = 1

    def frameMapped(event=None):
        global z
        window.overrideredirect(True)
        # window.iconbitmap("ANAL_OG.ico")
        if z == 1:
            set_appwindow(window)
            z = 0

    def set_appwindow(window):
        # Honestly forgot what most of this stuff does. I think it's so that you can see
        # the program in the task bar while using overridedirect. Most of it is taken
        # from a post I found on stackoverflow.
        hwnd = windll.user32.GetParent(window.winfo_id())
        stylew = windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
        stylew = stylew & ~WS_EX_TOOLWINDOW
        stylew = stylew | WS_EX_APPWINDOW
        #res = windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, stylew)
        # re-assert the new window style
        window.wm_withdraw()
        window.after(10, lambda: window.wm_deiconify())

    window.overrideredirect(True)

    title_bar = Frame(window, bg='#1C2833', relief='raised', bd=0)
    title_bar.pack(fill='x', expand=True)
    # title_label = Label(window, text="Replacer is working",
    # bg='#1C2833', fg='grey')
    #title_label.pack(side=LEFT, pady=2)
    close_button = Button(title_bar, text='X', command=w_exit, relief="raised", bg='#212F3D', padx=2,
                          pady=2, bd=1, font="bold", fg='white', highlightthickness=0)
    close_button.pack(side=RIGHT)
    close_button.bind("<Enter>", on_enter)
    close_button.bind("<Leave>", on_leave2)

    min_button = Button(title_bar, text='-', command=minimizeGUI, relief="raised", bg='#212F3D', padx=2,
                        pady=2, bd=1, font="bold", fg='white', highlightthickness=0)
    min_button.pack(side=RIGHT)
    min_button.bind("<Enter>", on_enter2)
    min_button.bind("<Leave>", on_leave2)

    window.wm_attributes('-alpha', 0.92)
    window.attributes('-topmost', 1)
    #window.title("Replacer is working")
    window.configure(bg='#1C2833')
    window.bind('<B1-Motion>', move_window)
    window.bind("<Map>", frameMapped)
    styl = ttk.Style()
    styl.configure('TSeparator', background='grey')
    '''separator = ttk.Separator(
        window, orient='horizontal', style='TSeparator', class_=ttk.Separator)
    separator.pack(fill='x', expand=True)'''
    lbl = Label(window, text="Switch: start - activated, stop - deactivated", font=(
        "Arial Bold", 14), fg='#7FB3D5', bg='#1C2833')
    lbl.pack()
    lbl = Label(window, text="!closing is possible only in the deactivated state", font=(
        "Arial", 10), bg='#1C2833', fg='grey')
    lbl.pack()

    def changeText():
        makeSomething(not opt)
        if close_button["state"] == "normal":
            close_button["state"] = "disabled"
            close_button['background'] = "white"
            close_button.bind("<Enter>", on_enter3)
            close_button.bind("<Leave>", on_leave3)
            button["text"] = "STOP"
        else:
            close_button["state"] = "normal"
            close_button['background'] = '#212F3D'
            close_button.bind("<Enter>", on_enter)
            close_button.bind("<Leave>", on_leave2)
            button["text"] = "START"

    button = Button(window, text='START', font=("Arial Bold", 13), pady=15, bg='#212F3D', fg='#7FB3D5', activebackground="#F67280", activeforeground="#000",
                    command=changeText)
    button.pack(fill=BOTH, expand=True)
    button.bind("<Enter>", on_enter)
    button.bind("<Leave>", on_leave)
    lbl = Label(
        window, text="   Press ` key on keyboard to activate autoreplacement\n   Then press combination of keys which creates keyword\n(!the order in which the keys are used is important, \nnot the written word)\n  Then press Shift button to apply replacement", font=("Arial", 10), bg='#1C2833', fg='grey')
    lbl.pack(side=TOP, anchor=SW)
    lisrR = ['LIST OF REPLACEMENTS:']
    for key in replacements:
        liskey.append(key)
        rowy = key+'       ' + replacements[key]
        lisrR.append(rowy)
    longest_string = len(max(liskey, key=len))
    for x in range(len(lisrR)):
        lbl = Label(window, text='   '+lisrR[x], font=(
            "Arial Bold", 11), fg='#98AFC7', bg='#1C2833')
        lbl.pack(side=TOP, anchor=SW)
    window.mainloop()

# ---------------------------------------------------------MAIN


def on_press(key):
    '''on_press function for replacement. Activated by `symbol, disable by space button. If combination of letters meets combination in excel it will be removed and changed'''
    global opt
    global typed_keys
    global listening
    global alphabet
    global replacements
    global longest_string
    global ex
    if opt == False:
        if hasattr(key, 'char') and key.char == macro_starter:
            typed_keys = []
            listening = True
        if listening:
            if hasattr(key, 'char'):
                for a in alphabet:
                    if len(typed_keys) <= longest_string:
                        if key.char == a:
                            typed_keys.append(a)
                    else:
                        typed_keys = []
            if key == macro_ender:
                candidate_keyword = ""
                candidate_keyword = candidate_keyword.join(typed_keys)
                if candidate_keyword != "":
                    if candidate_keyword in replacements.keys():
                        pyautogui.press(
                            'backspace', presses=len(candidate_keyword)+1)
                        pyautogui.typewrite(
                            replacements[candidate_keyword])
                        listening = False
    elif ex == True:
        return False


if __name__ == '__main__':
    if fexcel() == True:
        # ---------------------------------------------------------THREADING
        t = threading.Thread(target=menu)
        t.start()
        # ---------------------------------------------------------LISTENER
        with Listener(on_press=on_press) as listener:
            listener.join()

    else:
        nofile()
