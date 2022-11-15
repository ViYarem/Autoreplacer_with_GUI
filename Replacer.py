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
COMBINATION_S = {keyboard.Key.shift, keyboard.KeyCode(char='`')}
#COMBINATION_E = {keyboard.Key.shift, keyboard.Key.space}
current_s = set()
#current_e = set()
board = Controller()
opt = False
macro_starter = '`'
macro_ender = Key.space
listening = False
typed_keys = []
ex = False
longest_string = 0
z = 0
n = 0
# ---------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------Excel


def fexcel():
    '''check if excel with replacements exists: no-False,yes-create dictionary {keywords:replacement}'''
    global replacements
    if os.path.isfile('autoreplacement.xlsx') == False:
        return False
    else:
        df = pd.read_excel('autoreplacement.xlsx', sheet_name=0)[
            ['Keyword', 'Replacement']].dropna()
        df['Keyword'] = df['Keyword'].astype('str')
        df['Replacement'] = df['Replacement'].astype('str')
        replacements = df.set_index('Keyword')['Replacement'].to_dict()
        return True


def nofile():
    '''create necessary excel with example'''
    workbook = xlsxwriter.Workbook('autoreplacement.xlsx', {
                                   'strings_to_numbers':  False, 'number_to_string':  False})
    worksheet = workbook.add_worksheet()
    cf1 = workbook.add_format({'bg_color': 'green'})
    worksheet.set_column(0, 3, 70)
    worksheet.write('A1', 'Keyword', cf1)
    worksheet.write('B1', 'Replacement', cf1)
    worksheet.write('A2', 'hi')
    worksheet.write('B2', 'Hi, how are you?')
    worksheet.write('A3', 'br')
    worksheet.write('B3', 'Best regards,')
    worksheet.write('A4', '1')
    worksheet.write('B4', 'Please correct as shown at ')
    cf2 = workbook.add_format({'bg_color': 'red'})
    cf3 = workbook.add_format({'bg_color': 'yellow'})
    worksheet.write(
        'C1', 'A1 (Keyword) and B1 (Replacement) must not be changed', cf2)
    worksheet.write(
        'C2', 'If cell Ax is not empty, Bx must also not be empty, and vice versa', cf2)
    worksheet.write(
        'C3', 'Keywords are not sensitive to capital and small letters, and keyboard layout', cf2)
    worksheet.write(
        'C4', '(qwe and QWE and Qwe and йце (cyrillic keyboard layout) - the same keyword with 1st replacement)', cf2)
    worksheet.write(
        'C5', 'Keywords must not include gaps(space)', cf2)
    worksheet.write(
        'C6', 'All questions and suggestions write to Vira Yaremchuk', cf3)
    workbook.close()
    root = Tk()
    root.configure(bg='#1C2833')
    w1 = Label(root, text='\n1) A1 (Keyword) and B1 (Replacement) must not be changed', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w1.pack(side=TOP, anchor=W)
    w2 = Label(root, text='2) If cell Ax is not empty, Bx must also not be empty, and vice versa', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w2.pack(side=TOP, anchor=W)
    w3 = Label(root, text='3) Keywords are not sensitive to capital and small letters, and keyboard layout (qwe=QWE=йце)', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w3.pack(side=TOP, anchor=W)
    w31 = Label(root, text='    (qwe = QWE = QwE = йце = йwe)', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w31.pack(side=TOP, anchor=W)
    w4 = Label(root, text='4) Keywords must not include gaps(space)\n', font=(
        "Arial", 15), fg='#7FB3D5', bg='#1C2833')
    w4.pack(side=TOP, anchor=W)
    messagebox.showerror("ERROR: autoreplacement.xlsx NOT FOUND",
                         "Necessary autoreplacement.xlsx was not found in directory:(\n\n\n\nScript has created this file for you :)\n\nYou can fill it using rules in 2nd window\n Please read them carefully")

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
        '''tkinter window destroy'''
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
        '''buttons styles'''
        global opt
        opt = not opt
        if close_button["state"] == "normal":
            close_button["state"] = "disabled"
            close_button['background'] = "white"
            close_button.bind("<Enter>", on_enter3)
            close_button.bind("<Leave>", on_leave3)
            lbl_title["text"] = "ACTIVATED"
            lbl_title["foreground"] = "#83f28f"
            button["text"] = "STOP"
        else:
            close_button["state"] = "normal"
            close_button['background'] = '#212F3D'
            close_button.bind("<Enter>", on_enter)
            close_button.bind("<Leave>", on_leave2)
            lbl_title["text"] = "DEACTIVATED"
            lbl_title["foreground"] = "#f94449"
            button["text"] = "START"

    def listbox_copy(event):
        '''copy row from listbox'''
        j = 0
        window.clipboard_clear()
        selected = mylist.get(ANCHOR)
        selected_index = mylist.curselection()[0]
        print(selected_index)
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
    lbl_title = Label(title_bar, text="ACTIVATED ", font=(
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
    lbl1 = Label(window, text="Close is enabled only when replacer is deactivated", font=(
        "Arial", 9), bg='#1C2833', fg='#AAAAAA')
    lbl1.pack(anchor=E)
    lbl2 = Label(window, text="PRESS START to activate, STOP to deactivate", font=(
        "Arial", 11), fg='#7FB3D5', bg='#1C2833')
    lbl2.pack()
    button = Button(window, text='STOP', font=("Arial Bold", 14), pady=12, bg='#212F3D', fg='#7FB3D5', activebackground="#000000", activeforeground="#FFFFFF",
                    command=changeText)
    button.pack(fill=BOTH, expand=True)
    button.bind("<Enter>", on_enter2)
    button.bind("<Leave>", on_leave2)
    lbl3 = Label(
        window, text="Check keyboard layout (US/UKR)!-->Press shift+` to start->\n->(char * will be auto-printed)-->Press keys combination->\n->Press space key (gap) to apply replacement", font=("Arial", 10), bg='#1C2833', fg='#AAAAAA')
    lbl3.pack(side=TOP)
    lbl4 = Label(window, text="LIST OF REPLACEMENTS", font=(
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
    lbl5 = Label(window, text="double-click row to copy",
                 font=("Arial", 10), bg='#1C2833', fg='#AAAAAA')
    lbl5.pack(side=TOP, anchor=E)
    window.mainloop()

# ---------------------------------------------------------MAIN
# create shortcut begin --->


def comb_press(key):
    global COMBINATION_S
    #global COMBINATION_E
    global current_s
    #global current_e
    if key in COMBINATION_S:
        current_s.add(key)
        if all(k in current_s for k in COMBINATION_S):
            return 'start'
        # elif all(k in current_e for k in COMBINATION_E):
            # return 'end'


def on_release(key):
    global current_s
    #global current_e
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
    '''on_press function for replacement. Activated by ctrl+`symbol, disable by space key. If combination of letters meets combination in excel it will be removed and changed'''
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
                    print(typed_keys)
                    typed_keys.append(key.char)
                    '''if macro_starter in typed_keys:
                        typed_keys.remove(macro_starter)
                        if st in typed_keys:
                        typed_keys.remove(st)
                        n -= 1'''
                if key == Key.backspace and len(typed_keys) != 0:
                    n -= 2
                    typed_keys.pop()
                if key == Key.backspace and len(typed_keys) == 0:
                    listening = False

            else:
                listening = False

            if key == Key.space:
                listening = False
                #listening = True
                if typed_keys[0] == '`' and typed_keys[1] == '*':
                    print(typed_keys[2:])
                    candidate_keyword = ""
                    candidate_keyword = candidate_keyword.join(
                        typed_keys[2:])
                    if candidate_keyword != "":
                        if candidate_keyword in replacements.keys():
                            # if candidate_keyword.isalpha() and replacements[candidate_keyword].isalpha():
                            pyautogui.press(
                                'backspace', presses=len(candidate_keyword)+3)
                            pyautogui.typewrite(
                                replacements[candidate_keyword])
                            listening = False
                        '''else:
                            pyautogui.press(
                                'backspace', presses=len(candidate_keyword)+2)
                            pyautogui.press(
                                'shift')
                            pyautogui.typewrite(
                                replacements[candidate_keyword])
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
        dem()
        # ---------------------------------------------------------LISTENER
    else:
        nofile()
