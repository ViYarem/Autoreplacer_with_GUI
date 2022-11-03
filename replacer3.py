from pynput.keyboard import Key, Listener, KeyCode
from tkinter import *
import threading
import pyautogui
import string
options = False
replacements = {"ruf": "Please remove under-segmentation on FC at ss",
                "rof": "Please remove over-segmentation on FC at ss"
                }

alphabet = list(string.ascii_lowercase)

macro_starter = '`'
macro_ender = Key.space
listening = True
typed_keys = []


def on_enter(e):
    e.widget['background'] = '#28c7e1'


def on_leave(e):
    e.widget['background'] = 'SystemButtonFace'


def makeSomething(value, xwindow):
    global options
    options = value
    xwindow.destroy()


def on_press(key):
    global options
    global typed_keys
    global listening
    global alphabet
    if options == False:
        if hasattr(key, 'char') and key.char == macro_starter:
            typed_keys = []
            listening = True
        if listening:
            if hasattr(key, 'char'):
                for a in alphabet:
                    if key.char == a:
                        typed_keys.append(a)
            if key == macro_ender:
                candidate_keyword = ""
                candidate_keyword = candidate_keyword.join(typed_keys)
                if candidate_keyword != "":
                    if candidate_keyword in replacements.keys():
                        pyautogui.press(
                            'backspace', presses=len(candidate_keyword)+2)
                        pyautogui.typewrite(replacements[candidate_keyword])
                        listening = False

    # Stop listener
    elif options == True:
        return False


def menu():
    global options
    window = Tk()
    window.title("Replacer is working")
    window.configure(bg='#000000')
    lblx = Label(window, text=" ", bg='#000000')
    lblx.grid(column=0, row=0)
    lbl1 = Label(window, text="If you want to stop replacer, please put on button below:", font=(
        "Arial Bold", 16), fg='#28c7e1', bg='#000000')
    lbl1.grid(column=0, row=1)
    lbl2 = Label(window, text=" ", bg='#000000')
    lbl2.grid(column=0, row=2)
    btn1 = Button(window, text="STOP", font=("Arial", 12), pady=15, activebackground="#28c7e1",
                  activeforeground="#fff", command=lambda: makeSomething(True, window))
    btn1.grid(column=0, row=3, sticky='we')
    btn1.bind("<Enter>", on_enter)
    btn1.bind("<Leave>", on_leave)

    window.mainloop()
    return


t = threading.Thread(target=menu)
t.start()

with Listener(on_press=on_press) as listener:
    listener.join()
