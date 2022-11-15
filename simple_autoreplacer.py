from pynput.keyboard import Key, Listener
from tkinter import *
import threading
import pyautogui
import string
import pandas as pd
import openpyxl
# ---------------------------------------------------------
# VARIABLES
# ---------------------------------------------------------
df = pd.read_excel('autoreplacement.xlsx', sheet_name=0)
test_keys = df["Word for auto-replacement"].values.tolist()
test_values = df["Auto-replacement"].values.tolist()
replacements = {test_keys[i]: test_values[i] for i in range(len(test_keys))}
# ---------------------------------------------------------
options = False
'''replacements = {"ref": "Please remove over-segmentation on FC at ss ",
                "ruf": "Please remove under-segmentation on FC at ss ",
                "ret": "Please remove over-segmentation on TC at ss ",
                "rut": "Please remove under-segmentation on TC at ss ",
                "csfb": "Please correct FB as shown at ",
                "cstb": "Please correct TB as shown at ",
                "cstc": "Please correct TC as shown at ",
                "csfc": "Please correct FC as shown at ",
                "pru": "Please remove undercut ",
                "fcb": "Please correct FC and FB as shown at ss ",
                "tcb": "Please correct TC and TB as shown at ss ",
                "refb": "Please remove over-segmentation on FB at AS ",
                "rufb": "Please remove under-segmentation on FB at AS ",
                "retb": "Please remove over-segmentation on TB at AS ",
                "rutb": "Please remove under-segmentation on TB at AS ",
                "pcas": "Please correct as shown at ",

                "fhc": "Please adjust Femoral Head Center ",
                "pfp": "Please adjust Piriformis Fossa Point position ",
                "ame": "Please adjust Medial Epicondyle Point position ",
                "ale": "Please adjust Lateral Epicondyle Point position ",
                "mnp": "Please adjust Middle Notch Point position ",
                "aap": "Please adjust Anterior Point position ",
                "mpbp": "Please adjust Medial Posterior Bone Point position ",
                "lpbp": "Please adjust Lateral Posterior Bone Point position ",
                "fc": "Please correct Femoral Curve as shown ",
                "dtp": "Please adjust Distal Tibia Point position ",
                "aptp": "Please adjust Proximal Tibia Point position ",
                "fhp": "Please adjust Fibular Head Point position ",
                "msp": "Please adjust Medial Spine Point position ",
                "lsp": "Please adjust Lateral Spine Point position ",
                "mtp": "Please adjust Medial Tuberosity Point position ",
                "ltp": "Please adjust Lateral Tuberosity Point position ",
                "msup": "Please adjust Medial Sulcus Point position ",
                "lsup": "Please adjust Lateral Sulcus Point position ",
                "tmc": "Please correct Tibial Medial Curve as shown ",
                "tlc": "Please correct Tibial Lateral Curve as shown "
                }'''
alphabet = list(string.ascii_lowercase)
macro_starter = '`'
macro_ender = Key.space
listening = True
typed_keys = []
# ---------------------------------------------------------
# FUNCTIONS
# ---------------------------------------------------------MENU


def on_enter(e):
    e.widget['background'] = '#28c7e1'


def on_leave(e):
    e.widget['background'] = 'SystemButtonFace'


def makeSomething(value, xwindow):
    global options
    options = value
    xwindow.destroy()


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
# ---------------------------------------------------------MAIN


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


# ---------------------------------------------------------THREAD
t = threading.Thread(target=menu)
t.start()
# ---------------------------------------------------------LISTENER
with Listener(on_press=on_press) as listener:
    listener.join()
