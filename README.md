# Replacer.py
## Overview
The script does the autoreplacement of the keywords are entered from keyboard. The keywords and words/phrases for replacement are being field in excel file. The script launches a user-friendly GUI which includes small instruction for using.
![image1](https://github.com/ViYarem/Autoreplacer_with_GUI/assets/68001529/1a6a04d2-a9fc-4c50-ac01-b76cd92dbc2d)

## What You Will Need
Python>=3.0;
pynput==1.4.5;
tk==0.1.0;
PyAutoGUI==0.9.53;
pandas==1.3.5;
openpyxl==3.1.2;
xlsWriter==3.0.3

## First Launch
If there is no excel file with keywords and replacements in the folder with Replacer.py, the file will be created automatically with an example of filling and instruction, a warning window with rules will also appear.
Click 'OK' button to close.
Open the folder where Replacer.py is located and find the autoreplacement.xlsx file.
Filled fields (excepted A1 and B1) are for reference only and can be changed (or used).
Do not change A1 (Keyword) and B1 (Replacement) cells
Keywords are entered in column A, replacements - in column B opposite the key word. All characters except spaces are allowed in column A, all characters including spaces are allowed in column B

## GUI
Close button is available only in inactive state.
The script is activated and deactivated using the 'START'/'STOP' button, which changes accordingly.
Values from autoreplacement.xlsx are exported to the replacement lis.
Autoreplacement from the corresponding line can be copied to the clipboard by double-clicking the left mouse button. This option is valid in active and inactive state. 

## Keyboard replacement
To apply the keyboard replacement, the program must be in active mode. 
Press the keyboard shortcut shift+` that implements the symbol '~'. Then * will be automatically printed, that means replacer works. 
Enter the key combination in sequence that creates the keyword, then press space.
You can cancel the start of the replacement if you delete the automatic * using Backspace. Also, can be deleted accidental chars after *.

## An example:

Key - hi, replacement - Hi, how are you?

press: Shift+` --> ~*

press: hi -->  ~*hi

Press the space key --> Hi, how are you?

These same actions before each keyword.




