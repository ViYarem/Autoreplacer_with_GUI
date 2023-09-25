# Replacer.py
## Overview
The script does the autoreplacement of the keywords are entered from keyboard. The keywords and words/phrases for replacement are being field in excel file. If you often fill documents with the same phrases and want to speed up this process, then this script may be usefull for you. The script launches a user-friendly GUI which includes small instruction for using.

# What You Will Need
Python>=3.0;
pynput==1.4.5;
tk==0.1.0;
PyAutoGUI==0.9.53;
pandas==1.3.5;
openpyxl==3.1.2;
xlsWriter==3.0.3;

## First Launch
If there is no excel file with keywords and replacements in the folder with Replacer.py, the file will be created automatically with an example of filling and instruction, a warning window with rules will also appear
Click 'OK' button to close.
Open the folder where Replacer.py is located and find the autoreplacement.xlsx file.
Filled fields except (A1 and B1) are for reference only and can be changed (or used).
Do not change A1 (Keyword) and B1 (Replacement) cells
Keywords are entered in column A, replacements - in column B opposite the key word. All characters except spaces are allowed in column A, all characters including spaces are allowed in column B

## GUI
Close button is available only in inactive state.
The script is activated and deactivated using the 'START'/'STOP' button, which changes accordingly.
Values from autoreplacement.xlsx are exported to the replacement lis.
Autoreplacement from the corresponding line can be copied to the clipboard by double-clicking the left mouse button. This option is valid in active and inactive state. 

## Keyboard replacement
To apply the keyboard replacement, the program must be in active mode. 
Be careful: for replacement, it is not the word displayed in the text field, but the replacement of keys entered on the keyboard. Check the keyboard layout before using the program, as the program simulates your keyboard strokes in the current layout.
Press the keyboard shortcut shift+` at the same time which implements the symbol ~ . After it, * is automatically printed, which means that the key is ready to be entered. 
Enter the key combination in sequence that creates the keyword. After - press space.
You can cancel the start of the replacement if you immediately delete the automatic * using Backspace (also restore). You can also use Backspace to delete letters at the end of what was written, which were printed by accident and are not included in the keyword (delete, etc. do not work). 
Backspace removes characters from the end of the sequence of entered characters (it does not depend on the position of the cursor and what the test field displays).

## An example:

Key - hi, replacement - Hi, how are you?

press: Shift+` --> ~*

press: hi -->  ~*hi

Press the space key --> Hi, how are you?

These same actions before each keyword.




