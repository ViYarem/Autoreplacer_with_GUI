# Replacer.py
# Overview
The script does the autoreplacement of the keywords are entered from keyboard. The keywords and words/phrases for replacement are being field in excel file. If you often fill documents with the same phrases and want to speed up this process, then this script may be usefull for you. The script launches a user-friendly GUI which includes small instruction for using:

![image](https://github.com/ViYarem/autoreplacer/assets/68001529/73077dfb-5b16-4814-8f20-85576d112613)

# What You Will Need
Python>=3.0;
pynput==1.4.5;
tk==0.1.0;
PyAutoGUI==0.9.53;
pandas==1.3.5;
openpyxl==3.1.2;
xlsWriter==3.0.3;

# First Launch
If there is no excel file with keywords and replacements in the folder with Replacer.py, the file will be created automatically with an example of filling and instruction, a warning window with rules will also appear:
![image](https://github.com/ViYarem/autoreplacer/assets/68001529/a5725aa4-6349-44c8-8fc6-17cc4dee23b7)
Click 'OK' button to close.
Open the folder where Replacer.py is located and find the autoreplacement.xlsx file. It will be filled in as follows:
![image](https://github.com/ViYarem/autoreplacer/assets/68001529/f31ed1cd-dce2-4649-ad3d-8901de33d377)
Filled fields except (A1 and B1) are for reference only and can be changed (or used).
! Do not change A1 (Keyword) and B1 (Replacement) cells
! Keywords are entered in column A, replacements - in column B opposite the key word. All characters except spaces are allowed in column A, all characters including spaces are allowed in column B

# GUI
At the next start, a window will appear:
Indicators: ![image](https://github.com/ViYarem/autoreplacer/assets/68001529/2df46063-3d1d-4fea-a197-b46265fe884c) indicate that the replacer is active.
Close button ![image](https://github.com/ViYarem/autoreplacer/assets/68001529/76a24f8f-cad8-492a-9d35-1731ee47e57a) is available only in inactive state: ![image](https://github.com/ViYarem/autoreplacer/assets/68001529/da461802-1585-4e42-8149-60ea29da8a08)
The script is activated and deactivated using the 'START'/'STOP' button, which changes accordingly.
The window can be moved by clicking on this area: ![image](https://github.com/ViYarem/autoreplacer/assets/68001529/7a5cec10-232b-4e96-925b-17bfb659f8ad) 
Values from autoreplacement.xlsx are exported to the replacement list ![image](https://github.com/ViYarem/autoreplacer/assets/68001529/9a8d46ff-67fe-4310-9092-b1eed8e93c9b)
Autoreplacement from the corresponding line can be copied to the clipboard by double-clicking the left mouse button. This option is valid in active and inactive state. 

# Keyboard replacement
!To apply the keyboard replacement, the program must be in active mode: ![image](https://github.com/ViYarem/autoreplacer/assets/68001529/271fd480-0604-4bb0-b552-25fba688cd3f)
!Be careful: for replacement, it is not the word displayed in the text field, but the replacement of keys entered on the keyboard. Check the keyboard layout before using the program, as the program simulates your keyboard strokes in the current layout.
Press the keyboard shortcut shift+` at the same time ![image](https://github.com/ViYarem/autoreplacer/assets/68001529/f57c5e56-f023-491c-bfa5-17098470030b) which implements the symbol ~ . After it, * is automatically printed, which means that the key is ready to be entered. 
Enter the key combination in sequence that creates the keyword. 
After - press space.
You can cancel the start of the replacement if you immediately delete the automatic * using Backspace (also restore). You can also use Backspace to delete letters at the end of what was written, which were printed by accident and are not included in the keyword (delete, etc. do not work). 
Backspace removes characters from the end of the sequence of entered characters (it does not depend on the position of the cursor and what the test field displays).

# An example:
Key - hi, replacement - Hi, how are you?
press: Shift+` --> ~*
press: hi -->  ~*hi
Press the space key --> Hi, how are you?
These same actions before each keyword.


