# VBAScripts
Collection of useful VBA scripts. This could be useful during a pentest for a situation where you have remote desktop access to a pc with MS Office installed, but for whatever reason you can't open a command prompt or run any arbitrary code.

# Usage
 - Download the required script(s)
 - Copy the entire .bas file to the clipboard
 - Create a new macro in MS Office with the same name as the downloaded script
 - Replace all code in this new macro with the code copied to the clipboard
 - Run the macro

# Scripts

### [ExecCmd](https://github.com/joseph-dillon/VBAScripts/tree/master/ExecCmd)
Execute a cmd command and return the output in a message box. Type `exit` to exit.

### [ExecPython](https://github.com/joseph-dillon/VBAScripts/tree/master/ExecPython)
Execute python code (all code has to be formatted onto a single line). If Python isn't installed on the PC, a portable installation of Python can be used instead as the location of `python.exe` can be specified.

### [ReverseShell](https://github.com/joseph-dillon/VBAScripts/tree/master/ReverseShell)
Start a reverse shell (interactive `cmd.exe`). Python **2** is required for this script to work (a portable installation can be used if Python 2 isn't installed on the PC).
