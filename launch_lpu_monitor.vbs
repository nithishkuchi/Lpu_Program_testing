' launch_lpu_monitor.vbs
' Double-click this file to open the LPU Monitor tool -- no terminal, no
' commands. Works no matter where the project folder is placed, as long as
' this .vbs file stays in the same folder as app.py.

Set fso = CreateObject("Scripting.FileSystemObject")
scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

Set WshShell = CreateObject("WScript.Shell")
WshShell.CurrentDirectory = scriptDir
WshShell.Run "cmd /c python_embedded\python.exe -m streamlit run app.py", 0, False