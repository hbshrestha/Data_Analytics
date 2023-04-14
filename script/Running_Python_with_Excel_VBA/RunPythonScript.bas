Attribute VB_Name = "RunPythonScript"
Sub RunPythonScript()

Dim objShell As Object
Dim PythonExePath As String, PythonScriptPath As String
ActiveWorkbook.Save

'Enter into the path of given workbook
ChDir Application.ThisWorkbook.Path

    Set objShell = VBA.CreateObject("Wscript.Shell")
    
    'Goto cmd. Type where python to get this path. Note that there are three quotes below.
    PythonExePath = """C:\Users\HShrestha\AppData\Local\Microsoft\WindowsApps\python.exe"""
    
    'Get the path of the file.
    PythonScriptPath = Application.ThisWorkbook.Path & "\python_script.py"
    
    objShell.Run PythonExePath & PythonScriptPath


End Sub
