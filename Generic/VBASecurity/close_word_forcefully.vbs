'VB Script listing                                                    -------------------------------------------------------------------
On Error Resume Next

dim objWMIService, colProcess, objProcess, returnValue
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set colProcess = objWMIService.ExecQuery("Select * from Win32_Process Where Name = 'winword.exe'")
For Each objProcess in colProcess
    returnValue = objProcess.Terminate()
Next

WScript.Quit