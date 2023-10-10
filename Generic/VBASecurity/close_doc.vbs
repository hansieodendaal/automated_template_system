'VB Script listing                                                    -------------------------------------------------------------------
On Error Resume Next

WScript.Sleep 7000

dim found, count, wd
found = false
count = 0
Do
    WScript.Sleep 250
    Set wd = Nothing
    Set wd = GetObject(, "Word.Application")
    If Not wd Is Nothing Then
        wd.Quit(wdDoNotSaveChanges) 'SaveChanges:=wdDoNotSaveChanges
    End If
    
    count = count + 1
    If count > 20 then
        Exit Do
    End If
Loop

WScript.Quit