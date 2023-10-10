'VB Script listing                                                    -------------------------------------------------------------------
On Error Resume Next

dim found, count, wd, msg

found = false
count = 0
Do
    Set wd = Nothing
    Set wd = GetObject(, "Word.Application")
    If Not wd Is Nothing Then
        found = true
        msg = ""
		If Not wd.ActiveDocument.Name = "" Then
			msg = "'" & wd.ActiveDocument.Name & "'."
			msg = "Save & Close or Close Microsoft Word" & chr(13) & msg & chr(13) & chr(13) & "Click OK to continue."
		Else
			msg = "Dead Microsoft Word process found. Please manually kill the Microsoft Word process" & chr(13) & "or restart the computer and try the installation process again." & chr(13) & chr(13) & "Click OK to continue."
		End If
        msg_response = MsgBox(msg, vbOKOnly + vbDefaultButton1, "Microsoft Word Template Update")
    Else
        Exit Do
    End If
    
    count = count + 1
    If count > 20 then
        Exit Do
    End If
Loop
Set wd = Nothing
Set wd = GetObject(, "Word.Application")
If Not wd Is Nothing Then
    wd.Quit(wdDoNotSaveChanges) 'SaveChanges:=wdDoNotSaveChanges
End If

found = false
count = 0
Do
    Set wd = Nothing
    Set wd = GetObject(, "Excel.Application")
    If Not wd Is Nothing Then
        found = true
        msg = ""
		If Not wd.ActiveWorkbook.Name = "" Then
			msg = "'" & wd.ActiveWorkbook.Name & "'."
			msg = "Save & Close or Close Microsoft Excel" & chr(13) & msg & chr(13) & chr(13) & "Click OK to continue."
		Else
			msg = "Dead Microsoft Excel process found. Please manually kill the Microsoft Excel process" & chr(13) & "or restart the computer and try the installation process again." & chr(13) & chr(13) & "Click OK to continue."
		End If
        msg_response = MsgBox(msg, vbOKOnly + vbDefaultButton1, "Microsoft Word Template Update")
    Else
        Exit Do
    End If
    
    count = count + 1
    If count > 20 then
        Exit Do
    End If
Loop
Set wd = Nothing
Set wd = GetObject(, "Excel.Application")
If Not wd Is Nothing Then
    wd.ActiveWorkbook.Saved = true
    wd.Quit
End If

WScript.Quit