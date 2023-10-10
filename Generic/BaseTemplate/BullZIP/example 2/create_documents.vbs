Const number_of_documents = 10

Rem -- Get current path
Set fso = CreateObject("Scripting.FileSystemObject")
currentdir = fso.GetAbsolutePathName(".")

Rem -- Create test documents
For i = 1 To number_of_documents
	filecontent = "This is test document #" & i & "." & vbcrlf & "This file was created at " & now
	filename = currentdir & "\in\" & i & ".txt"
	Set objFile = fso.CreateTextFile(filename, true, false)
	objFile.WriteLine filecontent
Next

Wscript.Echo number_of_documents & " documents were created in the 'in' folder."