'VB Script listing                                                    -------------------------------------------------------------------

dim wshShellm, wshUserEnv, strItem
Set wshShell = CreateObject( "WScript.Shell" )
Set wshUserEnv = wshShell.Environment( "Process" )

dim oWord, oControl
Set oWord = CreateObject("Word.Application")
oWord.Visible = True
oWord.Documents.Open(wshUserEnv("targetDocument"))

oWord.ActiveDocument.BuiltInDocumentProperties("Comments").value ="<//([{**** Word 2016 Fields Update Fix ****}])\\>"

dim objExec
Set objExec = wshShell.Exec("ping localhost -n 3")
objExec.StdOut.ReadAll

WScript.Echo "With the '" & wshUserEnv("Company") & "' ribbon, run 'Update Document -> Update Details', then close the document." & chr(13) & chr(13) & "BuiltInDocumentProperty 'Comments' =" & chr(13) &  "   '" & oWord.ActiveDocument.BuiltInDocumentProperties("Comments").value & "'"

Set wshUserEnv = Nothing
Set wshShell   = Nothing

WScript.Quit