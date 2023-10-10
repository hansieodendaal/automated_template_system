'VB Script listing                                                    -------------------------------------------------------------------
On Error Resume Next

Set wshShell = CreateObject( "WScript.Shell" )
dim strComputer
strComputer = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

dim objWMIService
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

dim colAccounts
Set colAccounts = objWMIService.ExecQuery _
    ("Select * From Win32_UserAccount Where Domain = '" & strComputer & "'")

dim objAccount
For Each objAccount in colAccounts
    If Left (objAccount.SID, 6) = "S-1-5-" and Right(objAccount.SID, 4) = "-500" Then
        Wscript.Echo objAccount.Name
    End If
    Wscript.Echo objAccount.SID & ": " & objAccount.Name
Next

