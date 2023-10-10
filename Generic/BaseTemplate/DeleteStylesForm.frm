VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DeleteStylesForm 
   Caption         =   "Insert new section..."
   ClientHeight    =   1950
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4545
   OleObjectBlob   =   "DeleteStylesForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DeleteStylesForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
    cancelInsertSection = True
    Unload DeleteStylesForm
End Sub

Private Sub ContinueButton_Click()
    cancelInsertSection = False
    endOfDoc = EndOfDocOption.value
    Unload DeleteStylesForm
End Sub

Private Sub CurrentPositionOption_Click()
    If CurrentPositionOption.value = True Then
        CurrentPositionOption.font.Bold = True
        EndOfDocOption.font.Bold = False
        If EndOfDocOption.value = True Then
            EndOfDocOption.value = False
        End If
    End If
End Sub

Private Sub EndOfDocOption_Click()
    If EndOfDocOption.value = True Then
        EndOfDocOption.font.Bold = True
        CurrentPositionOption.font.Bold = False
        If CurrentPositionOption.value = True Then
            CurrentPositionOption.value = False
        End If
    End If
End Sub

Private Sub FooterCheckBox_Click()
    cleanFooter = FooterCheckBox.value
End Sub

Private Sub HeaderCheckBox_Click()
    cleanHeader = HeaderCheckBox.value
End Sub

Private Sub UserForm_Activate()
    If insertLandscapeSection = True Then
        Caption = "Insert new landscape section..."
    Else
        Caption = "Insert new portrait section..."
    End If
    CurrentPositionOption.value = True
    CurrentPositionOption.font.Bold = True
    If Selection.Range.text <> "" Then
        CurrentPositionOption.Caption = "... using selected text"
    Else
        CurrentPositionOption.Caption = "... at cursor position"
    End If
    EndOfDocOption.value = False
    endOfDoc = False
    HeaderCheckBox.value = cleanHeader
    FooterCheckBox.value = cleanFooter
End Sub
