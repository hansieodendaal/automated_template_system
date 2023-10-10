Attribute VB_Name = "ClipBoardTest"
Option Explicit

Sub testClipboard1()
    'This routine tests the vbaClipboard object.
    'Before running this, copy some text from Word. This will place Rich Text Format data
    'on the clipboard. The test will preserve the RTF data, then use the clipboard
    'to manipulate some plain text ("CF_TEXT"). Finally, the test will put the
    'RTF data back on the clipboard. When the test is finished, you should be able
    'to go back into Word and hit Ctrl+V and paste your original copied text (with formatting).
    
    'Instantiate a vbaClipboard object
    Dim myClipboard As New vbaClipboard
    
    'The ClipboardFormat class encapsulates a clipboard format number and a name
    Dim ClipboardFormat As ClipboardFormat
    
    'Handle errors below
    On Error GoTo ErrorHandler
    
    'Show the currently available formats
    'The ClipboardFormatsAvailable property returns a collection of ClipboardFormat objects
    'representing all formats currently available on the clipboard.
    
    Debug.Print "===================================================================="
    
    For Each ClipboardFormat In myClipboard.ClipboardFormatsAvailable
        Debug.Print ClipboardFormat.Number, ClipboardFormat.Name
    Next ClipboardFormat
        
    'Preserve the RTF currently on the clipboard (you did copy some, right?)
    Dim oldRTF As String
    'Get the format number value for Rich Text Format
    Dim richTextFormatNumber As Long
    On Error Resume Next
    richTextFormatNumber = myClipboard.ClipboardFormatsAvailable("Rich Text Format").Number
    If Err.Number <> 0 Then
        On Error GoTo ErrorHandler
        Err.Raise vbObjectError + 1, , "The clipboard does not have any Rich Text Format data."
    End If
    On Error GoTo ErrorHandler
    
    'Get the RTF data from the clipboard
    oldRTF = myClipboard.GetClipboardText(richTextFormatNumber)
    'Debug.Print oldRTF
    
    'Use the clipboard for something else
    Dim s As String
    s = "Hello, world!"
    myClipboard.SetClipboardText s, "CF_TEXT"
    
    'Get it back again
    Debug.Print myClipboard.GetClipboardText(1)
    
    'Show the currently available formats
    Debug.Print "===================================================================="
    For Each ClipboardFormat In myClipboard.ClipboardFormatsAvailable
        Debug.Print ClipboardFormat.Number, ClipboardFormat.Name
    Next ClipboardFormat
    
    'Now put back the RTF
    myClipboard.SetClipboardText oldRTF, "Rich Text Format"
    
    'Show the currently available formats
    Debug.Print "===================================================================="
    For Each ClipboardFormat In myClipboard.ClipboardFormatsAvailable
        Debug.Print ClipboardFormat.Number, ClipboardFormat.Name
    Next ClipboardFormat
    'You can now paste back into Word, and you'll get whatever text you selected
    Exit Sub
ErrorHandler:
    MsgBox Err.description
End Sub

Sub testClipboard2()
    'This tests stuffs some formatted text (RTF) onto the clipboard. Run the test, then
    'go into word and hit Ctrl+V to paste it in.
    Dim myClipboard As New vbaClipboard
    Dim sRTF As String
    sRTF = "{\rtf1\ansi\ansicpg1252\deff0\deftab720{\fonttbl" & _
           "{\f0\fswiss MS Sans Serif;}{\f1\froman\fcharset2 Symbol;}" & _
           "{\f2\froman\fprq2 Times New Roman;}}" & _
           "{\colortbl\red0\green0\blue0;\red255\green0\blue0;}" & _
           "\deflang1033\horzdoc{\*\fchars }{\*\lchars }" & _
           "\pard\plain\f2\fs24 This is some \plain\f2\fs24\cf1" & _
           "formatted\plain\f2\fs24  text.\par }"
              
    myClipboard.SetClipboardText sRTF, "Rich Text Format"
End Sub



