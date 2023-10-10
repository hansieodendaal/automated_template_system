VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "vbaClipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function EnumClipboardFormats Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function RegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
'Note that we do not use the GetClipboardDataA declaration
'Public Declare Function GetClipboardData Lib "user32" Alias "GetClipboardDataA" (ByVal wFormat As Long) As Long
Private Declare Function GetClipBoardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
'NOTE: the lstrCpy declaration you get from the VB6 API Viewer is WRONG. It's version is this:
'Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
'The correct version for (at least for Windows 7 / 64-bit is this:
Private Declare Function lstrCpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32" (ByVal wFormat As Long) As Long
Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long

Private Const GMEM_MOVEABLE = &H2
Private Const GMEM_ZEROINIT = &H40
Private Const GHND = (GMEM_MOVEABLE Or GMEM_ZEROINIT) 'Use for hwnd
Private Const NAME_MAX_LENGTH = 1024

Private Const APINULL = 0
Private Const CF_TEXT = 1 'Text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data. Use this format for ANSI text.
Private Const CF_BITMAP = 2 'A handle to a bitmap (HBITMAP).
Private Const CF_METAFILEPICT = 3 'Handle to a metafile picture format as defined by the METAFILEPICT structure. When passing a CF_METAFILEPICT handle by means of DDE, the application responsible for deleting hMem should also free the metafile referred to by the CF_METAFILEPICT handle.
Private Const CF_SYLK = 4 'Microsoft Symbolic Link (SYLK) format.
Private Const CF_TIFF = 6 'Tagged-image file format.
Private Const CF_DIF = 5 'Software Arts' Data Interchange Format.
Private Const CF_OEMTEXT = 7 'Text format containing characters in the OEM character set. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
Private Const CF_DIB = 8 'A memory object containing a BITMAPINFO structure followed by the bitmap bits.
Private Const CF_PALETTE = 9 'Handle to a color palette. Whenever an application places data in the clipboard that depends on or assumes a color palette, it should place the palette on the clipboard as well.
Private Const CF_PENDATA = 10 'Data for the pen extensions to the Microsoft Windows for Pen Computing.
Private Const CF_RIFF = 11 'Represents audio data more complex than can be represented in a CF_WAVE standard wave format.
Private Const CF_WAVE = 12 'Represents audio data in one of the standard wave formats, such as 11 kHz or 22 kHz PCM.
Private Const CF_UNICODETEXT = 13 'Unicode text format. Each line ends with a carriage return/linefeed (CR-LF) combination. A null character signals the end of the data.
Private Const CF_ENHMETAFILE = 14 'A handle to an enhanced metafile (HENHMETAFILE).
Private Const CF_HDROP = 15  'A handle to type HDROP that identifies a list of files. An application can retrieve information about the files by passing the handle to the DragQueryFile function.
Private Const CF_LOCALE = 16 'The data is a handle to the locale identifier associated with text in the clipboard. When you close the clipboard, if it contains CF_TEXT data but no CF_LOCALE data, the system automatically sets the CF_LOCALE format to the current input language. You can use the CF_LOCALE format to associate a different locale with the clipboard text.
Private Const CF_DIBV5 = 17 'A memory object containing a BITMAPV5HEADER structure followed by the bitmap color space information and the bitmap bits.
Private Const CF_DSPBITMAP = &H82 'Bitmap display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in bitmap format in lieu of the privately formatted data.
Private Const CF_DSPENHMETAFILE = &H8E  'Enhanced metafile display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in enhanced metafile format in lieu of the privately formatted data.
Private Const CF_DSPMETAFILEPICT = &H83 'Metafile-picture display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in metafile-picture format in lieu of the privately formatted data.
Private Const CF_DSPTEXT = &H81 'Text display format associated with a private format. The hMem parameter must be a handle to data that can be displayed in text format in lieu of the privately formatted data.
Private Const CF_GDIOBJFIRST = &H300 'Start of a range of integer values for application-defined GDI object clipboard formats. The end of the range is CF_GDIOBJLAST.
Private Const CF_GDIOBJLAST = &H3FF 'See CF_GDIOBJFIRST.
Private Const CF_OWNERDISPLAY = &H80 'Owner-display format. The clipboard owner must display and update the clipboard viewer window, and receive the WM_ASKCBFORMATNAME, WM_HSCROLLCLIPBOARD, WM_PAINTCLIPBOARD, WM_SIZECLIPBOARD, and WM_VSCROLLCLIPBOARD messages. The hMem parameter must be NULL.
Private Const CF_PRIVATEFIRST = &H200 'Start of a range of integer values for private clipboard formats. The range ends with CF_PRIVATELAST. Handles associated with private clipboard formats are not freed automatically; the clipboard owner must free such handles, typically in response to the WM_DESTROYCLIPBOARD message.
Private Const CF_PRIVATELAST = &H2FF 'See CF_PRIVATEFIRST.

Public Property Get ClipboardFormatsAvailable() As Collection
 On Error GoTo ErrorHandler
 Dim thisClipboardFormat As Long
 Dim returnStringLength As Long
 Dim myCFAvailable As New Collection
 Dim clipBoardFormatName As String
 Dim ClipboardFormat As ClipboardFormat
 Dim success As Boolean
 success = OpenClipboard(0)
 If success Then
  thisClipboardFormat = 0
  thisClipboardFormat = EnumClipboardFormats(thisClipboardFormat)
  While thisClipboardFormat <> 0
   Set ClipboardFormat = New ClipboardFormat
   clipBoardFormatName = String$(NAME_MAX_LENGTH, vbNullChar)
   returnStringLength = GetClipboardFormatName(thisClipboardFormat, _
    clipBoardFormatName, Len(clipBoardFormatName))
   clipBoardFormatName = TrimNull(clipBoardFormatName)
   If clipBoardFormatName = "" Then
    clipBoardFormatName = BuiltInClipboardFormatName(thisClipboardFormat)
   End If
   ClipboardFormat.Name = clipBoardFormatName
   ClipboardFormat.Number = thisClipboardFormat
   myCFAvailable.Add ClipboardFormat, ClipboardFormat.Name
   thisClipboardFormat = EnumClipboardFormats(thisClipboardFormat)
  Wend
  Set ClipboardFormatsAvailable = myCFAvailable
  CloseClipboard
 Else
  Set ClipboardFormatsAvailable = Nothing
 End If
 
 Exit Property
ErrorHandler:
 On Error Resume Next
 CloseClipboard
End Property

Public Function GetClipboardText(ByVal aClipboardFormatNumber As Long) As String
 'Do not handle errors - let them bubble up
 Dim wLen As Integer
 Dim hMemory As Long
 Dim hMyMemory As Long

 Dim lpMemory As Long
 Dim lpMyMemory As Long

 Dim RetVal As Variant
 Dim haveMemoryLocked As Boolean
 Dim wClipAvail As Integer
 Dim szText As String
 Dim wSize As Long
 
 
 Dim clipBoardText As String
 clipBoardText = ""
 
 'Before accessing the clipboard, find out if the requested format is available
 If IsClipboardFormatAvailable(aClipboardFormatNumber) = APINULL Then
  Err.Raise vbObjectError + 1, "vbaClipboard", "Requested clipboard format number " & aClipboardFormatNumber & " is not available on the clipboard."
  Exit Function
 End If
 
 
 Dim success As Boolean
 success = OpenClipboard(0)
 If success Then
  'Get a handle to a memory structure containing the clipboard data in the requested format
  hMemory = GetClipBoardData(aClipboardFormatNumber)
  CloseClipboard
  'If the handle is null, something went wrong
  If hMemory = APINULL Then
   'Throw an error
   Err.Raise vbObjectError + 1, "vbaClipboard", "Unable to retrieve data from the Clipboard."
  End If
  'The handle is good. How much data came back?
  wSize = GlobalSize(hMemory)
  'Fill our destination string with nulls
  clipBoardText = Space(wSize)

  'Lock the memory
  'Get a pointer to the locked memory area
  lpMemory = GlobalLock(hMemory)
  If lpMemory = APINULL Then
   'CloseClipboard
   Err.Raise vbObjectError + 1, "vbaClipboard", "Unable to lock clipboard memory."
  End If
 
  ' Copy the locked memory into our string
  RetVal = lstrCpy(clipBoardText, lpMemory)
  
  'Unlock memory
  GlobalUnlock hMemory
  
  ' Get rid of trailing stuff.
  clipBoardText = Trim(clipBoardText)
  GetClipboardText = TrimNull(clipBoardText)
 Else
  Err.Raise vbObjectError + 1, "vbaClipboard", "Unable to open Clipboard. Perhaps some other application is using it."
 End If
End Function

Public Sub SetClipboardText(ByVal aText As String, ByVal aClipboardFormatName As String)
  
 Dim wLen As Integer
 Dim hMemory As Long
 Dim lpMemory As Long
 Dim RetVal As Variant
 Dim memoryIsLocked As Boolean
 Dim memoryIsAllocated As Boolean
 Dim clipBoardIsOpen As Boolean
 
 memoryIsAllocated = False
 memoryIsLocked = False
 clipBoardIsOpen = False
 
 On Error GoTo ErrorHandler
 
 ' Get the length, including one extra for a CHR$(0) at the end.
 wLen = Len(aText) + 1
 'Add a null to the end
 aText = aText & Chr$(0)
 'Allocate some memory
 hMemory = GlobalAlloc(GHND, wLen + 1)
 If hMemory = APINULL Then
  Err.Raise vbObjectError + 1001, "vbaClipboard", "Unable to allocate memory."
 Else
  memoryIsAllocated = True
 End If
 
 lpMemory = GlobalLock(hMemory)
 If lpMemory = APINULL Then
  'Throw an error
  Err.Raise vbObjectError + 1001, "vbaClipboard", "Unable to lock memory."
 Else
  memoryIsLocked = True
 End If

 ' Copy our string into the locked memory.
 RetVal = lstrCpy(lpMemory, aText)
 ' Don't send clipboard locked memory.
 RetVal = GlobalUnlock(hMemory)
 'If the preceding throws an error, it will be handled in ErrorHandler
 memoryIsLocked = True
 If OpenClipboard(0&) = APINULL Then
  Err.Raise vbObjectError + 1, "vbaClipboard", "Unable to open Clipboard. Perhaps some other application is using it."
 Else
  clipBoardIsOpen = True
 End If
 
 
 'Is the requested format one of the Windows built-in formats?
 Dim i As Integer
 Dim thisClipboardFormatNumber As Long
 thisClipboardFormatNumber = BuiltInClipboardFormatNumber(aClipboardFormatName)
 If thisClipboardFormatNumber = 0 Then
  'Nope. Register the format
  On Error Resume Next
  thisClipboardFormatNumber = RegisterClipboardFormat(aClipboardFormatName)
  If Err.Number <> 0 Then
   Err.Raise vbObjectError + 1, "vbaClipboard", "Unable to register clipboard format: " & aClipboardFormatName & _
    ". Error message: " & Err.description
  End If
  
  On Error GoTo ErrorHandler
  If thisClipboardFormatNumber = 0 Then
   Err.Raise vbObjectError + 1, "vbaClipboard", "Unable to register clipboard format: " & aClipboardFormatName
  End If
 End If
 
 'Empty the clipboard
 If EmptyClipboard() = APINULL Then
  Err.Raise vbObjectError + 1, "vbaClipboard", "Unable to empty the clipboard."
 End If
 
 If SetClipboardData(thisClipboardFormatNumber, hMemory) = APINULL Then
  Err.Raise vbObjectError + 1, "vbaClipboard", "Unable to set the clipboard data."
 End If
 
 CloseClipboard
 GlobalFree hMemory
 
 Exit Sub
ErrorHandler:
 Dim description As String
 description = Err.description
 On Error Resume Next
 If memoryIsLocked Then GlobalUnlock hMemory
 If memoryIsAllocated Then GlobalFree hMemory
 If clipBoardIsOpen Then CloseClipboard
 On Error GoTo 0
 Err.Raise vbObjectError + 1, "vbaClipboard", description
End Sub


Private Function TrimNull(ByVal aString As String) As String
 TrimNull = Left(aString, _
   InStr(1, aString, vbNullChar) - 1)
End Function

Private Function BuiltInClipboardFormatNumber(ByVal aClipboardFormatName As String) As Long
 Dim result As Long
 Select Case UCase(aClipboardFormatName)
  Case "CF_TEXT"
   result = 1
  Case "CF_BITMAP"
   result = 2
  Case "CF_METAFILEPICT"
   result = 3
  Case "CF_SYLK"
   result = 4
  Case "CF_DIF"
   result = 5
  Case "CF_TIFF"
   result = 6
  Case "CF_OEMTEXT"
   result = 7
  Case "CF_DIB"
   result = 8
  Case "CF_PALETTE"
   result = 9
  Case "CF_PENDATA"
   result = 10
  Case "CF_RIFF"
   result = 11
  Case "CF_WAVE"
   result = 12
  Case "CF_UNICODETEXT"
   result = 13
  Case "CF_ENHMETAFILE"
   result = 14
  Case "CF_HDROP"
   result = 15
  Case "CF_LOCALE"
   result = 16
  Case "CF_DIBV5"
   result = 17
  Case "CF_DSPBITMAP"
   result = &H82
  Case "CF_DSPENHMETAFILE"
   result = &H8E
  Case "CF_DSPMETAFILEPICT"
   result = &H83
  Case "CF_DSPTEXT"
   result = &H81
  Case "CF_GDIOBJFIRST"
   result = &H300
  Case "CF_GDIOBJLAST"
   result = &H3FF
  Case "CF_OWNERDISPLAY"
   result = &H80
  Case "CF_PRIVATEFIRST"
   result = &H200
  Case "CF_PRIVATELAST"
   result = &H2FF
  Case Else
   result = 0
 End Select
  BuiltInClipboardFormatNumber = result
End Function

Private Function BuiltInClipboardFormatName(ByVal aIndex As Integer) As String
 Dim n As String
 Select Case aIndex
  Case 1
   n = "CF_TEXT"
  Case 2
   n = "CF_BITMAP"
  Case 3
   n = "CF_METAFILEPICT"
  Case 4
   n = "CF_SYLK"
  Case 5
   n = "CF_DIF"
  Case 6
   n = "CF_TIFF"
  Case 7
   n = "CF_OEMTEXT"
  Case 8
   n = "CF_DIB"
  Case 9
   n = "CF_PALETTE"
  Case 10
   n = "CF_PENDATA"
  Case 11
   n = "CF_RIFF"
  Case 12
   n = "CF_WAVE"
  Case 13
   n = "CF_UNICODETEXT"
  Case 14
   n = "CF_ENHMETAFILE"
  Case 15
   n = "CF_HDROP"
  Case 16
   n = "CF_LOCALE"
  Case 17
   n = "CF_DIBV5"
  Case &H82
   n = "CF_DSPBITMAP"
  Case &H8E
   n = "CF_DSPENHMETAFILE"
  Case &H83
   n = "CF_DSPMETAFILEPICT"
  Case &H81
   n = "CF_DSPTEXT"
  Case &H300
   n = "CF_GDIOBJFIRST"
  Case &H3FF
   n = "CF_GDIOBJLAST"
  Case &H80
   n = "CF_OWNERDISPLAY"
  Case &H200
   n = "CF_PRIVATEFIRST"
  Case &H2FF
   n = "CF_PRIVATELAST"
 End Select
 BuiltInClipboardFormatName = n
End Function

