VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClipboardFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mNumber As Long
Private mName As String

Public Property Get Number() As Long
  Number = mNumber
End Property

Public Property Let Number(ByVal value As Long)
  mNumber = value
End Property

Public Property Get Name() As String
  Name = mName
End Property

Public Property Let Name(ByVal value As String)
  mName = value
End Property
 

