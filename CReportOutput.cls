VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CReportOutput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msKey As String
Private oExceptions As Object
Private msExceptions As String

Public Property Let Key(ByVal sValue As String)
    msKey = sValue
End Property

Public Property Let ExceptionsLabel(ByVal oValue As Object)
    Set oExceptions = oValue
End Property

Public Property Let ExceptionsSetting(ByVal sValue As String)
    msExceptions = sValue
End Property


Public Property Get Key() As String
    Key = msKey
End Property

Public Property Get ExceptionsLabel() As Object
    Set ExceptionsLabel = oExceptions
End Property

Public Property Get ExceptionsSetting() As String
    ExceptionsSetting = msExceptions
End Property

