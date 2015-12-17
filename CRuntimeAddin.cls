VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRuntimeAddin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' =====================================================================================================================
'
' Project Status and Activity Report Distribution Add-in
' Research Accounts, Finance Division, University of Oxford
' v1.4 - James Boon - Jan 2015
'
' ---------------------------------------------------------------------------------------------------------------------
'
' CRuntimeAddin
' -------------
'
' Module to contain code for class to hold runtime addins
'
' =====================================================================================================================

Option Explicit
'Option Private Module - not permitted in object module

Private Const msMODULE As String = "CRuntimeAddin"

' =====================================================================================================================
' Module variable declarations
' =====================================================================================================================

Private msAddinName As String
Private mbAddinStatus As Boolean

' =====================================================================================================================
' Let Properties
' =====================================================================================================================

Public Property Let Name(sValue As String)
    msAddinName = sValue
End Property

Public Property Let Status(bValue As Boolean)
    mbAddinStatus = bValue
End Property

' =====================================================================================================================
' Get Properties
' =====================================================================================================================

Public Property Get Name() As String
    Name = msAddinName
End Property

Public Property Get Status() As Boolean
    Status = mbAddinStatus
End Property
