VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CRuntimeSettings"
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
' CRuntimeSettings
' ----------------
'
' Module for Runtime Settings class
' Takes initial values for applictaion.calculation and installed status of all add-ins
' so that we can reinstate these at the end of runtime where required
'
' =====================================================================================================================

Option Explicit
'Option Private Module - not permitted in object module

Private Const msMODULE As String = "CRuntimeSettings"

' =====================================================================================================================
' Module variable declarations
' =====================================================================================================================

Private mbCalculationWasAlreadyManual As Boolean
Private mcolAddins As Collection

' =====================================================================================================================
' Let Properties
' =====================================================================================================================

Public Property Let ManualCalculation(bValue As Boolean)
    mbCalculationWasAlreadyManual = bValue
End Property

Public Property Let Addins(colValue As Collection)
    Set mcolAddins = colValue
End Property

' =====================================================================================================================
' Get Properties
' =====================================================================================================================

Public Property Get ManualCalculation() As Boolean
    ManualCalculation = mbCalculationWasAlreadyManual
End Property

Public Property Get Addins() As Collection
    Set Addins = mcolAddins
End Property

' =====================================================================================================================
' Subroutines
' =====================================================================================================================

Private Sub class_initialize()
    If Not bInitialise Then Err.Raise Number:=glHANDLED_ERROR
End Sub

Private Sub class_terminate()
    On Error Resume Next
    Set mcolAddins = Nothing
End Sub

' =====================================================================================================================
' Functions
' =====================================================================================================================

Private Function bInitialise() As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to collect titles and installed status of all addins
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    class_initialize
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 16 Dec 2014   James Boon      1.2         Created
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bInitialise"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim wkb As Excel.Workbook

' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    ' get manual calculation
    If Application.ActiveWorkbook Is Nothing Then Set wkb = Workbooks.Add 'prevents error on reading application.calculation
    Me.ManualCalculation = (Application.Calculation = xlCalculationManual)
    
    ' get status of add-ins
    If Not bCollectAddins Then Err.Raise Number:=glHANDLED_ERROR
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    On Error Resume Next
    wkb.Close savechanges:=False
    Set wkb = Nothing
    bInitialise = bReturn
    Exit Function

ErrorHandler:
    bReturn = False
    If bCentralErrorHandler(msMODULE, sSource) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
    
End Function

' =====================================================================================================================

Private Function bCollectAddins() As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to collect titles and installed status of all addins
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    bInitialise
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 16 Dec 2014   James Boon      1.2         Created
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bCollectAddins"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim vAddIn As Excel.AddIn
    Dim clsRuntimeAddin As CRuntimeAddin

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Set mcolAddins = New Collection
    
    For Each vAddIn In Application.Addins
        Set clsRuntimeAddin = New CRuntimeAddin
        clsRuntimeAddin.Name = vAddIn.Title
        clsRuntimeAddin.Status = vAddIn.Installed
        Me.Addins.Add Item:=clsRuntimeAddin, Key:=clsRuntimeAddin.Name
    Next 'vAddin

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bCollectAddins = bReturn
    On Error Resume Next
    Set vAddIn = Nothing
    Set clsRuntimeAddin = Nothing
    Exit Function

ErrorHandler:
    bReturn = False
    If bCentralErrorHandler(msMODULE, sSource) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If
    
End Function

