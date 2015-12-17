VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private WithEvents mcbButton As CommandButton
Attribute mcbButton.VB_VarHelpID = -1
Private msKey As String
Private mlButtonType As glButtonType
Private mclsReport As CReport

Private Const msMODULE As String = "CListener"

Public Property Let Button(ByVal cbButton As CommandButton)
    Set mcbButton = cbButton
End Property

Public Property Let Key(ByVal sValue As String)
    msKey = sValue
End Property

Public Property Let ButtonType(ByVal lValue As glButtonType)
    mlButtonType = lValue
End Property

Public Property Let Report(ByVal oReport As CReport)
    Set mclsReport = oReport
End Property

Public Property Get Button() As CommandButton
    Set Button = mcbButton
End Property

Public Property Get Key() As String
    Key = msKey
End Property

Public Property Get ButtonType() As glButtonType
    ButtonType = mlButtonType
End Property

Public Property Get Report() As CReport
    Set Report = mclsReport
End Property

' =====================================================================================================================
' Functions
' =====================================================================================================================

Private Function bAddButton( _
    ByVal cbButton As CommandButton _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to initialise a button's settings
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    class_initialize
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 12 Jan 2015   James Boon      0.1         Created
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bAddButton"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim sButtonType As String
    Dim lButtonType As glButtonType
    Dim sKey As String
    Dim clsReport As CReport
    Dim clsReportOutput As CReportOutput
    Dim lFirstUnderscore As Long
    Dim lSecondUnderscore As Long
    Dim lLength As Long
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    Me.Button = cbButton

    ' extract info from button name
    lLength = Len(Me.Button.Name)
    lFirstUnderscore = InStr(1, Me.Button.Name, "_")
    lSecondUnderscore = InStr(lFirstUnderscore + 1, Me.Button.Name, "_")
    sButtonType = Mid(Me.Button.Name, lFirstUnderscore + 1, lSecondUnderscore - lFirstUnderscore - 1)
    sKey = Mid(Me.Button.Name, lSecondUnderscore + 1, lLength - lSecondUnderscore)
    Me.Key = sKey

    ' set button type according to button name
    Select Case sButtonType
        Case "Input"
            lButtonType = glButtonType.InputReport
        Case "Output"
            lButtonType = glButtonType.OutputReport
        Case "Action"
            lButtonType = glButtonType.Action
        Case Else
            Err.Raise Number:=glHANDLED_ERROR, Description:=gsOUTPUT_UNEXPECTED_CASE
    End Select
    Me.ButtonType = lButtonType

    ' set report according to button name
    For Each clsReport In gclsReportCollection.Reports
        Select Case Me.ButtonType
            Case glButtonType.InputReport
                If clsReport.Key = Me.Key Then
                    Me.Report = clsReport
                End If
            Case glButtonType.OutputReport
                For Each clsReportOutput In clsReport.ReportOutputs
                    If clsReportOutput.Key = Me.Key Then
                        Me.Report = clsReport
                    End If
                Next
            Case glButtonType.Action 'do nothing
            Case Else
                Err.Raise Number:=glHANDLED_ERROR, Description:=gsOUTPUT_UNEXPECTED_CASE
        End Select
    Next 'For Each clsReport In gclsReportCollection.Reports
    
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bAddButton = bReturn
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
' Subroutines
' =====================================================================================================================

Public Sub mcbButton_Click() 'this is an entry point so requires error handling wrapper
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "mcbButton_Click"
    
' ---------------------------------------------------------------------------------------------------------------------
' declarations
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    If Not bFormButtonHandler(Me.ButtonType, Me.Key, Me.Report) Then Err.Raise Number:=glHANDLED_ERROR

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    Exit Sub

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSource, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Sub

Public Sub AddButton(ByVal cbButton As CommandButton)
    If Not bAddButton(cbButton) Then Err.Raise Number:=glHANDLED_ERROR
End Sub
