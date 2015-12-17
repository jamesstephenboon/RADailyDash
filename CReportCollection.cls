VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CReportCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit



Private Const msMODULE As String = "CReportCollection"

Private mcolReports As Collection

Public Property Get Reports() As Collection
    Set Reports = mcolReports
End Property

' =====================================================================================================================
' Functions
' =====================================================================================================================

Public Function bReadOutlookEmails( _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to get reports from outlook inbox
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
    Const sSource As String = "bReadOutlookEmails"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim clsReport As CReport

' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    For Each clsReport In Me.Reports
        If Not clsReport.bGatherEmails Then Err.Raise Number:=glHANDLED_ERROR
        DoEvents
    Next 'For Each clsReport In Me.Reports
        
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bReadOutlookEmails = bReturn
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

Public Sub class_initialize()

    Dim lReport As glReport
    Dim clsReport As CReport

    Set mcolReports = New Collection
    
    For lReport = glReport.[_First] To glReport.[_Last]
    
            Set clsReport = New CReport
            clsReport.ReportIndex = lReport
            Me.Reports.Add Item:=clsReport
    
    Next 'For lReport = glReport.[_First] To glReport.[_Last]
    
    
End Sub

