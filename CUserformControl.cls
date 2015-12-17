VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CUserformControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mfrmDashboard As FDailyDashboard
Private mcolListener As New Collection
Private Const msMODULE As String = "CUserformControl"

Public Enum glButtonType
    [_First] = 1
        InputReport = glButtonType.[_First]
        OutputReport
        Action
    [_Temp]
    [_Last] = glButtonType.[_Temp] - 1
End Enum

Public Property Get Form() As Object
    Set Form = mfrmDashboard
End Property

' =====================================================================================================================
' Functions
' =====================================================================================================================

Public Function bInitialise( _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to initialise userform
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    class_initialize
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 08 Dec 2015   James Boon      1.0         Created
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bInitialise"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim ctLoop As Object
    Dim clsListener As CListener

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Set mfrmDashboard = New FDailyDashboard

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
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

Public Function bGetButtons( _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to initialise userform
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    class_initialize
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 08 Dec 2015   James Boon      1.0         Created
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bGetButtons"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim ctLoop As Object
    Dim clsListener As CListener

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    For Each ctLoop In mfrmDashboard.Controls
    
        If TypeName(ctLoop) = "CommandButton" Then
            Set clsListener = New CListener
            Call clsListener.AddButton(ctLoop)
            mcolListener.Add clsListener
        End If
    
    Next
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bGetButtons = bReturn
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

Public Function bRefreshForm( _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to refresh data on form
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 02 Dec 2015   James Boon      1.0         Added to dashboard project
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bRefreshForm"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim sCaption As String
    Dim dteNow As Date
    Dim sExceptions As String
    
    Dim clsReport As CReport
    Dim clsReportOutput As CReportOutput
    

    Dim lbl As Object

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    dteNow = Now()
    
    For Each clsReport In gclsReportCollection.Reports
    
        ' data source check
        
        If Not bGetSetting(gsSETTINGS_FILE, clsReport.Key & "LatestEmailTime", sCaption) Then Err.Raise Number:=glHANDLED_ERROR
        
        If sCaption <> vbNullString Then
            
            clsReport.TimeLabel.Caption = sCaption
            
            If dteNow - DateValue(sCaption) > 1 Then
                clsReport.TimeFrame.BackColor = RGB(235, 196, 203)
                clsReport.TimeHeading.BackColor = RGB(235, 196, 203)
                clsReport.TimeLabel.BackColor = RGB(235, 196, 203)
            Else
                clsReport.TimeFrame.BackColor = RGB(206, 219, 175)
                clsReport.TimeHeading.BackColor = RGB(206, 219, 175)
                clsReport.TimeLabel.BackColor = RGB(206, 219, 175)
            End If
        
        End If
    
        For Each clsReportOutput In clsReport.ReportOutputs
        
            sCaption = vbNullString
            If Not bGetSetting(gsSETTINGS_FILE, clsReportOutput.ExceptionsSetting, sCaption) Then Err.Raise Number:=glHANDLED_ERROR
            clsReportOutput.ExceptionsLabel.Caption = sCaption
            If sCaption = vbNullString Then
            Else
                If sCaption > 0 Then
                    clsReportOutput.ExceptionsLabel.BackColor = RGB(235, 196, 203)
                Else
                    clsReportOutput.ExceptionsLabel.BackColor = RGB(206, 219, 175)
                End If
            End If
        
        Next 'For Each clsReportOutput In clsReport.ReportOutputs
    
    Next 'For Each clsReport In gclsReportCollection.Reports
    
    mfrmDashboard.Repaint
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bRefreshForm = bReturn
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

Public Function bShowForm( _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to close down userform
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    class_initialize
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 08 Dec 2015   James Boon      1.0         Created
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bShowForm"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations


' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    mfrmDashboard.Show
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bShowForm = bReturn
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

Public Function bClose( _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to close down userform
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    class_initialize
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 08 Dec 2015   James Boon      1.0         Created
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bClose"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations


' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Unload mfrmDashboard
    Set mfrmDashboard = Nothing
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bClose = bReturn
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

