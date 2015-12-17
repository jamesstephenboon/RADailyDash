Attribute VB_Name = "M_4_ProjectCode"
' =====================================================================================================================
'
' Daily Exceptions Dashboard
' Research Accounts, Finance Division, University of Oxford
' v1.0 - James Boon - Dec 2015
'
' ---------------------------------------------------------------------------------------------------------------------
'
' M_4_ProjectCode
' ---------------
'
' Module for project specific code. This includes a button handler, the main wrapper routine, and some small utility
' functions and routines.
'
' ---------------------------------------------------------------------------------------------------------------------
'
' Flow Chart
' ----------
'
' bButtonHandler (called by M_1_EntryPoints)

'
' ---------------------------------------------------------------------------------------------------------------------
'
' Requirements
' ------------
'
' Needs following modules: M_1_EntryPoints, M_2_ErrorHandler, M_3_GlobalVariables
'
' ---------------------------------------------------------------------------------------------------------------------
'
' Version Control
' ---------------
'
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 02 Dec 2015   James Boon      1.0         Updating from v0.7 to tidy up coding, add user forms, turn into addin
' =====================================================================================================================

Option Explicit
Option Private Module

' =====================================================================================================================
' Global and Module variable declarations
' =====================================================================================================================

Public Const gsAPP_TITLE As String = "Daily Exceptions Dashboard"
Private Const msMODULE As String = "M_4_ProjectCode"

' ribbon buttons
Public Const gsDAILY_UPDATE As String = "btnDailyUpdate"
Public Const gsOPEN_DASHBOARD As String = "btnOpenDashboard"

' other declarations
Public gclsDashboard As CUserformControl
Public gclsReportCollection As CReportCollection
Public gclsOutlook As COutlook

' =====================================================================================================================
' Functions
' =====================================================================================================================

Public Function bButtonHandler( _
    ByVal ircControl As IRibbonControl _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to handle program flow depending on which ribbon button was pressed
' Arguments:    ircControl                  Tells us which ribbon button was pressed
' Returns:      Boolean                     True on success, false on error
' Called by:    M_1_EntryPoints.RibbonEntryPoint
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 18 Aug 2015   James Boon      1.0         Created template version
' 02 Dec 2015   James Boon      1.0         Added to dashboard project
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bButtonHandler"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Select Case ircControl.id

'        Case gsDAILY_UPDATE
'            If Not bDailyUpdate Then Err.Raise Number:=glHANDLED_ERROR
            
        Case gsOPEN_DASHBOARD
            If Not bOpenDashboard Then Err.Raise Number:=glHANDLED_ERROR
    
        Case Else
            Err.Raise Number:=glHANDLED_ERROR, Description:=gsOUTPUT_UNEXPECTED_CASE

    End Select

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bButtonHandler = bReturn
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

Private Function bOpenDashboard( _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to open dashboard userform
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    bButtonHandler
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 02 Dec 2015   James Boon      1.0         Added to dashboard project
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bOpenDashboard"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations


' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Set gclsDashboard = New CUserformControl
    If Not gclsDashboard.bInitialise Then Err.Raise Number:=glHANDLED_ERROR
    Set gclsReportCollection = New CReportCollection 'initialises with set of report classes
    If Not gclsDashboard.bGetButtons Then Err.Raise Number:=glHANDLED_ERROR
    If Not gclsDashboard.bRefreshForm Then Err.Raise Number:=glHANDLED_ERROR
    If Not gclsDashboard.bShowForm Then Err.Raise Number:=glHANDLED_ERROR


' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bOpenDashboard = bReturn
    On Error Resume Next
    Set gclsDashboard = Nothing
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

Public Function bFormButtonHandler( _
    ByVal lButtonType As glButtonType, _
    ByVal sKey As String, _
    ByVal clsReport As CReport _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to handle button clicks on form
' Arguments:    none
' Returns:      Boolean                     True on success, false on error
' Called by:    bButtonHandler
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 02 Dec 2015   James Boon      1.0         Added to dashboard project
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bFormButtonHandler"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim sFile As String
    Dim wkbOpen As Excel.Workbook

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    
    Select Case lButtonType
        
        Case glButtonType.InputReport
        
            If Not bGetSetting(gsSETTINGS_FILE, clsReport.Key & "LatestFilePath", sFile) Then Err.Raise Number:=glHANDLED_ERROR
            If sFile = "none" Then
                MsgBox "no file"
            Else
                Set wkbOpen = Workbooks.Open(FileName:=sFile)
                wkbOpen.Activate
                If Not gclsDashboard.bClose Then Err.Raise Number:=glHANDLED_ERROR
            End If
        
        Case glButtonType.OutputReport
        
            sFile = clsReport.ReportOutputPath
            Set wkbOpen = Workbooks.Open(FileName:=sFile)
            wkbOpen.Activate
            If Not gclsDashboard.bClose Then Err.Raise Number:=glHANDLED_ERROR
        
        Case glButtonType.Action
        
            Select Case sKey
                Case "UpdateFromEmails"
                    Set gclsOutlook = New COutlook
                    If Not gclsReportCollection.bReadOutlookEmails Then Err.Raise Number:=glHANDLED_ERROR
                Case Else
                    Err.Raise Number:=glHANDLED_ERROR, Description:=gsOUTPUT_UNEXPECTED_CASE
            End Select
        
        Case Else
            
            Err.Raise Number:=glHANDLED_ERROR, Description:=gsOUTPUT_UNEXPECTED_CASE
        
    End Select
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bFormButtonHandler = bReturn
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

