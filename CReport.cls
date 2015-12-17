VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcolEmails As Collection
Private mlReport As glReport
Private msLatestAttachmentPath As String
Private msMatchingString As String
Private msKey As String
Private moTimeLabel As Object
Private moTimeFrame As Object
Private moTimeHeading As Object
Private mcolReportOutputs As Collection
Private msReportOutputFile As String
Private msProcessingSubroutine As String
Private moStatusLabel As Object

Private Const msMODULE As String = "CReport"

Public Enum glReport
    [_First] = 1
        CostToRevenue = glReport.[_First]
        FecOverheads
        ScheduleToBudget
        BurdenToSchedule
        BurdenCostCode
        BurdenBug
    [_Temp]
    [_Last] = glReport.[_Temp] - 1
End Enum

Public Property Let ReportIndex(lValue As glReport)
    mlReport = lValue
    If Not bGetReportSpecificAttributes Then Err.Raise Number:=glHANDLED_ERROR
End Property

Public Property Let LatestAttachmentPath(sValue As String)
    msLatestAttachmentPath = sValue
End Property

Public Property Get Emails() As Collection
    Set Emails = mcolEmails
End Property

Public Property Get ReportOutputs() As Collection
    Set ReportOutputs = mcolReportOutputs
End Property

Public Property Get MatchingString() As String
    MatchingString = msMatchingString
End Property

Public Property Get Key() As String
    Key = msKey
End Property

Public Property Get TimeLabel() As Object
    Set TimeLabel = moTimeLabel
End Property

Public Property Get TimeFrame() As Object
    Set TimeFrame = moTimeFrame
End Property

Public Property Get TimeHeading() As Object
    Set TimeHeading = moTimeHeading
End Property

Public Property Get StatusLabel() As Object
    Set StatusLabel = moStatusLabel
End Property

Public Property Get ReportIndex() As glReport
    ReportIndex = mlReport
End Property

Public Property Get LatestAttachmentPath() As String
    LatestAttachmentPath = msLatestAttachmentPath
End Property

Public Property Get ReportOutputPath() As String
    ReportOutputPath = msReportOutputFile
End Property

Public Property Get ProcessingSubroutine() As String
    ProcessingSubroutine = msProcessingSubroutine
End Property

' =====================================================================================================================
' Functions
' =====================================================================================================================

Public Function bGetReportSpecificAttributes( _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to get attributes depending on report type
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
    Const sSource As String = "bGetReportSpecificAttributes"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim clsReportOutput As CReportOutput
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    Select Case Me.ReportIndex
        
        Case glReport.CostToRevenue
            msMatchingString = "*UO Projects Cost to Revenue Reconciliation*"
            msKey = "CostToRevenue"
            Set moTimeLabel = gclsDashboard.Form.lblCostToRevenueTime
            Set moTimeFrame = gclsDashboard.Form.frmCostToRevenue
            Set moStatusLabel = gclsDashboard.Form.lblStatusCostToRevenue
            Set moTimeHeading = gclsDashboard.Form.lblCostToRevenue
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "CostToRevenueExceptions"
            clsReportOutput.ExceptionsSetting = "CostToRevenueExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExCostToRevenue
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            msReportOutputFile = "O:\Deptserv\RA Development\James Boon\Other Projects\Dashboard\Report Output.xlsm"
            
        Case glReport.FecOverheads
            msMatchingString = "*UO Projects DAC and IC Reconciliation Data*"
            msKey = "FecOverheads"
            Set moTimeLabel = gclsDashboard.Form.lblFecOhReconciliationDataTime
            Set moTimeFrame = gclsDashboard.Form.frmFecOhReconciliationData
            Set moStatusLabel = gclsDashboard.Form.lblStatusFecOhReconciliationData
            Set moTimeHeading = gclsDashboard.Form.lblFecOhReconciliationData
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "DACICExceptions"
            clsReportOutput.ExceptionsSetting = "DACICExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExDACIC
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "PABudgetExceptions"
            clsReportOutput.ExceptionsSetting = "PABudgetExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExPABudget
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "PACostExceptions"
            clsReportOutput.ExceptionsSetting = "PACostExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExPACost
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            msReportOutputFile = "O:\Deptserv\RA Development\James Boon\Other Projects\Dashboard\FEC Overheads Output.xlsm"
            
        Case glReport.ScheduleToBudget
            msMatchingString = "*UO Projects Burden Schedule and Budget Validation Data*"
            msKey = "ScheduleToBudget"
            Set moTimeLabel = gclsDashboard.Form.lblBurdenToBudgetValidationTime
            Set moTimeFrame = gclsDashboard.Form.frmBurdenToBudgetValidation
            Set moStatusLabel = gclsDashboard.Form.lblStatusBurdenToBudgetValidation
            Set moTimeHeading = gclsDashboard.Form.lblBurdenToBudgetValidation
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "FECScheduleToBudgetExceptions"
            clsReportOutput.ExceptionsSetting = "FECScheduleToBudgetExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExFECSchToBud
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "NonFECScheduleToBudgetExceptions"
            clsReportOutput.ExceptionsSetting = "NonFECScheduleToBudgetExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExNonFECSchToBud
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            msReportOutputFile = "O:\Deptserv\RA Development\James Boon\Other Projects\Dashboard\Burden Schedule Creation and Maintenance v3.1.xlsm"
        
        Case glReport.BurdenToSchedule
            msMatchingString = "*UO Project Burdening Validation*"
            msKey = "BurdenToSchedule"
            Set moTimeLabel = gclsDashboard.Form.lblBurdenValidationTime
            Set moTimeFrame = gclsDashboard.Form.frmBurdenValidation
            Set moStatusLabel = gclsDashboard.Form.lblStatusBurdenValidation
            Set moTimeHeading = gclsDashboard.Form.lblBurdenValidation
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "BurdenToScheduleExceptions"
            clsReportOutput.ExceptionsSetting = "BurdenToScheduleExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExBurdenToSchedule
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            msReportOutputFile = "O:\Deptserv\RA Development\James Boon\Other Projects\Dashboard\Report Output.xlsm"
        
        Case glReport.BurdenCostCode
            msMatchingString = "*UO Projects Burden Cost Code Mismatch*"
            msKey = "BurdenCostCode"
            Set moTimeLabel = gclsDashboard.Form.lblBurdenCostCodeMismatchTime
            Set moTimeFrame = gclsDashboard.Form.frmBurdenCostCodeMismatch
            Set moStatusLabel = gclsDashboard.Form.lblStatusBurdenCostCodeMismatch
            Set moTimeHeading = gclsDashboard.Form.lblBurdenCostCodeMismatch
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "BurdenCostCodeExceptions"
            clsReportOutput.ExceptionsSetting = "BurdenCostCodeExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExBurdenCostCode
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            msReportOutputFile = "O:\Deptserv\RA Development\James Boon\Other Projects\Dashboard\Report Output.xlsm"
        
        Case glReport.BurdenBug
            msMatchingString = "*UO Projects Burdening Bug Data*"
            msKey = "BurdenBug"
            Set moTimeLabel = gclsDashboard.Form.lblBurdeningBugTime
            Set moTimeFrame = gclsDashboard.Form.frmBurdeningBug
            Set moStatusLabel = gclsDashboard.Form.lblStatusBurdeningBug
            Set moTimeHeading = gclsDashboard.Form.lblBurdeningBug
            Set clsReportOutput = New CReportOutput
            clsReportOutput.Key = "BurdenBugExceptions"
            clsReportOutput.ExceptionsSetting = "BurdenBugExceptions"
            clsReportOutput.ExceptionsLabel = gclsDashboard.Form.lblExBurdeningBug
            Me.ReportOutputs.Add Item:=clsReportOutput, Key:=clsReportOutput.Key
            msReportOutputFile = "O:\Deptserv\RA Development\James Boon\Other Projects\Dashboard\Report Output.xlsm"
            
    End Select
    

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bGetReportSpecificAttributes = bReturn
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

Public Function bGatherEmails( _
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
    Const sSource As String = "bGatherEmails"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

'    Dim attLoop As Object
    Dim sOutputFileName As String
    Dim dteLatestEmail As Date
    Dim wkbReportOutput As Excel.Workbook
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    Me.StatusLabel.Caption = "Getting Emails"
    
    If Not gclsOutlook.bGatherEmails(Me.MatchingString, Me.Emails, dteLatestEmail) Then Err.Raise Number:=glHANDLED_ERROR
    
    If Me.Emails.Count > 0 Then
        
        If Not bSaveSetting(gsSETTINGS_FILE, Me.Key & "LatestEmailTime", dteLatestEmail) Then Err.Raise Number:=glHANDLED_ERROR
    
        Me.StatusLabel.Caption = "Saving Attachments"
        If Not gclsOutlook.bSaveLatestAttachment(Me.Emails, sOutputFileName) Then Err.Raise Number:=glHANDLED_ERROR
        Me.LatestAttachmentPath = sOutputFileName
        If Not bSaveSetting(gsSETTINGS_FILE, Me.Key & "LatestFilePath", Me.LatestAttachmentPath) Then Err.Raise Number:=glHANDLED_ERROR

        Me.StatusLabel.Caption = "Processing"
        
        If Me.ReportIndex = glReport.ScheduleToBudget Then gclsDashboard.Form.Hide
        
        
        Set wkbReportOutput = Workbooks.Open(Me.ReportOutputPath)
        Application.Run "'" & wkbReportOutput.Name & "'!" & "ProcessData", Me.ReportIndex, Me.LatestAttachmentPath, Me.Key & "Exceptions"
        wkbReportOutput.Save
        wkbReportOutput.Close savechanges:=False
        Set wkbReportOutput = Nothing
        
        If Me.ReportIndex = glReport.ScheduleToBudget Then gclsDashboard.Form.Show vbModeless
        
        Me.StatusLabel.Caption = "Refreshing"
        If Not gclsDashboard.bRefreshForm Then Err.Raise Number:=glHANDLED_ERROR
        
        Me.StatusLabel.Caption = "Archiving"
        If Not gclsOutlook.bArchiveEmails(Me.Emails) Then Err.Raise Number:=glHANDLED_ERROR
    
    End If

    Me.StatusLabel.Caption = "Done"


' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bGatherEmails = bReturn
    On Error Resume Next
    wkbReportOutput.Close savechanges:=False
    Set wkbReportOutput = Nothing
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

Private Sub class_initialize()
    Set mcolEmails = New Collection
    Set mcolReportOutputs = New Collection
End Sub






