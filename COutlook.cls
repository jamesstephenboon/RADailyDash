VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COutlook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE As String = "COutlook"

Private molApp As Object
Private molNamespace As Object
Private molfInbox As Object
Private molfArchive As Object


Private Sub class_initialize()
    Set molApp = CreateObject("Outlook.Application")
    Set molNamespace = molApp.GetNamespace("MAPI")
    Set molfInbox = molNamespace.GetDefaultFolder(6) 'this causes error if no outlook session is allowed
    Set molfArchive = molfInbox.Folders("Archive")
End Sub

' =====================================================================================================================
' Functions
' =====================================================================================================================

Public Function bGatherEmails( _
    ByVal sMatchingString As String, _
    ByVal colEmails As Collection, _
    ByRef dteLatestEmail As Date _
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

    Dim olEmail As Object
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    
    For Each olEmail In molfInbox.items
        If olEmail.Subject Like sMatchingString Then
            colEmails.Add Item:=olEmail
        End If
    Next
    
    If Not bGetLatestEmail(colEmails, olEmail) Then Err.Raise Number:=glHANDLED_ERROR
    
    If Not olEmail Is Nothing Then
        dteLatestEmail = olEmail.ReceivedTime
    End If
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bGatherEmails = bReturn
    On Error Resume Next
    Set olEmail = Nothing
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

Public Function bSaveLatestAttachment( _
    ByVal colEmails As Collection, _
    ByRef sOutputFileName As String _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to save latest attachment and return file path
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
    Const sSource As String = "bSaveLatestAttachment"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim attLoop As Object
    Dim olLatestEmail As Object
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    If colEmails.Count > 0 Then
    
        If Not bGetLatestEmail(colEmails, olLatestEmail) Then Err.Raise Number:=glHANDLED_ERROR
        
        If olLatestEmail.Attachments.Count > 0 Then
            
            For Each attLoop In olLatestEmail.Attachments
                
                If Not bGetSetting(gsSETTINGS_FILE, "DefaultOutputPath", sOutputFileName) Then Err.Raise Number:=glHANDLED_ERROR
                sOutputFileName = sOutputFileName & Application.PathSeparator & gsTimeStamp
                If Dir(sOutputFileName, vbDirectory) = vbNullString Then MkDir Path:=sOutputFileName
                sOutputFileName = sOutputFileName & Application.PathSeparator & attLoop.FileName
                
                If Dir(sOutputFileName) = vbNullString Then
                    attLoop.SaveAsFile sOutputFileName
                Else
                    Err.Raise Number:=glHANDLED_ERROR
                End If
            
            Next
        
        Else
            
            sOutputFileName = "none"
            
        End If
        
    End If
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bSaveLatestAttachment = bReturn
    On Error Resume Next
    Set olLatestEmail = Nothing
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

Public Function bArchiveEmails( _
    ByVal colEmails As Collection _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to archive emails from a collection
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
    Const sSource As String = "bArchiveEmails"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim olEmail As Object
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    For Each olEmail In colEmails
        olEmail.UnRead = False
        olEmail.Move molfArchive
    Next
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bArchiveEmails = bReturn
    On Error Resume Next
    Set olEmail = Nothing
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

Private Function bGetLatestEmail( _
    ByVal colEmails As Collection, _
    ByRef olLatestEmail As Object _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to get latest email from a collection of emails
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
    Const sSource As String = "bGetLatestEmail"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim dteLatest As Date
    Dim olEmail As Object
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    Set olLatestEmail = Nothing
    
    For Each olEmail In colEmails
        If olEmail.ReceivedTime > dteLatest Then
            dteLatest = olEmail.ReceivedTime
            Set olLatestEmail = olEmail
        End If
    Next 'For Each oEmail In Me.Emails
    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bGetLatestEmail = bReturn
    On Error Resume Next
    Set olEmail = Nothing
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
