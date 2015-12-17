Attribute VB_Name = "M_2_ErrorHandler"
' =====================================================================================================================
'
' Generic central error handler (adapted from Professional Excel Development (Bovey, Wallentin, Bullen, Green))
'
' ---------------------------------------------------------------------------------------------------------------------
'
' MErrorHandler
' -------------
'
' Module to hold code for error handler
'
' =====================================================================================================================

Option Explicit
Option Private Module

' =====================================================================================================================
' Module constant declarations
' =====================================================================================================================

Public Const gbDEBUG_MODE As Boolean = True
Public Const glHANDLED_ERROR As Long = 9999
Public Const glUSER_CANCEL As Long = 18

Private Const msSILENT_ERROR As String = gsUSER_CANCEL
Private Const msFILE_ERROR_LOG As String = gsAPP_TITLE & gsERROR_LOG

' =====================================================================================================================
' Functions
' =====================================================================================================================

Public Function bCentralErrorHandler( _
    ByVal sModule As String, _
    ByVal sProc As String, _
    Optional ByVal sFile As String, _
    Optional ByVal bEntryPoint As Boolean _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure for central error handler, run each time a bFunction ends
' Arguments:    sModule                     Module housing procedure which called the function
'               sProc                       Procedure which called the function
'               sFile                       optional file name to output in error log
'               bEntryPoint                 True if called by an Entry Point procedure
' Returns:      Boolean                     True if in debug mode, false if silent error or in production mode
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 10 Jul 2014   James Boon                  Created
' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Static sErrMsg As String

    Dim iFile As Integer
    Dim lErrNum As Long
    Dim sFullSource As String
    Dim sPath As String
    Dim sLogText As String

' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    
    lErrNum = Err.Number
    If lErrNum = glUSER_CANCEL Then sErrMsg = msSILENT_ERROR
    If Len(sErrMsg) = 0 Then sErrMsg = Err.Description
    
    On Error Resume Next
    
    If Len(sFile) = 0 Then sFile = ThisWorkbook.Name
    
    sPath = ThisWorkbook.Path
    If Right$(sPath, 1) <> Application.PathSeparator Then sPath = sPath & Application.PathSeparator
    
    sFullSource = "[" & sFile & "]" & sModule & "." & sProc
    sLogText = " " & sFullSource & ", Error " & CStr(lErrNum) & ": " & sErrMsg
    
    iFile = FreeFile()
    Open sPath & msFILE_ERROR_LOG For Append As #iFile
    Print #iFile, Format$(Now(), "dd-mmm-yyyy hh:mm:ss"); sLogText
    If bEntryPoint Then Print #iFile,
    Close #iFile
    
    If sErrMsg <> msSILENT_ERROR Then
    
        If bEntryPoint Or gbDEBUG_MODE Then
            
            Application.ScreenUpdating = True
            MsgBox Prompt:=sErrMsg, Buttons:=vbOKOnly + vbCritical, Title:=gsAPP_TITLE
            sErrMsg = vbNullString
        
        End If
            
        bCentralErrorHandler = gbDEBUG_MODE
            
    Else
        
        If bEntryPoint Then sErrMsg = vbNullString
        bCentralErrorHandler = False
            
    End If

End Function


