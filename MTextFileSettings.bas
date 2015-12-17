Attribute VB_Name = "MTextFileSettings"
'
' Module for utilities code
'

Private Const msMODULE As String = "MTextFileSettings"

Option Explicit
Option Private Module


Public Function bSaveSetting(ByVal sFile As String, ByVal sName As String, ByVal sValue As String) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to save setting to text file
'
' Arguments:    sFile           settings file
'               sName           variable name
'               sValue          variable value
'
' Returns:      Boolean         True on success, false on error
'
' Called by:
'
'
' Date          Developer       Action
' ---------------------------------------------------------------------------------------------------------------------
' 10 Jul 2014   James Boon      Created
'
'
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bSaveSetting"
    Dim bReturn As Boolean
    bReturn = True
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Dim lFileNumA As Long
    Dim lFileNumB As Long
    Dim sXFile As String
    Dim sVarName As String
    Dim sVarValue As String
    
    If Not IsFullName(sFile) Then
        sXFile = ThisWorkbook.Path & "\X" & sFile
        sFile = ThisWorkbook.Path & "\" & sFile
    Else
        sXFile = FullNameToPath(sFile) & "\X" & FullNameToFileName(sFile)
    End If
        
    If FileExists(sFile) Then
        lFileNumA = FreeFile
        Open sFile For Input As lFileNumA
        lFileNumB = FreeFile
        Open sXFile For Output As lFileNumB
        Do While Not EOF(lFileNumA)
            Input #lFileNumA, sVarName, sVarValue
            If sVarName <> sName Then
                Write #lFileNumB, sVarName, sVarValue
            End If
        Loop 'Do While Not EOF(iFileNumA)
        Write #lFileNumB, sName, sValue
        Close #lFileNumA
        Close #lFileNumB
        FileCopy sXFile, sFile
        Kill sXFile
    Else 'If FileExists(sFile) Then
        lFileNumB = FreeFile
        Open sFile For Output As lFileNumB
        Write #lFileNumB, sName, sValue
        Close #lFileNumB
    End If 'If FileExists(sFile) Then

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bSaveSetting = bReturn
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





Public Function bGetSetting(ByVal sFile As String, ByVal sName As String, ByRef sValue As String) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to get setting from text file
'
' Arguments:    sFile           settings file
'               sName           variable name
'               sValue          variable value
'
' Returns:      Boolean         True on success, false on error
'
' Called by:
'
'
' Date          Developer       Action
' ---------------------------------------------------------------------------------------------------------------------
' 10 Jul 2014   James Boon      Created
'
'
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bGetSetting"
    Dim bReturn As Boolean
    bReturn = True
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Dim lFileNum As Long
    Dim sVarName As String
    Dim sVarValue As String
    Dim bFoundSetting As Boolean
    
    bFoundSetting = False
    
    If Not IsFullName(sFile) Then
        sFile = ThisWorkbook.Path & "\" & sFile
    End If
    
    If FileExists(sFile) Then
        lFileNum = FreeFile
        Open sFile For Input As lFileNum
        Do While Not EOF(lFileNum) And Not bFoundSetting
            Input #lFileNum, sVarName, sVarValue
            If sVarName = sName Then
                sValue = CStr(sVarValue)
                bFoundSetting = True
            End If
        Loop 'Do While Not EOF(lFileNum)
        Close #lFileNum
    End If 'If FileExists(sFile) Then
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bGetSetting = bReturn
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





