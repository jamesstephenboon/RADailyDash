Attribute VB_Name = "MUtilities"
''
'' Module for utilities code
''
'
'Private Const msMODULE As String = "MUtilities"
'
'Option Explicit
'Option Private Module
'
'
'
'
'
'Public Function bChooseFile(ByVal sTitle As String, ByRef sChooseFile As String, Optional ByVal vInitialFileName As Variant)
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bChooseFile"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Dim fd As FileDialog
'    Dim lFileChosen As Long
'
'    Set fd = Application.FileDialog(msoFileDialogFilePicker)
'    fd.Title = sTitle
'    If Not IsMissing(vInitialFileName) Then
'        fd.InitialFileName = vInitialFileName
'    End If
'    lFileChosen = fd.Show
'    If lFileChosen <> -1 Then
'        sChooseFile = gsUSER_CANCEL
'    Else
'        sChooseFile = fd.SelectedItems(1)
'    End If
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bChooseFile = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'
'
'
'Public Function bChooseFolder(ByVal sTitle As String, ByRef sChooseFolder As String)
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bChooseFolder"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Dim fd As FileDialog
'    Dim lFolderChosen As Long
'
'    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
'    fd.Title = sTitle
'    lFolderChosen = fd.Show
'    If lFolderChosen <> -1 Then
'        sChooseFolder = gsUSER_CANCEL
'    Else
'        sChooseFolder = fd.SelectedItems(1)
'    End If
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bChooseFolder = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
Public Function bFormatSheet(ByRef wksFormatSheet As Worksheet, Optional ByVal lColFreeze As Long = 0, Optional ByVal lRowFreeze As Long = 1)

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

    On Error GoTo ErrorHandler
    Const sSource As String = "bFormatSheet"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' procedure code


    Dim wksActive As Worksheet

    Set wksActive = ActiveSheet

    ' apply formatting rules
    With wksFormatSheet
        .Cells.Font.Name = "Calibri"
        .Cells.Font.Size = 11
        .Cells.Font.Bold = False
        .Cells.Font.Color = RGB(0, 0, 0)
        .Cells.Interior.ColorIndex = xlNone
        .Range("1:" & lRowFreeze).Font.Bold = True
        .Range("1:" & lRowFreeze).Font.Color = glWhite
        .Range("1:" & lRowFreeze).Interior.Color = glOxfordBlue
        .Select
        ActiveWindow.FreezePanes = False
        Application.GoTo reference:=.Cells(1, 1), Scroll:=True
        .Cells(lRowFreeze + 1, lColFreeze + 1).Select
        ActiveWindow.FreezePanes = True
'        .Cells(1, 1).Select
        .AutoFilterMode = False
        If WorksheetFunction.CountA(.Cells) > 0 Then .Cells.AutoFilter
        .Cells.Columns.AutoFit
    End With        'With wksFormatSheet

    ' re-select sheet that was initially selected
    wksActive.Select

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bFormatSheet = bReturn
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
'
'
'Public Function bDoesObjectExist(ByRef colTest As Collection, ByVal strKey As String, ByRef clsObject As Variant)
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bDoesObjectExist"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Set clsObject = Nothing
'    On Error Resume Next
'        Set clsObject = colTest(strKey)
'    On Error GoTo ErrorHandler
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bDoesObjectExist = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'Public Function bShowAllData(ByRef wksSheet As Variant) As Boolean
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bShowAllData"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    On Error Resume Next
'        wksSheet.ShowAllData
'    On Error GoTo ErrorHandler
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bShowAllData = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'
'
Public Function bTimeStamp(ByRef sTimeStamp As String, Optional ByRef dteTimeStamp As Date) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

    On Error GoTo ErrorHandler
    Const sSource As String = "bTimeStamp"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Dim dteNow As Date

    dteNow = Now()

    sTimeStamp = CStr(Year(dteNow))
    If Month(dteNow) >= 10 Then sTimeStamp = sTimeStamp & CStr(Month(dteNow)) Else sTimeStamp = sTimeStamp & "0" & CStr(Month(dteNow))
    If Day(dteNow) >= 10 Then sTimeStamp = sTimeStamp & CStr(Day(dteNow)) Else sTimeStamp = sTimeStamp & "0" & CStr(Day(dteNow))
    If Hour(dteNow) >= 10 Then sTimeStamp = sTimeStamp & CStr(Hour(dteNow)) Else sTimeStamp = sTimeStamp & "0" & CStr(Hour(dteNow))
    If Minute(dteNow) >= 10 Then sTimeStamp = sTimeStamp & CStr(Minute(dteNow)) Else sTimeStamp = sTimeStamp & "0" & CStr(Minute(dteNow))
    If Second(dteNow) >= 10 Then sTimeStamp = sTimeStamp & CStr(Second(dteNow)) Else sTimeStamp = sTimeStamp & "0" & CStr(Second(dteNow))

    If Not (IsMissing(dteTimeStamp)) Then
        dteTimeStamp = dteNow
    End If

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bTimeStamp = bReturn
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
'
'
'
'Public Function bOpenExcelWorkbook( _
'    ByRef wkbToOpen As Workbook, _
'    ByVal sPathToOpen As String, _
'    ByRef bWasAlreadyOpen As Boolean _
') As Boolean
'
'
'' ---------------------------------------------------------------------------------------------------------------------
'' Comments:     Procedure to open an excel file that may already be open and may not actually exist
''
'' Arguments:    wkbToOpen       the index worksheet
''               sPathToOpen     the path of the file we want to open
''               bWasAlreadyOpen the worksheet to link to
''
'' Returns:      Boolean         True on success, false on error
''
'' Date          Developer       Action
'' ---------------------------------------------------------------------------------------------------------------------
'' 17 Jun 2014   James Boon      Created
''
''
''
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bOpenExcelWorkbook"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Dim sFileName As String
'    Set wkbToOpen = Nothing
'    bWasAlreadyOpen = False
'
'    sFileName = sPathToOpen
'    sFileName = Right(sFileName, Len(sFileName) - InStrRev(sFileName, "\"))
'
'
'    ' check to see if already open
'    On Error Resume Next
'        Set wkbToOpen = Workbooks(sFileName)
'    On Error GoTo ErrorHandler
'
'    ' if not then try to open it and, if fail, raise error
'    If wkbToOpen Is Nothing Then
'
'        Application.DisplayAlerts = False
'            On Error Resume Next
'                Set wkbToOpen = Workbooks.Open(sPathToOpen)
'            On Error GoTo ErrorHandler
'        Application.DisplayAlerts = True
'
'        If wkbToOpen Is Nothing Then Err.Raise Number:=glHANDLED_ERROR, Description:=sPathToOpen & " not found"
'
'    Else
'
'        bWasAlreadyOpen = True
'
'    End If
'
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bOpenExcelWorkbook = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'
'Public Function bChooseOpenFileOrLoadExistingFile( _
'    ByRef bUserCancelled As Boolean, _
'    ByRef bSourceFileWasAlreadyOpen As Boolean, _
'    ByVal sPrompt As String, _
'    ByRef wkbChosen As Workbook, _
'    Optional ByVal vInitialFileName As Variant _
') As Boolean
'
'' ---------------------------------------------------------------------------------------------------------------------
'' Comments:     Procedure to get source file, either from file already open, or by opening a new file
''
'' Arguments:    bUserCancelled  returns whether user cancelled this procedure or not
''               bSourceFileWasAlreadyOpen  tells the calling procedure if we needed to open a file, so the calling procedure
''                               knows to close it again
''               sPrompt         the prompt in the dialogue box
''               wkbChosen       returns the chosen workbook to the calling procedure
''
'' Returns:      Boolean         True on success, false on error
''
'' Called by:    bSplitActivityReport, bSendActivityReport
''
''
'' Date          Developer       Action
'' ---------------------------------------------------------------------------------------------------------------------
'' 13 Jun 2014   James Boon      Created
'' 20 Jun 2014   James Boon      Adapted to be used more generally
''
''
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bGetSourceFile"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Dim sFileName As String
'    Dim bFormStopped As Boolean
'    Dim bOpenOtherFile As Boolean
'
'    bUserCancelled = False
'    bSourceFileWasAlreadyOpen = False
'
'    FChooseWorkbook.Stopped = False
'    FChooseWorkbook.Show
'    sFileName = FChooseWorkbook.FileName
'    bFormStopped = FChooseWorkbook.Stopped
'    bOpenOtherFile = FChooseWorkbook.OpenOtherFile
'    Unload FChooseWorkbook
'
'    If Not (bFormStopped) Then
'        If bOpenOtherFile Then
'            If IsMissing(vInitialFileName) Then
'                If Not bChooseFile(sPrompt, sFileName) Then Err.Raise Number:=glHANDLED_ERROR
'            Else
'                If Not bChooseFile(sPrompt, sFileName, vInitialFileName) Then Err.Raise Number:=glHANDLED_ERROR
'            End If
'            If sFileName <> gsUSER_CANCEL Then
'                Set wkbChosen = Workbooks.Open(sFileName)
'            Else
'                bUserCancelled = True
'            End If
'        Else
'            Set wkbChosen = Workbooks(sFileName)
'            bSourceFileWasAlreadyOpen = True
'        End If
'    Else
'        bUserCancelled = True
'    End If
'
'    'If bUserCancelled Then MsgBox gsOUTPUT_USER_CANCEL_OUTPUT
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bChooseOpenFileOrLoadExistingFile = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'
'Public Function bDeleteAllButOneWorksheet(ByRef wkb As Workbook) As Boolean
'
'' ---------------------------------------------------------------------------------------------------------------------
'' Comments:     Procedure to delete all but one worksheet in a workbook
''
'' Arguments:    wkb             the workbook to delete sheets from (if required)
''
'' Returns:      Boolean         True on success, false on error
''
'' Called by:    bSplitActivityReport, bSendActivityReport
''
''
'' Date          Developer       Action
'' ---------------------------------------------------------------------------------------------------------------------
'' 27 Jun 2014   James Boon      Created
''
''
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bDeleteAllButOneWorksheet"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Dim i As Integer
'
'    If wkb.Sheets.Count > 1 Then
'        Application.DisplayAlerts = False
'        For i = wkb.Sheets.Count To 2 Step -1
'            wkb.Sheets(i).Delete
'        Next i
'        Application.DisplayAlerts = True
'    End If
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bDeleteAllButOneWorksheet = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'Public Function bBackupFile(ByVal wkbFileToBackup As Workbook, Optional ByVal sSetting As String) As Boolean
'
'' ---------------------------------------------------------------------------------------------------------------------
'' Comments:     Procedure to backup a file
''
'' Arguments:    wkbFileToBackup the workbook to back up
''
'' Returns:      Boolean         True on success, false on error
''
'' Called by:
''
''
'' Date          Developer       Action
'' ---------------------------------------------------------------------------------------------------------------------
'' 11 Jul 2014   James Boon      Created
''
''
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bBackupFile"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Dim sTimeStamp As String
'    Dim sBackupFileName As String
'
'    If Not bTimeStamp(sTimeStamp) Then Err.Raise Number:=glHANDLED_ERROR
'    sBackupFileName = wkbFileToBackup.Path & "\" & gsPREFIX_BACKUP_FILE & sTimeStamp & "-" & wkbFileToBackup.Name
'    wkbFileToBackup.SaveCopyAs FileName:=sBackupFileName
'    If Not IsMissing(sSetting) Then
'        If Not bSaveSetting(gsSETTINGS_FILE, sSetting, sBackupFileName) Then Err.Raise Number:=glHANDLED_ERROR
'    End If
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bBackupFile = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'Public Function bDeleteFile(ByRef sFileToDelete As String, Optional sSetting As String) As Boolean
'
'' ---------------------------------------------------------------------------------------------------------------------
'' Comments:     Procedure to backup a file
''
'' Arguments:    wkbFileToBackup the workbook to back up
''
'' Returns:      Boolean         True on success, false on error
''
'' Called by:
''
''
'' Date          Developer       Action
'' ---------------------------------------------------------------------------------------------------------------------
'' 11 Jul 2014   James Boon      Created
''
''
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bDeleteFile"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Kill sFileToDelete
'    If Not IsMissing(sSetting) Then
'        If Not bSaveSetting(gsSETTINGS_FILE, sSetting, "") Then Err.Raise Number:=glHANDLED_ERROR
'    End If
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bDeleteFile = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
Public Function IsFullName(sFile As String) As Boolean
  ' if sFile includes path, it contains path separator "\"
  IsFullName = InStr(sFile, "\") > 0
End Function

Public Function FullNameToPath(sFullName As String) As String
  ''' does not include trailing backslash
  Dim k As Integer
  For k = Len(sFullName) To 1 Step -1
    If Mid(sFullName, k, 1) = "\" Then Exit For
  Next k
  If k < 1 Then
    FullNameToPath = ""
  Else
    FullNameToPath = Mid(sFullName, 1, k - 1)
  End If
End Function

Public Function FullNameToFileName(sFullName As String) As String
  Dim k As Integer
  Dim sTest As String
  If InStr(1, sFullName, "[") > 0 Then
    k = InStr(1, sFullName, "[")
    sTest = Mid(sFullName, k + 1, InStr(1, sFullName, "]") - k - 1)
  Else
    For k = Len(sFullName) To 1 Step -1
      If Mid(sFullName, k, 1) = "\" Then Exit For
    Next k
    sTest = Mid(sFullName, k + 1, Len(sFullName) - k)
  End If
  FullNameToFileName = sTest
End Function

Public Function FileExists(ByVal FileSpec As String) As Boolean
   ' by Karl Peterson MS MVP VB
   Dim Attr As Long
   ' Guard against bad FileSpec by ignoring errors
   ' retrieving its attributes.
   On Error Resume Next
   Attr = GetAttr(FileSpec)
   If Err.Number = 0 Then
      ' No error, so something was found.
      ' If Directory attribute set, then not a file.
      FileExists = Not ((Attr And vbDirectory) = vbDirectory)
   End If
End Function
'
'Public Function GetWorkbookAddressFromExternalAddress(ByVal rngCell As Range) As String
'
'    Dim sAddress As String, lLeft As Long, lRight As Long
'
'    sAddress = rngCell.Address(external:=True)
'    lLeft = InStr(1, sAddress, "[")
'    lRight = InStr(lLeft, sAddress, "]")
'    sAddress = Left(sAddress, lLeft - 1) + Mid(sAddress, lRight + 1)
'    GetWorkbookAddressFromExternalAddress = sAddress
'
'End Function
'
'
'Public Sub HideGridlines(wks As Worksheet)
'
'    wks.Activate
'    wks.Select
'    wks.Parent.Windows(1).DisplayGridlines = False
'
'End Sub
'
'Public Function bCopyNamedRanges(ByVal wksFrom As Worksheet, ByVal wksTo As Worksheet) As Boolean
'
'' ---------------------------------------------------------------------------------------------------------------------
'' Comments:     Procedure to copy named ranges from one worksheet to another
''
'' Arguments:    wksFrom         source worksheet
''               wksTo           target worksheet
''
'' Returns:      Boolean         True on success, false on error
''
'' Called by:
''
''
'' Date          Developer       Action
'' ---------------------------------------------------------------------------------------------------------------------
'' 26 Aug 2014   James Boon      Created
''
''
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bCopyNamedRanges"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Dim nameLoop As Name
'
'    For Each nameLoop In wksFrom.Names
'        wksTo.Names.Add Name:=Right(nameLoop.Name, Len(nameLoop.Name) - InStr(nameLoop.Name, "!")), RefersTo:=wksTo.Range(nameLoop.RefersToRange.Address)
'    Next
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bCopyNamedRanges = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'Public Function bPageSetup(ByRef wks As Worksheet) As Boolean
'
'' ---------------------------------------------------------------------------------------------------------------------
'' Comments:     Procedure to manage page setup
''
'' Arguments:    wks             worksheet to setup
''
'' Returns:      Boolean         True on success, false on error
''
'' Called by:
''
''
'' Date          Developer       Action
'' ---------------------------------------------------------------------------------------------------------------------
'' 26 Aug 2014   James Boon      Created
''
''
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bPageSetup"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    ' print areas and setup
''    With wks.PageSetup
''        .Zoom = False
''        .FitToPagesWide = glONE_PAGE_WIDE
''        .Orientation = xlLandscape
''        .PrintArea = wks.UsedRange.Resize( _
''            wks.UsedRange.Rows.Count - glHEADER_ROWS_TO_REMOVE_FROM_PRINT_AREA, _
''            wks.UsedRange.Columns.Count _
''        ).Offset(glHEADER_ROWS_TO_REMOVE_FROM_PRINT_AREA, 0).Address
''    End With
'
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bPageSetup = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'Public Function wksGetWorksheetFromWorksheetName(ByVal wkb As Workbook, ByVal sWksName As String) As Worksheet
'
'    Set wksGetWorksheetFromWorksheetName = wkb.Sheets(CStr(wkb.VBProject.VBComponents(sWksName).Properties(7)))
'
'End Function
'
'Public Function bDeleteAllRowsButHeader(ByVal wks As Worksheet, ByVal lHeaderRows As Long) As Boolean
'
'' ---------------------------------------------------------------------------------------------------------------------
'' Comments:     Procedure to manage page setup
''
'' Arguments:    wks             worksheet to delete rows on
''               lHeaderRows     number of header rows to not delete
''
'' Returns:      Boolean         True on success, false on error
''
'' Called by:
''
''
'' Date          Developer       Action
'' ---------------------------------------------------------------------------------------------------------------------
'' 26 Aug 2014   James Boon      Created
''
''
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'    On Error GoTo ErrorHandler
'    Const sSource As String = "bDeleteAllRowsButHeader"
'    Dim bReturn As Boolean
'    bReturn = True
'
'' ---------------------------------------------------------------------------------------------------------------------
'' procedure code
'
'    Dim rngDelete As Range
'
'    ' delete current distribution list for repopulation
'    Set rngDelete = wks.UsedRange.Columns(glFIRST_COLUMN)
'    If rngDelete.Rows.Count > glFIRST_ROW Then
'        Set rngDelete = rngDelete.Resize(rngDelete.Rows.Count - lHeaderRows, glFIRST_COLUMN).Offset(lHeaderRows, 0)
'        rngDelete.EntireRow.Delete 'clear this now so we can repopulate later in the procedure
'    End If
'
'' ---------------------------------------------------------------------------------------------------------------------
'' error handler code
'
'ErrorExit:
'    bDeleteAllRowsButHeader = bReturn
'    Exit Function
'
'ErrorHandler:
'    bReturn = False
'    If bCentralErrorHandler(msMODULE, sSource) Then
'        Stop
'        Resume
'    Else
'        Resume ErrorExit
'    End If
'
'End Function
'
'
'Public Function sRemoveTrailingSingleQuote(ByVal sInput As String) As String
'
'    sRemoveTrailingSingleQuote = Left(sInput, 1) & Replace(sInput, "'", "", 2)
'
'End Function
'
'
'
'
'
Public Function bCopyData(ByVal sPath As String, ByVal sInput As String, ByVal wksOutput As Worksheet) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to open file, copy data into cleared worksheet
'
' Arguments:    sPath           path of file to copy from
'               sFile           file to copy from
'               wksInput        name of worksheet to copy from
'               wksOutput       worksheet to copy to
'
' Returns:      Boolean         True on success, false on error
'
' Called by:    MDailyUpdate
'
'
' Date          Developer       Action
' ---------------------------------------------------------------------------------------------------------------------
' 13 Oct 2014   James Boon      Created
'
'
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

    On Error GoTo ErrorHandler
    Const sSource As String = "bProcessFile"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Dim wkbCopyFrom As Workbook
    Dim wksCopyFrom As Worksheet

    Set wkbCopyFrom = Workbooks.Open(sPath)
    Set wksCopyFrom = wkbCopyFrom.Worksheets(sInput)

    On Error Resume Next
    wksOutput.ShowAllData
    On Error GoTo ErrorHandler
    wksOutput.Cells.Delete
    wksCopyFrom.UsedRange.Cells.Copy Destination:=wksOutput.Cells(1, 1)

    wkbCopyFrom.Close savechanges:=False

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bCopyData = bReturn
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
'
