Attribute VB_Name = "M_1_EntryPoints"
' =====================================================================================================================
'
' Daily Exceptions Dashboard
' Research Accounts, Finance Division, University of Oxford
' v1.0 - James Boon - Dec 2015
'
' ---------------------------------------------------------------------------------------------------------------------
'
' M_1_EntryPoints
' ---------------
'
' Module for code to deal with entry points (calls from ribbon ui)
' Runs housekeeping at beginning and end of runtime
'
' ---------------------------------------------------------------------------------------------------------------------
'
' Flow Chart
' ----------
'
' MEntryPoints.RibbonEntryPoint
'   bHousekeeping
'       bEnableDisableAddin(disable)
'       bCalculation(manual)
'   M_4_ProjectCode.bButtonHandler
'   bHousekeeping
'       bEnableDisableAddin(enable)
'       bCalculation(automatic)
'
' ---------------------------------------------------------------------------------------------------------------------
'
' Requirements
' ------------
'
' Needs following modules: M_2_ErrorHandler, M_4_ProjectCode, CRuntimeAddin, CRuntimeSettings
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
'Option Private Module

Private Const msMODULE As String = "M_1_EntryPoints"

' =====================================================================================================================
' Module variable declarations
' =====================================================================================================================

Private mclsRuntimeSettings As CRuntimeSettings  'stores runtime settings so we know which state to leave Excel in

' addins to disable
Public Const msGL_CONNECT_ADDIN_NAME_1 As String = "EiSGLConnect"
Public Const msGL_CONNECT_ADDIN_NAME_2 As String = "EiS GL Connect"

' enmumerations
Private Enum mlHousekeeping
    mlHousekeepingBeginning
    mlHousekeepingEnd
End Enum
Private Enum mlCalculation
    mlCalculationAutomatic
    mlCalculationManual
End Enum
Private Enum mlEnableDisable
    mlDisable
    mlEnable
End Enum


' =====================================================================================================================
' Subroutintes
' =====================================================================================================================

Public Sub RibbonEntryPoint_DASH( _
    ircControl As IRibbonControl, _
    Optional sID As String, _
    Optional iIndex As Integer _
)

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to handle entry point from ribbon button
' Arguments:    ircControl      Control number from ribbon
' Returns:      none
'
' Date          Developer       Action
' ---------------------------------------------------------------------------------------------------------------------
' 17 Dec 2013   James Boon      Created
' 02 Jun 2014   James Boon      Cater for no active workbook (failed on application.calculation = xlcalculationmanual)
'
'
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

    On Error GoTo ErrorHandler
    Const sSource As String = "RibbonEntryPoint_DASH"
    
' ---------------------------------------------------------------------------------------------------------------------
' declarations
    
' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    If Not bHousekeeping(mlHousekeepingBeginning) Then Err.Raise Number:=glHANDLED_ERROR
    If Not bButtonHandler(ircControl) Then Err.Raise Number:=glHANDLED_ERROR
   

    
' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    If Not bHousekeeping(mlHousekeepingEnd) Then Err.Raise Number:=glHANDLED_ERROR
    Exit Sub

ErrorHandler:
    If bCentralErrorHandler(msMODULE, sSource, , True) Then
        Stop
        Resume
    Else
        Resume ErrorExit
    End If

End Sub

' =====================================================================================================================
' Functions
' =====================================================================================================================

Private Function bHousekeeping( _
    ByVal lBeginningOrEnd As mlHousekeeping _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to handle housekeeping tasks at beginning/end of add-in
'               We know that the GL Connect Add-in causes a conflict, so we temporarily suspend it during processing
' Arguments:    lBeginningorEnd             Tells us whether to prepare for add-in, or tidy up at end of add-in
' Returns:      Boolean                     True on success, false on error
' Called by:    RibbonEntryPoint
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 15 Dec 2014   James Boon      1.2         Created
' 16 Dec 2014   James Boon      1.2         Using clsRuntimeSettings to hold initial application settings
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bHousekeeping"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    Select Case lBeginningOrEnd
        
        Case mlHousekeepingBeginning
        
            Set mclsRuntimeSettings = New CRuntimeSettings
            Application.ScreenUpdating = False
            InitGlobals
            If Not bEnableDisableAddin(mlDisable, msGL_CONNECT_ADDIN_NAME_1) Then Err.Raise Number:=glHANDLED_ERROR
            If Not bEnableDisableAddin(mlDisable, msGL_CONNECT_ADDIN_NAME_2) Then Err.Raise Number:=glHANDLED_ERROR
            If Not bCalculation(xlCalculationManual) Then Err.Raise Number:=glHANDLED_ERROR
        
        Case mlHousekeepingEnd
            
            If Not bEnableDisableAddin(mlEnable, msGL_CONNECT_ADDIN_NAME_1) Then Err.Raise Number:=glHANDLED_ERROR
            If Not bEnableDisableAddin(mlEnable, msGL_CONNECT_ADDIN_NAME_2) Then Err.Raise Number:=glHANDLED_ERROR
            If Not bCalculation(xlCalculationAutomatic) Then Err.Raise Number:=glHANDLED_ERROR
            Application.StatusBar = False
            Application.ScreenUpdating = True
            Set mclsRuntimeSettings = Nothing
    
    End Select

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bHousekeeping = bReturn
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

Private Function bCalculation( _
    ByVal lState As XlCalculation _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to turn calculation on or off
' Arguments:    lState                      Tells us whether to turn calculation on or off
' Returns:      Boolean                     True on success, false on error
' Called by:    bHousekeeping
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 15 Dec 2014   James Boon      1.2         Created
' 16 Dec 2014   James Boon      1.2         Ensures we actually have something to do before we try to change state
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bCalculation"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim wkb As Excel.Workbook

' ---------------------------------------------------------------------------------------------------------------------
' procedure code
    
    If Application.ActiveWorkbook Is Nothing Then Set wkb = Workbooks.Add 'prevents error on reading application.calculation
    If lState <> Application.Calculation And Not (mclsRuntimeSettings.ManualCalculation) Then
        Application.Calculation = lState
    End If

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bCalculation = bReturn
    On Error Resume Next
    wkb.Close savechanges:=False
    Set wkb = Nothing
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

Private Function bEnableDisableAddin( _
    ByVal lEnableDisable As mlEnableDisable, _
    ByVal sAddinName As String _
) As Boolean

' ---------------------------------------------------------------------------------------------------------------------
' Comments:     Procedure to enable/disable any add-in
' Arguments:    lEnableDisable              defines which way the routine is being run
'               sAddinName                  addin to enable/disable
' Returns:      Boolean                     True on success, false on error
' Called by:    bHousekeeping
' ---------------------------------------------------------------------------------------------------------------------
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 11 Dec 2014   James Boon      1.1         Created
' 15 Dec 2014   James Boon      1.2         Altered to use enum instead of string in parameter
'                                           Moved to entry points
' 16 Dec 2014   James Boon      1.2         Made more general (taken name of addin out of function)
' ---------------------------------------------------------------------------------------------------------------------
' error handler code
    
    On Error GoTo ErrorHandler
    Const sSource As String = "bEnableDisableAddin"
    Dim bReturn As Boolean
    bReturn = True

' ---------------------------------------------------------------------------------------------------------------------
' declarations

    Dim aiToChange As Excel.AddIn

' ---------------------------------------------------------------------------------------------------------------------
' procedure code

    On Error Resume Next
        Set aiToChange = Application.Addins(sAddinName)
    On Error GoTo ErrorHandler
    
    If Not (aiToChange Is Nothing) Then
    
        Select Case lEnableDisable
            
            Case mlDisable
                aiToChange.Installed = False
                        
            Case mlEnable
                If mclsRuntimeSettings.Addins(sAddinName).Status Then aiToChange.Installed = True
                
        End Select
                        
    End If

' ---------------------------------------------------------------------------------------------------------------------
' error handler code

ErrorExit:
    bEnableDisableAddin = bReturn
    On Error Resume Next
    Set aiToChange = Nothing
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


