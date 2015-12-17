Attribute VB_Name = "M_3_GlobalVariables"
' =====================================================================================================================
'
' Daily Exceptions Dashboard
' Research Accounts, Finance Division, University of Oxford
' v1.0 - James Boon - Dec 2015
'
' ---------------------------------------------------------------------------------------------------------------------
'
' M_3_GlobalVariables
' -------------------
'
' Module for code for global variables
'
' ---------------------------------------------------------------------------------------------------------------------
'
' Version Control
' ---------------
'
' Date          Developer       Version     Action
' ---------------------------------------------------------------------------------------------------------------------
' 02 Dec 2015   James Boon      1.0         Updating from v0.7 to tidy up coding, add user forms, turn into addin
'
' =====================================================================================================================

Option Explicit
Option Private Module

' Application global variables
' ----------------------------

' Session run data
Public gsTimeStamp As String
Public gdteTimeRan As Date

' Tolerances
Public Const glBURDEN_BUDGET_VALIDATION_TOLERANCE As Long = 1


' text file settings
Public Const gsSETTINGS_FILE As String = "Dashboard_Settings.txt"
Public Const gsOUTPUT_UNEXPECTED_CASE As String = "Unexpected value"

Public gvaReportTitles() As Variant
Public gvaReportRanges() As Variant

Public Enum glReportType
    glReportTypeCostToRevenue
    glReportTypeBurdenValidation
    glReportTypeBurdeningBug
    glReportTypeBurdenBudgetValidation
    glReportTypeBurdenCostCodeMismatch
    glReportTypeFecOhReconciliationData
End Enum

Public Enum glBudgetItemType
    glBudgetLineTypeDirectCost
    glBudgetLineTypeDirectlyAllocatedCost
    glBudgetLineTypeIndirectCost
    glBudgetLineTypePriceAdjustment
    glBudgetLineTypeOverhead
    glBudgetLineTypeOther
End Enum

' Generic Constants
' -----------------
Public Const glFIRST_SHEET As Long = 1
Public Const glFIRST_COLUMN As Long = 1
Public Const glFIRST_ROW As Long = 1
Public Const glONE_PAGE_WIDE As Long = 1
Public Const glADD_ONE_BECAUSE_ARRAY_INDEX_STARTS_WITH_ZERO As Long = 1
Public Const glNEXT_ROW As Long = 1
Public Const glNEXT_COLUMN As Long = 1
Public Const glPREVIOUS_ROW As Long = -1
Public Const glPREVIOUS_COLUMN As Long = -1
Public Const gsPERSONAL_WORKBOOK_NAME As String = "PERSONAL.XLSB"
Public Const gsUSER_CANCEL As String = "UserCancel"
Public Const gsXLSX_EXTENSION As String = ".xlsx"
Public Const gsERROR_LOG As String = " - Error.log"
Public Const gsPREFIX_BACKUP_FILE As String = "backup-"


' Oxford Brand Colours (maybe make this an enumaration or even a class)
' --------------------
Public Const glOxfordPink = 13354219        'RGB(235, 196, 203)
Public Const glLightGreen = 11525070        'RGB(206, 219, 175)
Public Const glLightTurquoise = 12833468    'RGB(188, 210, 195)
Public Const glOrange = 7659251             'RGB(243, 222, 116)
Public Const glGrey = 14278368              'RGB(224, 222, 217)
Public Const glWhite = 16777215             'RGB(255, 255, 255)
Public Const glOxfordBlue = 4661504         'RGB(0, 33, 71)


Public Sub InitGlobals()
    
    If Not bTimeStamp(gsTimeStamp, gdteTimeRan) Then Err.Raise Number:=glHANDLED_ERROR
    
    gvaReportTitles() = Array( _
        "Cost to Revenue", _
        "Burden Validation", _
        "Burdening Bug", _
        "Burden to Budget Validation", _
        "Burden Cost Code Mismatch", _
        "FEC OH Reconciliation Data" _
    )
    
    gvaReportRanges() = Array( _
        "rngCostToRevenue", _
        "rngBurdenValidation", _
        "rngBurdeningBug", _
        "rngBurdenBudgetValidation", _
        "rngBurdenCostCodeMismatch", _
        "rngFecOhReconciliationData" _
    )
    
End Sub
