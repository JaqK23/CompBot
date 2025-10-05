Attribute VB_Name = "modMisc"
Option Explicit

'Create a new blank sheet with no gridlines, a header and freeze frames
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Blank Sheet
' Description:            Creates a blank sheet named based on cell value (Sht otherwise)
' Macro Expression:       modMisc.CreateBlankSheet([[ActiveCell]])
' Generated:              01/03/2025 06:05 PM
'----------------------------------------------------------------------------------------------------
Sub CreateBlankSheet(Optional strWS As String)
    Dim WB As Workbook
    Dim WSNew As Worksheet
    Dim rng As Range
    Dim booShtExists As Boolean
    Dim strNewNm As String
    Dim intCnt As Integer
    
    Set WB = ActiveWorkbook
    If IsMissing(strWS) Or strWS = "" Then strWS = "Sht"
    
    ' Loop to determine the next available name
    booShtExists = True
    intCnt = 0
    Do While booShtExists
        If intCnt = 0 Then
            strNewNm = strWS
        Else
            strNewNm = strWS & intCnt
        End If
        
        ' Check if the name exists
        booShtExists = SheetExists(strNewNm)
        
        If Not booShtExists Then Exit Do
        
        intCnt = intCnt + 1
    Loop
    
    'worksheet exists
    Set WSNew = WB.Sheets.Add(After:=WB.ActiveSheet)
    On Error Resume Next
    WSNew.Name = strNewNm
    On Error GoTo 0
    WSNew.DisplayPageBreaks = False
    ActiveWindow.DisplayGridlines = False
    'make column A thin
    Columns("A:A").ColumnWidth = 1
    'set up a header in row 5
    Set rng = Range("B5:J5")
    rng.HorizontalAlignment = xlCenterAcrossSelection
    rng.Style = "Heading 1"
    rng.Cells(1, 1).Value = strWS
    'freeze panes below header
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 5
        .FreezePanes = True
    End With
        
End Sub

'Update Default Settings to those specified
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Default Settings
' Description:            Updates default settings to those in Default Settings sheet
' Macro Expression:       modMisc.DefaultSettings()
' Generated:              01/03/2025 10:19 PM
'----------------------------------------------------------------------------------------------------
Sub DefaultSettings()
    Dim decimalSeparator As String
    Dim thousandsSeparator As String
    Dim listSeparator As String
    Dim useSystemSeparators As Boolean
    Dim targetLanguage As String
    Dim WS As Worksheet
    
    ' Ensure named ranges exist
    On Error GoTo ErrorHandler
    Set WS = ThisWorkbook.Sheets("Default Settings")
    
    ' Read values from named ranges
    decimalSeparator = WS.Range("Decimal_Separator").Value
    thousandsSeparator = WS.Range("Thousands_Separator").Value
    listSeparator = WS.Range("List_Separator").Value
    useSystemSeparators = WS.Range("Use_System_Separators").Value
    targetLanguage = WS.Range("Language").Value
    
    ' Apply settings
    Application.useSystemSeparators = useSystemSeparators
    
    If Not useSystemSeparators Then
        Application.decimalSeparator = decimalSeparator
        Application.thousandsSeparator = thousandsSeparator
    End If

    ' List Separator: Update formula separators if needed
    ' This requires workarounds since `xlListSeparator` is read-only
    If listSeparator <> Application.International(xlListSeparator) Then
        MsgBox "Please ensure the regional settings in your system reflect the desired list separator: " & listSeparator, vbExclamation
        
    End If

    ' Notify user
    MsgBox "Excel settings updated to " & targetLanguage & " conventions.", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error updating regional settings. Ensure all named ranges exist and are populated.", vbCritical

End Sub

'Toggle calculation mode between manual and automatic
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Toggle Calculation Mode
' Description:            Toggles calculation mode and places current mode notice in StatusBar
' Macro Expression:       modMisc.ToggleCalculationMode()
' Generated:              01/08/2025 01:50 PM
'----------------------------------------------------------------------------------------------------
Sub ToggleCalculationMode()
    If Application.Calculation = xlCalculationAutomatic Then
        Application.Calculation = xlCalculationManual
        Application.StatusBar = "Calculation mode is now set to Manual."
    Else
        Application.Calculation = xlCalculationAutomatic
        Application.StatusBar = "Calculation mode is now set to Automatic."
    End If
End Sub

'Toggle iterative calculation
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Toggle Iterative Calculation
' Description:            Toggles iterative calculation and sets status in status bar
' Macro Expression:       modMisc.ToggleIterativeCalculation()
' Generated:              01/08/2025 01:50 PM
'----------------------------------------------------------------------------------------------------
Sub ToggleIterativeCalculation()
    If Application.Iteration Then
        Application.Iteration = False
        Application.StatusBar = "Iterative calculation is disabled."
    Else
        Application.Iteration = True
        Application.MaxIterations = 1000 ' Adjust as needed
        Application.MaxChange = 0.001 ' Adjust as needed
        Application.StatusBar = "Iterative calculation is enabled."
    End If
End Sub

