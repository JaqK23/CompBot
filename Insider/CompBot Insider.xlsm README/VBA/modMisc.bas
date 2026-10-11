Attribute VB_Name = "modMisc"
Option Explicit
Private Const SETTINGS_SHEET As String = "Regional Settings"
Private Const m_DEBUG_MODE   As Boolean = False

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Unmerge Multi-Row Merges
' Description:            Unmerges every merged area spanning more than one row (the selection, or the
'                         whole sheet), leaving the value in the top-left cell
' Macro Expression:       modMisc.UnmergeMultiRowMerges()
'----------------------------------------------------------------------------------------------------
' Promoted into CompBot from A-ZTraining 2026-09-22. The command review found it needed in 7 of 67
' surveyed cases - and every sighting was a STANDALONE competition workbook (Tic Tac Toe's 104
' three-row merges, Battleship's 408, Where's Wally's 142), not an A-Z training case. So the tool was
' living in the one collection that is NOT loaded when it is actually needed.
'
' Unmerges VERTICAL merges only. Single-row merges are left for Merged To Centre Across Selection,
' which replaces them without destroying the layout; vertical merges it cannot.
'
' SCOPE (Jaq, 2026-09-24): a selection of more than one cell, or a single merged cell, limits it to
' the merged blocks the selection touches; a single ordinary cell means the whole sheet. So a block
' can be unmerged in place without taking the rest of the sheet apart.
Public Sub UnmergeMultiRowMerges()
    UnmergeOrCentre True
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Merged To Centre Across Selection
' Description:            Merged Cells changed to Centre Across Selection
' Macro Expression:       modMisc.MergedToCAS()
'----------------------------------------------------------------------------------------------------
' Promoted into CompBot from Jaq's own WiE Robot collection 2026-09-24, with its name, description
' and launch codes (Unmerge, M2C). The partner of Unmerge Multi-Row Merges: SINGLE-row merges are
' unmerged and the text centred across the same cells with Centre Across Selection, so the layout
' looks the same but the cells behave (select, fill, spill). Multi-row merges are left alone.
' Same scope rule as Unmerge Multi-Row Merges: the selection, or the whole sheet.
Public Sub MergedToCAS()
    UnmergeOrCentre False
End Sub

' Purpose: the shared body of the two merge commands. blnMultiRow = True: unmerge merges of more
'          than one row. False: unmerge single-row merges and centre them across selection.
'          Reports on the status bar; no dialogs.
Private Sub UnmergeOrCentre(ByVal blnMultiRow As Boolean)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "UnmergeOrCentre"

    Dim wsTarget    As Worksheet
    Dim rngScan     As Range
    Dim rngCell     As Range
    Dim rngArea     As Range
    Dim rngTargets  As Range
    Dim dicSeen     As Object
    Dim varMerged   As Variant
    Dim blnOnSheet  As Boolean
    Dim lngCount    As Long
    Dim strTitle    As String
    Dim strStatus   As String

    On Error GoTo ErrHandler
    VBAInit

    If blnMultiRow Then strTitle = "Unmerge Multi-Row Merges" Else strTitle = "Merged To Centre Across Selection"

    If Not TypeOf ActiveSheet Is Worksheet Then
        strStatus = strTitle & ": the active sheet is not a worksheet."
        GoTo Cleanup
    End If
    Set wsTarget = ActiveSheet

    Set rngScan = MergeScope(wsTarget, blnOnSheet)

    ' False = nothing merged anywhere in scope. Null = a mix, so keep going.
    varMerged = rngScan.MergeCells
    If Not IsNull(varMerged) Then
        If varMerged = False Then
            strStatus = strTitle & ": no merged cells " & ScopeWords(blnOnSheet) & "."
            GoTo Cleanup
        End If
    End If

    ' Collect first, act second - unmerging inside the scan changes what is being walked. Each
    ' area is keyed by its address, so a block the selection only partly covers is taken whole
    ' and taken once.
    Set dicSeen = CreateObject("Scripting.Dictionary")
    For Each rngCell In rngScan
        If rngCell.MergeCells Then
            Set rngArea = rngCell.MergeArea
            If Not dicSeen.Exists(rngArea.Address) Then
                dicSeen.Add rngArea.Address, True
                If (rngArea.Rows.Count > 1) = blnMultiRow Then
                    lngCount = lngCount + 1
                    If rngTargets Is Nothing Then
                        Set rngTargets = rngArea
                    Else
                        Set rngTargets = Application.Union(rngTargets, rngArea)
                    End If
                End If
            End If
        End If
    Next rngCell

    If rngTargets Is Nothing Then
        If blnMultiRow Then
            strStatus = strTitle & ": no multi-row merges " & ScopeWords(blnOnSheet) & " (single-row ones are for M2C)."
        Else
            strStatus = strTitle & ": no single-row merges " & ScopeWords(blnOnSheet) & " (multi-row ones are for Unmerge Multi-Row Merges)."
        End If
        GoTo Cleanup
    End If

    ' One hit rather than area by area. UnMerge leaves the value in the top-left cell.
    rngTargets.UnMerge
    If Not blnMultiRow Then rngTargets.HorizontalAlignment = xlCenterAcrossSelection

    strStatus = strTitle & ": " & lngCount & " merged block" & IIf(lngCount = 1, "", "s") & " done " & ScopeWords(blnOnSheet) & "."

Cleanup:
    VBAFin                                        ' clears the status bar, so report after it
    Application.StatusBar = strStatus
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strStatus = strTitle & " failed: " & Err.Description & " (see " & LogLocation() & ")."
    Resume Cleanup
End Sub

' Purpose: the cells the merge commands look at. A selection of more than one cell, or a single
'          merged cell, on this sheet: the selection, trimmed to the sheet's content. Otherwise the
'          whole sheet, with UsedRange expanded back out to A1 (these sheets leave a blank spacer
'          column and top rows that UsedRange would miss). blnOnSheet reports which.
Private Function MergeScope(ByVal wsTarget As Worksheet, ByRef blnOnSheet As Boolean) As Range

    Dim rngSheet    As Range
    Dim rngSel      As Range
    Dim varMerged   As Variant
    Dim lngLastRow  As Long
    Dim lngLastCol  As Long
    Dim blnUseSel   As Boolean

    On Error GoTo ErrHandler

    lngLastRow = wsTarget.UsedRange.Row + wsTarget.UsedRange.Rows.Count - 1
    lngLastCol = wsTarget.UsedRange.Column + wsTarget.UsedRange.Columns.Count - 1
    Set rngSheet = wsTarget.Range(wsTarget.Cells(1, 1), wsTarget.Cells(lngLastRow, lngLastCol))

    If TypeName(Selection) = "Range" Then
        Set rngSel = Selection
        If rngSel.Worksheet Is wsTarget Then
            If rngSel.Cells.CountLarge > 1 Then
                blnUseSel = True
            Else
                varMerged = rngSel.MergeCells
                If Not IsNull(varMerged) Then blnUseSel = varMerged
            End If
        End If
    End If

    ' A selection is never widened to the sheet: one wholly outside the content just finds nothing.
    If blnUseSel Then
        Set MergeScope = Application.Intersect(rngSel, rngSheet)
        If MergeScope Is Nothing Then Set MergeScope = rngSel
    Else
        Set MergeScope = rngSheet
        blnOnSheet = True
    End If
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "MergeScope", Err.Number, Err.Description
    Set MergeScope = rngSheet
    blnOnSheet = True
End Function

' Purpose: "in the selection" or "on the sheet", for the status bar.
Private Function ScopeWords(ByVal blnOnSheet As Boolean) As String
    If blnOnSheet Then ScopeWords = "on the sheet" Else ScopeWords = "in the selection"
End Function

'Create a new blank sheet with no gridlines, a header and freeze frames
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Blank Sheet
' Description:            Creates a blank sheet named based on cell value (Sht otherwise)
' Macro Expression:       modMisc.CreateBlankSheet([[ActiveCell]])
' Generated:              01/03/2025 06:05 PM
'----------------------------------------------------------------------------------------------------
Sub CreateBlankSheet(Optional strWS As String)
    Dim wb As Workbook
    Dim wsNew As Worksheet
    Dim rng As Range
    Dim booShtExists As Boolean
    Dim strNewNm As String
    Dim intCnt As Integer
    
    Set wb = ActiveWorkbook
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
    Set wsNew = wb.Sheets.Add(After:=wb.ActiveSheet)
    On Error Resume Next
    wsNew.Name = strNewNm
    On Error GoTo 0
    wsNew.DisplayPageBreaks = False
    ActiveWindow.DisplayGridlines = False
    'make column A thin
    Columns("A:A").ColumnWidth = 1
    'set up a header in row 5
    Set rng = Range("B5:J5")
    rng.HorizontalAlignment = xlCenterAcrossSelection
    rng.Style = "Heading 1"
    rng.Cells(1, 1).value = strWS
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
' Command Name:           Update Settings
' Description:            Updates default settings to those in Regional Settings sheet
' Macro Expression:       modMisc.DefaultSettings()
' Generated:              01/03/2025 10:19 PM
'----------------------------------------------------------------------------------------------------
Sub UpdateSettings()
    Dim decimalSeparator As String
    Dim thousandsSeparator As String
    Dim listSeparator As String
    Dim useSystemSeparators As Boolean
    Dim ws As Worksheet
    
    ' Ensure named ranges exist
    On Error GoTo RangeError
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    
    ' Read values from named ranges
    decimalSeparator = ws.Range("Decimal_Separator").value
    thousandsSeparator = ws.Range("Thousands_Separator").value
    listSeparator = ws.Range("List_Separator").value
    useSystemSeparators = CBool(ws.Range("Use_System_Separators").value)
    
    ' Apply settings
    
    On Error GoTo ApplyError
    'Error checks
    If Len(decimalSeparator) <> 1 And Not useSystemSeparators Then Err.Raise vbObjectError, , "Decimal separator must be a single character"
    If Len(thousandsSeparator) <> 1 And Not useSystemSeparators Then Err.Raise vbObjectError, , "Thousands separator must be a single character"
    
    Application.useSystemSeparators = useSystemSeparators
        
    'Only update this if useSystemSeparators is FALSE
    If Not useSystemSeparators Then
        Application.decimalSeparator = decimalSeparator
        Application.thousandsSeparator = thousandsSeparator
    End If

    ' List Separator: Update formula separators if needed
    ' This requires workarounds since `xlListSeparator` is read-only
    If listSeparator <> Application.International(xlListSeparator) Then
        MsgBox "Please note that the formula separator is automatically set based on decimal separator: " & _
        Application.International(xlListSeparator) & " will be used", vbExclamation
        
    End If

    ' Notify user
    MsgBox "Excel settings updated to current conventions.", vbInformation

    Exit Sub

RangeError:
    MsgBox "Error with loading ranges. Ensure all named ranges exist and are populated.", vbCritical
    Exit Sub
ApplyError:
    MsgBox "Error updating settings. Ensure access is allowed.", vbCritical
    
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Revert Settings
' Description:            Reverts settings to those loaded into the Loaded column.
' Macro Expression:       modMisc.RevertSettings()
' Generated:              2026-05-21 11:46 AM
'----------------------------------------------------------------------------------------------------
Sub RevertSettings()
    Dim decimalSeparator As String
    Dim thousandsSeparator As String
    Dim listSeparator As String
    Dim useSystemSeparators As Boolean
    Dim ws As Worksheet
    
    ' Ensure named ranges exist
    On Error GoTo RangeError
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    
    ' Read values from named ranges
    decimalSeparator = ws.Range("Loaded_Decimal_Separator").value
    thousandsSeparator = ws.Range("Loaded_Thousands_Separator").value
    listSeparator = ws.Range("Loaded_List_Separator").value
    useSystemSeparators = CBool(ws.Range("Loaded_Use_System_Separators").value)
    
    
    ' Apply settings
    
    On Error GoTo ApplyError
    'Error checks
    If Len(decimalSeparator) <> 1 And Not useSystemSeparators Then Err.Raise vbObjectError, , "Decimal separator must be a single character"
    If Len(thousandsSeparator) <> 1 And Not useSystemSeparators Then Err.Raise vbObjectError, , "Thousands separator must be a single character"
    
    Application.useSystemSeparators = useSystemSeparators
        
    'Overwrite this to be sure you're reverting all settings, regardless of useSystemSeparators
    Application.decimalSeparator = decimalSeparator
    Application.thousandsSeparator = thousandsSeparator

    ' List Separator: Update formula separators if needed
    ' This requires workarounds since `xlListSeparator` is read-only
    If listSeparator <> Application.International(xlListSeparator) Then
        MsgBox "Please note that the formula separator is automatically set based on decimal separator: " & _
        Application.International(xlListSeparator) & " will be used", vbExclamation
        
    End If

    ' Notify user
    MsgBox "Excel settings updated to previously loaded conventions.", vbInformation

    Exit Sub

RangeError:
    MsgBox "Error with loading ranges. Ensure all named ranges exist and are populated.", vbCritical
    Exit Sub
ApplyError:
    MsgBox "Error updating settings. Ensure access is allowed.", vbCritical
    
End Sub
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Get Current Settings
' Description:            Get current settings for this computer.
' Macro Expression:       modMisc.GetCurrentSettings()
' Generated:              2026-05-21 09:48 AM
'----------------------------------------------------------------------------------------------------
Sub GetCurrentSettings()
    Dim decimalSeparator As String
    Dim thousandsSeparator As String
    Dim listSeparator As String
    Dim useSystemSeparators As Boolean
    Dim loadedLanguageID As Long
    Dim ws As Worksheet
    
    ' Ensure named ranges exist
    On Error GoTo ErrorHandler
    Set ws = ThisWorkbook.Sheets(SETTINGS_SHEET)
    
     ' Get settings
    useSystemSeparators = Application.useSystemSeparators
    If useSystemSeparators Then
        decimalSeparator = Application.International(xlDecimalSeparator)
        thousandsSeparator = Application.International(xlThousandsSeparator)
    Else
        decimalSeparator = Application.decimalSeparator
        thousandsSeparator = Application.thousandsSeparator
    End If
    
    ' List Separator: Update formula separators if needed
    ' This requires workarounds since `xlListSeparator` is read-only
    listSeparator = Application.International(xlListSeparator)
    
    ' Unable to update - set up for information only
    loadedLanguageID = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
   
   ' Read values from named ranges
    ws.Range("Loaded_Decimal_Separator").value = decimalSeparator
    ws.Range("Loaded_Thousands_Separator").value = thousandsSeparator
    ws.Range("Loaded_List_Separator").value = listSeparator
    ws.Range("Loaded_Use_System_Separators").value = useSystemSeparators
    ws.Range("Loaded_Language_ID").value = loadedLanguageID
    

    ' Notify user
    MsgBox "Current Excel settings loaded to 'Loaded' range.", vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Error getting current regional settings. Have some practice debugging VBA.", vbCritical

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




