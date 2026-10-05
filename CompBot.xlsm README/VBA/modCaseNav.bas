Attribute VB_Name = "modCaseNav"
Option Explicit

'==============================================================================
' modCaseNav - selection helpers for working a case, promoted into CompBot from
' A-ZTraining on 2026-09-22.
'
' WHY THEY MOVED: the command review scored every command against 67 surveyed
' cases. "Exclude Examples" was needed in 23 of them and "Select Example And
' Questions" in 14 - the 2nd and 5th most-needed commands of Jaq's own, and both
' were living in the training collection rather than the competition one.
'
' WHAT CHANGED ON THE WAY IN - neither is a straight copy:
'   1. Both now go through modCaseSetup's shared example-word list, so they work
'      on French ("Exemple"), Brazilian Portuguese ("Exemplo") and the older
'      English "Sample" cases, not only "Example".
'   2. Neither assumes there is exactly ONE worked example row. The A-Z originals
'      both hardcode "drop two rows" / "the row below"; several 2026 cases carry
'      two or three worked rows per level (Example3a / 3b). These walk the rows
'      instead and stop when they run out of example and blank rows.
'
' 2026-09-24: Save Answer To Bonus 1-5 (B1-B5) promoted from A-ZTraining too (Jaq),
' so the side-column bonus flow (Link Side Cell, RDT, a bonus formula, then B#)
' lives in one collection. A-ZTraining keeps its own copies.
'==============================================================================

' --- CONSTANTS (module) ---
Private Const m_DEBUG_MODE   As Boolean = False
Private Const m_LABEL_COL    As Long = 2     ' cases put Bonus / Example labels in column B
Private Const m_MAX_EXAMPLES As Long = 12    ' sanity stop when walking example rows
Private Const m_MEWC_GREEN   As Long = 3631104   ' the MEWC answer-cell fill
Private Const m_BONUS_SHEET  As String = "B"     ' Create Bonus Sheet's sheet
Private Const m_MAX_STATUS   As Long = 250       ' Application.StatusBar rejects long strings
Private Const m_MAX_HEADER_UP As Long = 30       ' how far above "Bonus N" its "Answer" header can be


'------------------------------------------------------------------------------
' Trim the worked-example rows off the top of the current selection, leaving the
' real questions selected.
'
' The A-Z original dropped exactly two rows - one example plus one blank spacer.
' This walks instead: it keeps dropping the top row while that row is an example
' label or is blank, so one example, three examples, or a missing spacer all
' behave. It falls back to the original two-row behaviour if it cannot read the
' label column at all, so it is never worse than what it replaces.
'------------------------------------------------------------------------------
Public Sub ExcludeExamples()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ExcludeExamples"

    Dim rngSel     As Range
    Dim ws         As Worksheet
    Dim lngRows    As Long
    Dim lngCols    As Long
    Dim lngDrop    As Long
    Dim lngRow     As Long
    Dim varLabel   As Variant
    Dim lngErrNum  As Long
    Dim strErrDesc As String

    On Error GoTo ErrHandler

    Set rngSel = Selection
    If rngSel Is Nothing Then GoTo Cleanup

    Set ws = rngSel.Worksheet
    lngRows = rngSel.Rows.Count
    lngCols = rngSel.Columns.Count

    If lngRows <= 2 Then
        Application.StatusBar = "Exclude Examples: select more than the example rows first"
        GoTo Cleanup
    End If

    ' Walk down from the top of the selection: drop example-label rows and blank rows.
    For lngRow = 1 To m_MAX_EXAMPLES
        If lngRow >= lngRows Then Exit For
        varLabel = ws.Cells(rngSel.Row + lngRow - 1, m_LABEL_COL).Value2
        If modCaseSetup.IsExampleLabel(varLabel) Then
            lngDrop = lngRow
        ElseIf Len(Trim$(CStr(varLabel & ""))) = 0 And lngDrop > 0 Then
            ' a blank spacer directly under an example row
            lngDrop = lngRow
        ElseIf lngDrop > 0 Then
            Exit For
        End If
    Next lngRow

    ' Nothing recognisable - behave exactly as the A-Z original did.
    If lngDrop = 0 Then lngDrop = 2

    If lngDrop >= lngRows Then
        Application.StatusBar = "Exclude Examples: nothing left once the example rows are dropped"
        GoTo Cleanup
    End If

    rngSel.Offset(lngDrop, 0).Resize(lngRows - lngDrop, lngCols).Select

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Capture first: LogError resets Err.
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    Application.StatusBar = "Exclude Examples failed: " & lngErrNum & " - " & strErrDesc
    LogError PROC_NAME, lngErrNum, strErrDesc
    Resume Cleanup

End Sub


'------------------------------------------------------------------------------
' Select from the worked-example row above the cursor down to the end of the
' level's questions.
'
' The A-Z original searched for the literal string "Example", so it found nothing
' on a Portuguese or French case. This tries every accepted example word.
'------------------------------------------------------------------------------
Public Sub SelectExampleAndQuestions()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "SelectExampleAndQuestions"

    Dim ws         As Worksheet
    Dim rngFound   As Range
    Dim rngBest    As Range
    Dim varWords   As Variant
    Dim varWord    As Variant
    Dim rngAfter   As Range
    Dim lngErrNum  As Long
    Dim strErrDesc As String

    On Error GoTo ErrHandler

    Set ws = ActiveSheet
    Selection.Offset(1, 0).Select
    Set rngAfter = Selection.Cells(1, 1)

    ' Search upwards for each accepted example word and keep the LOWEST hit, which is
    ' the example row belonging to this level rather than one from a level above.
    varWords = modCaseSetup.ExampleWords()
    For Each varWord In varWords
        Set rngFound = FindLabelAbove(ws, CStr(varWord), rngAfter)
        If Not rngFound Is Nothing Then
            If rngBest Is Nothing Then
                Set rngBest = rngFound
            ElseIf rngFound.Row > rngBest.Row Then
                Set rngBest = rngFound
            End If
        End If
    Next varWord

    If rngBest Is Nothing Then
        Application.StatusBar = "Select Example And Questions: no example row found above the selection"
        LogError PROC_NAME, 0, "no example label found on " & ws.Name
        GoTo Cleanup
    End If

    rngBest.Select
    ws.Range(Selection, Selection.End(xlDown).End(xlDown)).Select
    ActiveWindow.ScrollRow = Application.WorksheetFunction.Max(Selection.Row - 10, 1)
    ActiveWindow.ScrollColumn = 1

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    Application.StatusBar = "Select Example And Questions failed: " & lngErrNum & " - " & strErrDesc
    LogError PROC_NAME, lngErrNum, strErrDesc
    Resume Cleanup

End Sub


'------------------------------------------------------------------------------
' Nearest cell above rngAfter whose text is a real example LABEL for strWord.
' Up to three hits are tried, each search continuing on from the last, so a
' sentence that merely mentions the word is stepped over rather than ending the
' search - the same retry the A-Z original used, kept deliberately.
'------------------------------------------------------------------------------
Private Function FindLabelAbove(ByVal ws As Worksheet, _
                                ByVal strWord As String, _
                                ByVal rngAfter As Range) As Range

    Dim rngHit As Range
    Dim lngTry As Long

    On Error GoTo ErrHandler

    ' SearchOrder, MatchCase and SearchFormat are passed EXPLICITLY. Range.Find inherits
    ' any it is not given from the last search made anywhere in the session - including
    ' the user's own Find dialog - so leaving them out makes the result depend on history.
    ' Measured 2026-09-22: with SearchOrder inherited as xlByColumns, "previous" from a
    ' cell in Level 2 walked UP ITS OWN COLUMN and then round to the BOTTOM of column B,
    ' so Select Example And Questions found Level 3's example row and selected the wrong
    ' level entirely. xlByRows is what "above" means here. SearchFormat is the other
    ' classic: inherit a format filter from the user and Find quietly matches nothing.
    Set rngHit = Nothing
    For lngTry = 1 To 3
        If rngHit Is Nothing Then
            Set rngHit = ws.Cells.Find(What:=strWord, After:=rngAfter, LookIn:=xlValues, _
                                       LookAt:=xlPart, SearchOrder:=xlByRows, _
                                       SearchDirection:=xlPrevious, MatchCase:=False, _
                                       SearchFormat:=False)
        Else
            Set rngHit = ws.Cells.Find(What:=strWord, After:=rngHit, LookIn:=xlValues, _
                                       LookAt:=xlPart, SearchOrder:=xlByRows, _
                                       SearchDirection:=xlPrevious, MatchCase:=False, _
                                       SearchFormat:=False)
        End If
        If rngHit Is Nothing Then Exit For
        If modCaseSetup.IsExampleLabel(rngHit.Value2) Then
            Set FindLabelAbove = rngHit
            Exit Function
        End If
    Next lngTry

    Exit Function

ErrHandler:
    Set FindLabelAbove = Nothing

End Function


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Save Answer To Bonus 1 (... 5)
' Description:            Links the active cell (any sheet) into the Bonus N answer cell, goes there and
'                         copies it for the submission site.
' Macro Expression:       modCaseNav.SaveAnswerToBonus(1)
' Generated:              2026-09-24, promoted from A-ZTraining modNavigation
'----------------------------------------------------------------------------------------------------
' Target: sheet B when it holds "Bonus N" (Create Bonus Sheet links Case's bonus answers to B),
' otherwise the sheet holding "Bonus Questions" (normally Case). The answer cell is the MEWC green
' cell on the "Bonus N" row; failing that (another green), the cell on that row under the nearest
' "Answer" header above it. A live link is written, then the answer cell is selected and copied,
' ready to paste into the submission site. Nothing is written if the target cannot be found.
' Changed on the way in: every Find passes SearchOrder / MatchCase / SearchFormat explicitly (they are
' inherited from the last search otherwise, see FindLabelAbove), and failures are logged.
Public Sub SaveAnswerToBonus(ByVal lngBonus As Long)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "SaveAnswerToBonus"

    Dim rngSrc     As Range
    Dim wsTarget   As Worksheet
    Dim rngLabel   As Range
    Dim rngRow     As Range
    Dim rngCell    As Range
    Dim rngAns     As Range
    Dim rngLast    As Range
    Dim varShown   As Variant
    Dim strShown   As String
    Dim strLabel   As String
    Dim strAltLabel As String
    Dim strTitle   As String
    Dim lngCol     As Long
    Dim lngErrNum  As Long
    Dim strErrDesc As String

    On Error GoTo ErrHandler

    strTitle = "Save Answer To Bonus " & lngBonus
    If TypeName(Selection) <> "Range" Then
        Application.StatusBar = strTitle & ": select the answer cell first"
        GoTo Cleanup
    End If
    Set rngSrc = ActiveCell
    strLabel = "Bonus " & CStr(lngBonus)
    ' 2026-09-25: the 2025 MEWC qualification rounds label their bonuses "Bonus A" to "Bonus E",
    ' and B1 silently did nothing there. Bonus N also matches the Nth letter; the number wins.
    strAltLabel = "Bonus " & Chr$(64 + lngBonus)

    ' Sheet B first, else the sheet with the bonus block. Probe: Worksheets() raises when B is missing.
    On Error Resume Next
    Set wsTarget = ActiveWorkbook.Worksheets(m_BONUS_SHEET)
    On Error GoTo ErrHandler
    If Not wsTarget Is Nothing Then
        If FindWhole(wsTarget, strLabel) Is Nothing Then
            If FindWhole(wsTarget, strAltLabel) Is Nothing Then Set wsTarget = Nothing
        End If
    End If
    If wsTarget Is Nothing Then Set wsTarget = SheetWithValue("Bonus Questions")
    If wsTarget Is Nothing Then
        Application.StatusBar = strTitle & ": no 'Bonus Questions' block found"
        GoTo Cleanup
    End If

    Set rngLabel = FindWhole(wsTarget, strLabel)
    If rngLabel Is Nothing Then
        Set rngLabel = FindWhole(wsTarget, strAltLabel)
        If Not rngLabel Is Nothing Then strLabel = strAltLabel
    End If
    If rngLabel Is Nothing Then
        Application.StatusBar = strTitle & ": '" & strLabel & "' / '" & strAltLabel & "' not found on " & wsTarget.Name
        GoTo Cleanup
    End If

    ' The green answer cell on that row, up to the last column with content (UsedRange can run to XFD
    ' when whole rows are formatted, which would mean thousands of reads).
    Set rngLast = wsTarget.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, _
                                      SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    If Not rngLast Is Nothing Then
        Set rngRow = wsTarget.Range(wsTarget.Cells(rngLabel.Row, 1), wsTarget.Cells(rngLabel.Row, rngLast.Column))
    End If
    If Not rngRow Is Nothing Then
        For Each rngCell In rngRow.Cells
            If rngCell.Interior.Color = m_MEWC_GREEN Then
                Set rngAns = rngCell
                Exit For
            End If
        Next rngCell
    End If
    ' Fallback (Jaq, 2026-09-24): no MEWC green on the row, because the case (a local chapter, or
    ' the tutorial's demo sheet) uses another green. Take the column under the nearest "Answer"
    ' header above the label, the route Save From Example uses.
    If rngAns Is Nothing Then
        lngCol = BonusAnswerColumn(wsTarget, rngLabel.Row, rngLabel.Column)
        If lngCol > 0 And lngCol <> rngLabel.Column Then Set rngAns = wsTarget.Cells(rngLabel.Row, lngCol)
    End If
    If rngAns Is Nothing Then
        Application.StatusBar = strTitle & ": no green answer cell or 'Answer' header for row " & rngLabel.Row & _
                                " of " & wsTarget.Name
        GoTo Cleanup
    End If

    If rngSrc.Worksheet Is wsTarget And rngSrc.Address = rngAns.Address Then
        Application.StatusBar = strTitle & ": the active cell is the answer cell - select the result first"
        GoTo Cleanup
    End If

    If rngSrc.Worksheet Is wsTarget Then
        rngAns.Formula2 = "=" & rngSrc.Address(True, True)
    Else
        rngAns.Formula2 = "='" & Replace(rngSrc.Worksheet.Name, "'", "''") & "'!" & rngSrc.Address(True, True)
    End If
    rngAns.Calculate

    ' Going to the answer cell and copying it is the purpose of the command.
    wsTarget.Activate
    rngAns.Select
    ShowRange rngAns               ' .Select alone does not scroll while ScreenUpdating is off
    rngAns.Copy
    ' The value itself (.Text shows #### in a narrow column), and the StatusBar rejects long text.
    varShown = rngAns.Value2
    If IsError(varShown) Then strShown = "#error" Else strShown = CStr(varShown)
    Application.StatusBar = Left$(strLabel & " = " & strShown & " (from " & rngSrc.Worksheet.Name & "!" & _
                                  rngSrc.Address(False, False) & ") - copied", m_MAX_STATUS)

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Capture first: LogError resets Err.
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    Application.StatusBar = strTitle & " failed: " & lngErrNum & " - " & strErrDesc
    LogError PROC_NAME, lngErrNum, strErrDesc
    Resume Cleanup

End Sub


'------------------------------------------------------------------------------
' The column of the bonus block's "Answer" header above row lngRow, looking at
' most m_MAX_HEADER_UP rows up; 0 if there is none. The walk stays INSIDE the
' bonus block: it stops at the first row whose label (column lngLabelCol) is
' neither blank nor a "Bonus..." label, so a level's own "Answer" header above
' the block is never taken (it would link the bonus into a question cell).
' modCaseSetup.AnswerColumn reads each row as an array, so no Range.Find.
'------------------------------------------------------------------------------
Private Function BonusAnswerColumn(ByVal ws As Worksheet, ByVal lngRow As Long, _
                                   ByVal lngLabelCol As Long) As Long

    Dim lngScan  As Long
    Dim lngTop   As Long
    Dim lngCol   As Long
    Dim strLabel As String

    On Error GoTo ErrHandler

    lngTop = lngRow - m_MAX_HEADER_UP
    If lngTop < 1 Then lngTop = 1
    For lngScan = lngRow - 1 To lngTop Step -1
        lngCol = modCaseSetup.AnswerColumn(ws, lngScan)
        If lngCol > 0 Then
            BonusAnswerColumn = lngCol
            Exit Function
        End If
        strLabel = Trim$(CStr(ws.Cells(lngScan, lngLabelCol).Value2))
        If Len(strLabel) > 0 And StrComp(Left$(strLabel, 5), "Bonus", vbTextCompare) <> 0 Then Exit Function
    Next lngScan

    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    BonusAnswerColumn = 0

End Function


'------------------------------------------------------------------------------
' The cell on ws whose whole value is strValue, or Nothing. Find's sticky
' options are all passed, so the user's last search cannot change the result.
'------------------------------------------------------------------------------
Private Function FindWhole(ByVal ws As Worksheet, ByVal strValue As String) As Range

    On Error GoTo ErrHandler

    Set FindWhole = ws.Cells.Find(What:=strValue, LookIn:=xlValues, LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                  MatchCase:=False, SearchFormat:=False)
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Set FindWhole = Nothing

End Function


'------------------------------------------------------------------------------
' The worksheet holding a cell whose whole value is strValue: the active sheet
' first, then the other worksheets of the active workbook in tab order.
' Nothing if none has it.
'------------------------------------------------------------------------------
Private Function SheetWithValue(ByVal strValue As String) As Worksheet

    Dim wsTry As Worksheet

    On Error GoTo ErrHandler

    If TypeOf ActiveSheet Is Worksheet Then
        If Not FindWhole(ActiveSheet, strValue) Is Nothing Then
            Set SheetWithValue = ActiveSheet
            GoTo Cleanup
        End If
    End If
    For Each wsTry In ActiveWorkbook.Worksheets
        If Not FindWhole(wsTry, strValue) Is Nothing Then
            Set SheetWithValue = wsTry
            GoTo Cleanup
        End If
    Next wsTry

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Set SheetWithValue = Nothing
    Resume Cleanup

End Function









