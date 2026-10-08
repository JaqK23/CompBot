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


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Level Inputs   (launch codes LI, LLI)
' Macro Expression:       modCaseNav.LoadLevelInputs({{Levels}})
'----------------------------------------------------------------------------------------------------
' Purpose: put one or more levels' inputs into the active cell as a formula, wherever you are -
'          usually the bonus sheet (GitHub #3, Ben de Leon / Jaq, 2026-10-08). Type the levels the
'          way you would say them: "4", "4-5", "3,5-6"; any number of levels, so 5-level, 8-level
'          and 10-level cases all work. Left blank, it uses the level of the sheet you are on (L06).
'          The INPUTS spill as an array of their own (two columns right of the active cell, one row
'          down), ready to work on, with their headers above for information; Game and Level spill
'          as a separate array from below the active cell. Input columns blank for every chosen
'          level are left out, and empty inputs show blank, not 0. The names are made by Create Case Inputs Sheet (CIS, part of Full Setup
'          Case): each is that level's rows on CaseInputs - Game, Level and the input columns that
'          level uses. Never overwrites: the four anchor cells must be empty. Every outcome goes to
'          the status bar.
Public Sub LoadLevelInputs(Optional ByVal varLevels As Variant)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME     As String = "LoadLevelInputs"
    Const FAILED_PREFIX As String = "LEVEL INPUTS FAILED: "
    Const HEADER_ROW    As Long = 2     ' CaseInputs' header row (modCaseSetup.DetailedInputs)
    Const IS_IN_LIST    As String = "IsInList_byErikOehm"   ' CompBot lambda: (array, list) -> TRUE/FALSE each

    Dim wb          As Workbook
    Dim rngTarget   As Range
    Dim rngInfoHdr  As Range
    Dim rngInfo     As Range
    Dim rngInHdr    As Range
    Dim rngData     As Range
    Dim rngTable    As Range
    Dim strTable    As String
    Dim colLevels   As Collection
    Dim varLevel    As Variant
    Dim strLevels   As String
    Dim strName     As String
    Dim strArgs     As String
    Dim strStack    As String
    Dim strInputs   As String
    Dim strHeaderRow As String
    Dim strShown    As String
    Dim strStatus   As String
    Dim strWhy      As String
    Dim strBlocked  As String
    Dim strInLevels As String
    Dim strLoaded   As String

    On Error GoTo ErrHandler

    If Not IsMissing(varLevels) Then
        If Not IsError(varLevels) Then strLevels = Trim$(CStr(varLevels))
    End If

    Set wb = ActiveWorkbook
    If wb Is Nothing Or TypeName(Selection) <> "Range" Then
        strStatus = FAILED_PREFIX & "select the empty cell the inputs should go in."
        GoTo Cleanup
    End If
    Set rngTarget = ActiveCell

    ' Blank: the level of the sheet you are on.
    If Len(strLevels) = 0 Then
        If modCaseSetup.IsLevelSheetName(rngTarget.Worksheet.Name) Then
            strLevels = CStr(Val(Mid$(rngTarget.Worksheet.Name, 2)))
        Else
            strStatus = FAILED_PREFIX & "type the level(s), e.g. 4 or 4-5 or 3,5-6 (blank only works on a level sheet)."
            GoTo Cleanup
        End If
    End If

    Set colLevels = ParseLevelList(strLevels, strWhy)
    If colLevels Is Nothing Then
        strStatus = FAILED_PREFIX & """" & strLevels & """ - " & strWhy
        GoTo Cleanup
    End If

    ' The whole CaseInputs table the level names cover: Game in its first column, Level in its
    ' second, then every input column. No names at all: the case's games carry no level (it has
    ' no Level column - Jaq, 2026-10-08: say so, never guess), or CIS has not run.
    Set rngTable = LevelInputsTable(wb)
    If rngTable Is Nothing Then
        strStatus = FAILED_PREFIX & "no levels are tagged in this file - the case has no Level column, " & _
                    "or Create Case Inputs Sheet (CIS) has not run."
        GoTo Cleanup
    End If

    ' Every chosen level must exist (its L06_Inputs name is the proof CIS found it).
    For Each varLevel In colLevels
        strName = modCaseSetup.LevelInputsName(CLng(varLevel))
        If Not WorkbookNameExists(wb, strName) Then
            strStatus = FAILED_PREFIX & "there is no level " & varLevel & " on CaseInputs (no " & strName & ")."
            GoTo Cleanup
        End If
        strArgs = strArgs & "," & varLevel
        strShown = strShown & "," & varLevel
    Next varLevel
    strArgs = Mid$(strArgs, 2)

    ' THE LAYOUT (Jaq, 2026-10-08) - the inputs are an array to work on directly, nothing else in it:
    '   active cell (Z5)      Game | Level headers       two columns right (AB5)  the input headers
    '   below it    (Z6#)     game and level numbers     below that       (AB6#)  THE INPUTS
    ' Headers are for information; the cursor finishes on the inputs. All four anchors must be free.
    Set rngInfoHdr = rngTarget
    Set rngInfo = rngTarget.Offset(1, 0)
    Set rngInHdr = rngTarget.Offset(0, 2)
    Set rngData = rngTarget.Offset(1, 2)
    If Len(CStr(rngInfoHdr.Formula)) > 0 Or Len(CStr(rngInfo.Formula)) > 0 _
       Or Len(CStr(rngInHdr.Formula)) > 0 Or Len(CStr(rngData.Formula)) > 0 Then
        strStatus = FAILED_PREFIX & rngInfoHdr.Address(False, False) & ":" & rngData.Address(False, False) & _
                    " is not empty (game info here, inputs two columns right, headers above each). " & _
                    "Select an empty area; nothing was overwritten."
        GoTo Cleanup
    End If

    ' The formulas - written to be READ by the user who gets them (Jaq, 2026-10-08), so the LET
    ' names say what each step is:
    '   levels      the levels asked for, e.g. {3,5,6}
    '   caseInputs  the CaseInputs table: Game, Level, then every input column
    '   levelRows   FILTER of caseInputs by its Level column against levels (IsInList_byErikOehm,
    '               or XMATCH where the file lacks it) - one FILTER, so the
    '               rows come in the case's own order. (Not VSTACK of the level names: VSTACK of a
    '               blank-tested range gives #VALUE! on text over 255 characters, and dice-roll
    '               inputs run to 448 on 2023 SA Steppies - found 2026-10-08.)
    '   tidyRows    an EMPTY input shown blank: through a reference it reads 0, which looks like a
    '               real input (IWD case). LEN, not ="": Excel's = and <> give #VALUE! over 255
    '               characters.
    '   inputs      tidyRows without Game and Level
    '   usedCols    TRUE for an input column with anything in it for these levels; the others are
    '               left out of the inputs and their headers alike, so each header stays over its column.
    strTable = "'" & rngTable.Worksheet.Name & "'!" & rngTable.Address
    strHeaderRow = "'" & rngTable.Worksheet.Name & "'!" & _
                   rngTable.Worksheet.Cells(HEADER_ROW, rngTable.Column).Resize(1, rngTable.Columns.Count).Address
    ' The level test uses CompBot's IsInList lambda (Jaq, 2026-10-08), and the command makes sure it
    ' is there: Full Setup Case normally imports it, but not if the lambda import is switched off in
    ' SCS, so a missing copy is added from CompBot first (added only - an existing one is never
    ' replaced). Only if that copy fails is the same test written out with XMATCH, so the formula
    ' never shows #NAME?.
    If Not WorkbookNameExists(wb, IS_IN_LIST) Then
        If CopyLambdaFromCompBot(wb, IS_IN_LIST) Then strLoaded = "; " & IS_IN_LIST & " loaded from CompBot"
    End If
    If WorkbookNameExists(wb, IS_IN_LIST) Then
        strInLevels = IS_IN_LIST & "(INDEX(caseInputs,,2),levels)"
    Else
        strInLevels = "ISNUMBER(XMATCH(INDEX(caseInputs,,2),levels))"
    End If
    strStack = "LET(levels,{" & strArgs & "},caseInputs," & strTable & _
               ",levelRows,FILTER(caseInputs," & strInLevels & ")" & _
               ",tidyRows,IF(LEN(levelRows)=0,"""",levelRows),"
    strInputs = strStack & "inputs,DROP(tidyRows,,2),usedCols,BYCOL(inputs,LAMBDA(col,SUM(LEN(col))>0)),"
    rngInfoHdr.Formula2 = "=TAKE(" & strHeaderRow & ",,2)"
    rngInfo.Formula2 = "=" & strStack & "TAKE(tidyRows,,2))"
    rngInHdr.Formula2 = "=" & strInputs & "FILTER(DROP(" & strHeaderRow & ",,2),usedCols,""""))"
    rngData.Formula2 = "=" & strInputs & "FILTER(inputs,usedCols,""""))"

    ' A spill that runs into something below or to the right shows #SPILL! - the result never
    ' lands, so take all four back out and say so rather than leave a half-written block.
    strBlocked = SpillBlocked(Array(rngInfoHdr, rngInfo, rngInHdr, rngData))
    If Len(strBlocked) > 0 Then
        rngInfoHdr.ClearContents
        rngInfo.ClearContents
        rngInHdr.ClearContents
        rngData.ClearContents
        strStatus = FAILED_PREFIX & "the result needs more room - something is in the way of " & strBlocked & _
                    ". Pick a spot with empty space below and to the right; nothing was left behind."
        GoTo Cleanup
    End If

    ' Fit each column the result fills, but only a column holding nothing else (Jaq, 2026-10-08):
    ' short inputs in a wide column are hard to read; a column shared with the user's work is left alone.
    FitOwnColumns Array(rngInfoHdr, rngInfo, rngInHdr, rngData)

    ' Select is deliberate: finish on the inputs, ready to work on them (Jaq, 2026-10-08).
    rngData.Select
    strStatus = "Level Inputs: level " & Mid$(strShown, 2) & " - inputs in " & rngData.Address(False, False) & _
                "# (headers above), game and level in " & rngInfo.Address(False, False) & "#; blank columns left out" & _
                strLoaded & "."

Cleanup:
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strStatus = FAILED_PREFIX & Err.Number & ": " & Err.Description & " (see " & LogLocation() & ")"
    Resume Cleanup
End Sub


' Purpose: "3,5-6" -> 3, 5, 6 in the order typed (a range may run either way: 7-5 is 7, 6, 5).
'          Spaces are ignored; a repeated level is kept once. Returns Nothing, with the reason in
'          strWhy, for anything else - a letter, an empty part, a level below 1 or above 99.
Private Function ParseLevelList(ByVal strText As String, ByRef strWhy As String) As Collection

    ' --- CONSTANTS (local to function) ---
    Const MAX_LEVEL As Long = 99

    Dim colOut  As Collection
    Dim dicSeen As Object
    Dim varPart As Variant
    Dim strPart As String
    Dim lngDash As Long
    Dim lngFrom As Long
    Dim lngTo   As Long
    Dim lngStep As Long
    Dim lngL    As Long

    On Error GoTo ErrHandler

    Set colOut = New Collection
    Set dicSeen = CreateObject("Scripting.Dictionary")

    For Each varPart In Split(Replace(strText, " ", vbNullString), ",")
        strPart = CStr(varPart)
        lngDash = InStr(1, strPart, "-")
        If lngDash > 0 Then
            If Not IsWholeNumber(Left$(strPart, lngDash - 1)) Or Not IsWholeNumber(Mid$(strPart, lngDash + 1)) Then
                strWhy = "use level numbers, e.g. 4 or 4-5 or 3,5-6."
                GoTo Cleanup
            End If
            lngFrom = CLng(Left$(strPart, lngDash - 1))
            lngTo = CLng(Mid$(strPart, lngDash + 1))
        ElseIf IsWholeNumber(strPart) Then
            lngFrom = CLng(strPart)
            lngTo = lngFrom
        Else
            strWhy = "use level numbers, e.g. 4 or 4-5 or 3,5-6."
            GoTo Cleanup
        End If
        If lngFrom < 1 Or lngTo < 1 Or lngFrom > MAX_LEVEL Or lngTo > MAX_LEVEL Then
            strWhy = "levels run from 1 to " & MAX_LEVEL & "."
            GoTo Cleanup
        End If
        lngStep = IIf(lngTo >= lngFrom, 1, -1)
        For lngL = lngFrom To lngTo Step lngStep
            If Not dicSeen.Exists(lngL) Then
                dicSeen(lngL) = True
                colOut.Add lngL
            End If
        Next lngL
    Next varPart

    If colOut.Count > 0 Then Set ParseLevelList = colOut Else strWhy = "no level given."

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strWhy = "could not read it (" & Err.Description & ")."
    Set ParseLevelList = Nothing
    Resume Cleanup
End Function


' Purpose: Add CompBot's own lambda strName to wb, with its comment, when wb has none of that name.
'          Never replaces anything. Returns True if it was added. Never raises.
Private Function CopyLambdaFromCompBot(ByVal wb As Workbook, ByVal strName As String) As Boolean

    Dim nmSource As Name
    Dim nmNew    As Name

    On Error GoTo ErrHandler

    If WorkbookNameExists(wb, strName) Then GoTo Cleanup
    Set nmSource = ThisWorkbook.Names(strName)
    Set nmNew = wb.Names.Add(Name:=strName, RefersTo:=nmSource.RefersTo)
    On Error Resume Next                          ' narrow: a comment is nice to have, never fatal
    nmNew.Comment = nmSource.Comment
    On Error GoTo ErrHandler
    CopyLambdaFromCompBot = True

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "CopyLambdaFromCompBot", Err.Number, Err.Description
    CopyLambdaFromCompBot = False
    Resume Cleanup
End Function


' Purpose: The addresses (comma-separated) of the cells in varCells whose formula shows #SPILL!,
'          or "". Never raises: an unreadable cell counts as blocked.
Private Function SpillBlocked(ByVal varCells As Variant) As String

    ' --- CONSTANTS (local to function) ---
    Const ERR_SPILL As Long = 2045          ' xlErrSpill

    Dim varCell As Variant
    Dim varVal  As Variant
    Dim strOut  As String

    On Error GoTo ErrHandler

    For Each varCell In varCells
        varVal = varCell.value
        If IsError(varVal) Then
            If varVal = CVErr(ERR_SPILL) Then strOut = strOut & ", " & varCell.Address(False, False)
        End If
    Next varCell
    SpillBlocked = Mid$(strOut, 3)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    SpillBlocked = "the result"
    Resume Cleanup
End Function


' Purpose: AutoFit every column the results in varCells fill (each cell's spill, or the cell alone),
'          but only a column where those results are ALL there is - a column that also holds the
'          user's own work keeps its width. Never raises: a column that cannot be checked is left alone.
Private Sub FitOwnColumns(ByVal varCells As Variant)

    Dim varCell  As Variant
    Dim rngOurs  As Range
    Dim rngArea  As Range
    Dim rngCol   As Range
    Dim ws       As Worksheet
    Dim lngCol   As Long
    Dim lngFirst As Long
    Dim lngLast  As Long

    On Error GoTo ErrHandler

    For Each varCell In varCells
        Set rngArea = Nothing
        On Error Resume Next                      ' narrow: a single-value result has no spill
        If varCell.HasSpill Then Set rngArea = varCell.SpillingToRange
        On Error GoTo ErrHandler
        If rngArea Is Nothing Then Set rngArea = varCell
        If rngOurs Is Nothing Then Set rngOurs = rngArea Else Set rngOurs = Union(rngOurs, rngArea)
    Next varCell
    If rngOurs Is Nothing Then GoTo Cleanup

    Set ws = rngOurs.Worksheet
    lngFirst = ws.Columns.Count
    For Each rngArea In rngOurs.Areas
        If rngArea.Column < lngFirst Then lngFirst = rngArea.Column
        If rngArea.Column + rngArea.Columns.Count - 1 > lngLast Then lngLast = rngArea.Column + rngArea.Columns.Count - 1
    Next rngArea

    For lngCol = lngFirst To lngLast
        Set rngCol = ws.Columns(lngCol)
        If Not Intersect(rngCol, rngOurs) Is Nothing Then
            If Application.WorksheetFunction.CountA(rngCol) = _
               Application.WorksheetFunction.CountA(Intersect(rngCol, rngOurs)) Then rngCol.AutoFit
        End If
    Next lngCol

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Resume Cleanup
End Sub


' Purpose: The CaseInputs table the level names cover - from the first level's first row to the
'          last level's last row, and from Game across to the widest level's last column - or
'          Nothing if the file has no L##_Inputs names. Never raises.
Private Function LevelInputsTable(ByVal wb As Workbook) As Range

    Dim nmItem   As Name
    Dim rngName  As Range
    Dim wsTable  As Worksheet
    Dim lngTop   As Long
    Dim lngLeft  As Long
    Dim lngBottom As Long
    Dim lngRight As Long

    On Error GoTo ErrHandler

    For Each nmItem In wb.Names
        If nmItem.Name Like "L##_Inputs" And TypeName(nmItem.Parent) = "Workbook" Then
            Set rngName = Nothing
            On Error Resume Next                  ' narrow: a name whose range was deleted
            Set rngName = nmItem.RefersToRange
            On Error GoTo ErrHandler
            If Not rngName Is Nothing Then
                If wsTable Is Nothing Then
                    Set wsTable = rngName.Worksheet
                    lngTop = rngName.Row
                    lngLeft = rngName.Column
                End If
                If rngName.Row < lngTop Then lngTop = rngName.Row
                If rngName.Column < lngLeft Then lngLeft = rngName.Column
                If rngName.Row + rngName.Rows.Count - 1 > lngBottom Then lngBottom = rngName.Row + rngName.Rows.Count - 1
                If rngName.Column + rngName.Columns.Count - 1 > lngRight Then lngRight = rngName.Column + rngName.Columns.Count - 1
            End If
        End If
    Next nmItem

    If Not wsTable Is Nothing Then
        Set LevelInputsTable = wsTable.Range(wsTable.Cells(lngTop, lngLeft), wsTable.Cells(lngBottom, lngRight))
    End If

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Set LevelInputsTable = Nothing
    Resume Cleanup
End Function


' Purpose: True for 1 to 3 digits and nothing else ("06" counts; "4.5", "", "x" do not).
Private Function IsWholeNumber(ByVal strText As String) As Boolean
    IsWholeNumber = (Len(strText) >= 1 And Len(strText) <= 3 And Not strText Like "*[!0-9]*")
End Function


' Purpose: True if wb has a WORKBOOK-level name strName. Narrow probe; never raises.
Private Function WorkbookNameExists(ByVal wb As Workbook, ByVal strName As String) As Boolean

    Dim nmProbe As Name

    On Error Resume Next                          ' narrow: probing for a name that may not exist
    Set nmProbe = wb.Names(strName)
    On Error GoTo 0
    If Not nmProbe Is Nothing Then WorkbookNameExists = (TypeName(nmProbe.Parent) = "Workbook")
End Function










