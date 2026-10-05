Attribute VB_Name = "modDataTable"
Option Explicit

' High-resolution counter, for timing single recalculations (Run Data Table's recalc-scope check).
#If VBA7 Then
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#Else
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (ByRef lpPerformanceCount As Currency) As Long
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (ByRef lpFrequency As Currency) As Long
#End If

' --- MODULE CONSTANTS ---
Private Const m_DEBUG_MODE          As Boolean = False
Private Const m_MAX_STATUS          As Long = 250   ' Application.StatusBar rejects long strings
Private Const m_BODY_OFFSET_ROWS    As Long = 4     ' first game row, below the input cell
Private Const m_FORMULA_OFFSET_ROWS As Long = 3     ' the "ANSWER" formula cell, below the input cell
Private Const m_SCOPE_FULL          As Long = 0     ' Run Data Table recalc scopes
Private Const m_SCOPE_SHEET         As Long = 1
Private Const m_LABEL_PREFIX        As String = "RDT_Label_"   ' hidden names remembering Example# labels
Private Const m_SCAN_CELL_LIMIT     As Double = 1000000#   ' recalc check: most cells read from one block
Private Const m_KEY_INDIRECT        As String = "*INDIRECT" ' recalc check: marker key (* is never in a sheet name)
' Functions that force wider recalculation. Matched against formula text, so each ends in "(".
Private Const m_VOLATILE_LIST     As String = "INDIRECT(|OFFSET(|TODAY(|NOW(|RAND(|RANDBETWEEN(|RANDARRAY(|CELL(|INFO("
' Side columns (Link Side Cell): a second per-game output beside the answer column, for bonuses.
Private Const m_SIDE_PREFIX       As String = "SIDE "          ' label in the Check row: "SIDE 1", "SIDE 2", ...
Private Const m_MAX_SIDES         As Long = 5
Private Const m_NOTE_PREFIX       As String = "No room for SIDE "   ' Link Side Cell's refusal note
Private Const m_SIDE_FILL         As Long = 15652797            ' RGB(189, 215, 238): light blue, never ANSWER's green
Private Const m_NOTE_COLOUR       As Long = 1667760             ' RGB(176, 114, 25): burnt orange B07219

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Run Data Table
' Description:            Fills a Create Data Table block with static results, recalculating only when run
' Macro Expression:       modDataTable.RunDataTable()
' Generated:              2026-09-14
'----------------------------------------------------------------------------------------------------
' Purpose: Replacement for Excel's What-If data table, which recalculates unpredictably. Works on the
'          block Create Data Table builds:
'              input cell      "Example#" - the XLOOKUP input row beside it reads the game from it
'              formula cell    3 rows below, 1 column right - point it at the calc output
'              game numbers    from 4 rows below the input cell, down
'              results         the column to the right of the game numbers (answer cells link here)
'          For each game it writes the game number into the input cell, recalculates, and reads the
'          formula cell. Results are written as STATIC values in one hit, any What-If TABLE() formulas
'          are replaced, and the input cell is restored so the example check value stays visible.
'          A game that could not be calculated is written as #N/A, never left blank (a blank would
'          show as 0 in the linked answer cell and look like a real answer).
'          SIDE columns (Link Side Cell, 2026-09-24): the block's output row is ANSWER plus any SIDE
'          heads to its right. Each game is recalculated ONCE and the whole output row read together,
'          so every column comes from the same pass (Jaq: "calculate for each row and then both values
'          can be loaded in, row by row"). Every run refills every column whose head is a live formula;
'          a column whose head is not a formula is left untouched and named on the StatusBar, so a
'          constant is never written down every game.
Public Sub RunDataTable()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "RunDataTable"

    Dim wsCalc As Worksheet
    Dim rngInput As Range
    Dim rngFormula As Range          ' the output row: ANSWER, then any SIDE heads
    Dim rngGames As Range
    Dim rngResults As Range          ' every result column, answer column first
    Dim varGames As Variant
    Dim varOut() As Variant
    Dim varRow As Variant
    Dim varOriginal As Variant
    Dim blnLive() As Boolean         ' per output column: its head is a formula, so it is filled
    Dim blnClearedCol() As Boolean   ' per output column: What-If TABLE() formulas were cleared
    Dim lngErrors() As Long
    Dim lngFirstError() As Long
    Dim lngSides As Long
    Dim lngCols As Long
    Dim lngLive As Long
    Dim lngJ As Long
    Dim lngGames As Long
    Dim lngI As Long
    Dim lngInterruptKey As Long
    Dim blnDone As Boolean
    Dim blnWriting As Boolean        ' the per-column write has started (Cleanup finishes it)
    Dim strSkipped As String
    Dim strFilled As String
    Dim dblStart As Double
    Dim dblSecs As Double
    Dim blnInit As Boolean
    Dim blnEvents As Boolean
    Dim blnKeysSet As Boolean
    Dim blnRestore As Boolean
    Dim blnCleared As Boolean
    Dim blnWritten As Boolean
    Dim blnSheetOnly As Boolean
    Dim blnTouched As Boolean
    Dim dblCheckStart As Double
    Dim dblCheckSecs As Double
    Dim strScope As String
    Dim strSlow As String
    Dim strMsg As String
    Dim strErr As String

    On Error GoTo ErrHandler

    If TypeName(ActiveSheet) <> "Worksheet" Then
        strMsg = "Run Data Table: select a cell on a worksheet first."
        GoTo Cleanup
    End If
    Set wsCalc = ActiveSheet

    Set rngInput = FindInputCell(wsCalc, ActiveCell, "Run Data Table", strMsg)
    If rngInput Is Nothing Then GoTo Cleanup

    ' The output row: ANSWER plus the SIDE heads. Only columns whose head is a formula are filled.
    lngSides = SideColumnCount(rngInput)
    lngCols = 1 + lngSides
    Set rngFormula = rngInput.Offset(m_FORMULA_OFFSET_ROWS, 1).Resize(1, lngCols)
    ReDim blnLive(1 To lngCols)
    ReDim blnClearedCol(1 To lngCols)
    ReDim lngErrors(1 To lngCols)
    ReDim lngFirstError(1 To lngCols)
    For lngJ = 1 To lngCols
        blnLive(lngJ) = rngFormula.Cells(1, lngJ).HasFormula
        If blnLive(lngJ) Then
            lngLive = lngLive + 1
            strFilled = strFilled & IIf(Len(strFilled) > 0, ", ", "") & OutputTitle(lngJ)
        Else
            strSkipped = strSkipped & IIf(Len(strSkipped) > 0, ", ", "") & OutputTitle(lngJ)
        End If
    Next lngJ
    If lngLive = 0 Then
        strMsg = "Run Data Table: point " & rngFormula.Cells(1, 1).Address(False, False) & " (ANSWER) at your example's output first."
        GoTo Cleanup
    End If

    lngGames = CountGames(rngInput.Offset(m_BODY_OFFSET_ROWS, 0))
    If lngGames = 0 Then
        strMsg = "Run Data Table: no game numbers below " & rngInput.Address(False, False) & "."
        GoTo Cleanup
    End If

    Set rngGames = rngInput.Offset(m_BODY_OFFSET_ROWS, 0).Resize(lngGames, 1)
    Set rngResults = rngGames.Offset(0, 1).Resize(lngGames, lngCols)

    ' Refuse up front rather than after a full run.
    If wsCalc.ProtectContents Then
        strMsg = "Run Data Table: " & wsCalc.Name & " is protected - unprotect it first."
        GoTo Cleanup
    End If
    For lngJ = 1 To lngCols
        If blnLive(lngJ) Then
            If IsNull(rngResults.Columns(lngJ).MergeCells) Then
                strMsg = "Run Data Table: the " & OutputTitle(lngJ) & " results column has merged cells."
                GoTo Cleanup
            ElseIf rngResults.Columns(lngJ).MergeCells Then
                strMsg = "Run Data Table: the " & OutputTitle(lngJ) & " results column has merged cells."
                GoTo Cleanup
            End If
            ' An old What-If table cannot be cleared one column at a time; with SIDE columns, refuse.
            If lngSides > 0 Then
                If HasTableFormulas(rngResults.Columns(lngJ)) Then
                    strMsg = "Run Data Table: an old What-If data table is in the " & OutputTitle(lngJ) & _
                             " column - clear it before running with SIDE columns."
                    GoTo Cleanup
                End If
            End If
            ' A What-If table longer than the counted games cannot be partly cleared.
            If HasTableFormulas(rngResults.Columns(lngJ).Offset(lngGames, 0).Resize(1, 1)) Then
                strMsg = "Run Data Table: game numbers stop at row " & rngGames.Row + lngGames - 1 & _
                         " but the data table carries on - check " & _
                         rngResults.Columns(lngJ).Offset(lngGames, 0).Resize(1, 1).Address(False, False) & "."
                GoTo Cleanup
            End If
        End If
    Next lngJ

    ' One read for all game numbers. A single cell comes back as a scalar, not an array.
    If lngGames = 1 Then
        ReDim varGames(1 To 1, 1 To 1)
        varGames(1, 1) = rngGames.Value2
    Else
        varGames = rngGames.Value2
    End If

    ' Every result starts as #N/A, so a game that never completes is visibly wrong.
    ReDim varOut(1 To lngGames, 1 To lngCols)
    For lngI = 1 To lngGames
        For lngJ = 1 To lngCols
            varOut(lngI, lngJ) = CVErr(xlErrNA)
        Next lngJ
    Next lngI
    lngI = 0

    VBAInit
    blnInit = True
    blnEvents = Application.EnableEvents
    Application.EnableEvents = False
    ' Esc / Ctrl+Break raise error 18 into the handler (no dialog, cleanup still runs), and a key
    ' press cannot silently interrupt a recalculation mid-game.
    lngInterruptKey = Application.CalculationInterruptKey
    Application.CalculationInterruptKey = xlNoKey
    Application.EnableCancelKey = xlErrorHandler
    blnKeysSet = True

    ' What-If TABLE() formulas recalculate on every pass and slow the loop badly - remove them
    ' first. The results range is the whole table body, so Excel allows the clear.
    For lngJ = 1 To lngCols
        If blnLive(lngJ) Then
            If HasTableFormulas(rngResults.Columns(lngJ)) Then
                rngResults.Columns(lngJ).ClearContents
                blnClearedCol(lngJ) = True
                blnCleared = True
            End If
        End If
    Next lngJ

    ' Restore to the example label, even if a game was loaded with Load Game Into Calculation.
    ' A formula in the input cell is kept as a formula.
    If rngInput.HasFormula Then
        varOriginal = rngInput.Formula
    Else
        varOriginal = ExampleLabel(rngInput)
        If Len(varOriginal) = 0 Then varOriginal = rngInput.Formula
    End If
    blnRestore = True

    ' Recalc scope: the sheet only is ~5x faster than Application.Calculate (benchmarked on Bushland
    ' Level 6, 2026-09-14) but is only correct when the whole solve chain sits on this sheet.
    ' ChooseRecalcScope proves it (structural scan + sentinel probes) or picks the full recalc.
    ' Timed separately, so the games figure is comparable between the two scopes.
    blnTouched = True     ' from here the input cell and other sheets hold probe states
    dblCheckStart = Timer
    blnSheetOnly = (ChooseRecalcScope(wsCalc, rngInput, rngFormula, varGames, strScope) = m_SCOPE_SHEET)
    dblCheckSecs = Timer - dblCheckStart
    If dblCheckSecs < 0 Then dblCheckSecs = dblCheckSecs + 86400#

    dblStart = Timer
    For lngI = 1 To lngGames
        rngInput.Value2 = varGames(lngI, 1)
        ' Calculation is manual: recalculate this sheet only when verified safe, else every open
        ' workbook (correct across sheets).
        ' CalculationState is only meaningful after a full calculation - after a sheet-only one it can
        ' read "pending" because of other workbooks. Interrupt keys are off, so both run to the end.
        If blnSheetOnly Then
            wsCalc.Calculate
            blnDone = True
        Else
            Application.Calculate
            blnDone = (Application.CalculationState = xlDone)
        End If
        If blnDone Then
            ' One read of the whole output row: every column from the same recalculation.
            varRow = rngFormula.Value2
            For lngJ = 1 To lngCols
                If blnLive(lngJ) Then
                    If lngCols = 1 Then
                        varOut(lngI, 1) = varRow
                    Else
                        varOut(lngI, lngJ) = varRow(1, lngJ)
                    End If
                End If
            Next lngJ
        End If
        For lngJ = 1 To lngCols
            If blnLive(lngJ) Then
                If IsError(varOut(lngI, lngJ)) Then
                    lngErrors(lngJ) = lngErrors(lngJ) + 1
                    If lngFirstError(lngJ) = 0 Then lngFirstError(lngJ) = CLng(varGames(lngI, 1))
                End If
            End If
        Next lngJ
    Next lngI

    rngInput.Formula = varOriginal
    blnRestore = False
    If VarType(rngInput.Value2) = vbString Then ForgetLabel rngInput

    ' Write the results BEFORE the final recalculation, so the linked answer cells pick them up
    ' even when the user's calculation mode is manual. All in one hit when every column is live;
    ' otherwise column by column, so a column that is not filled keeps what it holds.
    If lngLive = lngCols Then
        rngResults.Value2 = varOut
    Else
        blnWriting = True
        For lngJ = 1 To lngCols
            If blnLive(lngJ) Then rngResults.Columns(lngJ).Value2 = OutputColumn(varOut, lngJ)
        Next lngJ
    End If
    blnWritten = True
    Application.Calculate

    dblSecs = Timer - dblStart
    If dblSecs < 0 Then dblSecs = dblSecs + 86400#     ' Timer wraps at midnight

    strMsg = "Run Data Table (" & wsCalc.Name & "): " & lngGames & " games in " & Format$(dblSecs, "0.00") & "s"
    If lngSides > 0 Then strMsg = strMsg & " (" & strFilled & ")"
    strMsg = strMsg & " - last run " & Format$(Now, "hh:nn:ss") & "." & ExampleCheckText(rngInput)
    If lngSides = 0 Then
        If lngErrors(1) > 0 Then
            strMsg = strMsg & " " & lngErrors(1) & " game(s) returned errors (first: game " & lngFirstError(1) & ")."
        End If
    Else
        For lngJ = 1 To lngCols
            If blnLive(lngJ) And lngErrors(lngJ) > 0 Then
                strMsg = strMsg & " " & OutputTitle(lngJ) & ": " & lngErrors(lngJ) & " error(s) (first: game " & _
                         lngFirstError(lngJ) & ")."
            End If
        Next lngJ
    End If
    If Len(strSkipped) > 0 Then
        strMsg = strMsg & " Not filled, not a formula: " & strSkipped & "."
    End If
    strSlow = SlowdownsIn(wsCalc.Parent)
    If Len(strSlow) > 0 Then
        strMsg = strMsg & " WARNING slow: " & strSlow
        If Right$(strMsg, 1) <> "." Then strMsg = strMsg & "."
    End If
    ' Scope last, so it is what gets cut if the StatusBar text is too long - never the warnings.
    strMsg = strMsg & " Recalc: " & strScope & " (check " & Format$(dblCheckSecs, "0.00") & "s)."

Cleanup:
    On Error Resume Next
    If blnRestore Then rngInput.Formula = varOriginal
    ' Put back what a failure left half-done: cleared What-If columns, and, if Esc or an error hit
    ' part-way through the per-column write, every live column - so no column is left from an older run
    ' beside newer ones.
    If (blnCleared Or blnWriting) And Not blnWritten Then
        For lngJ = 1 To lngCols
            If blnClearedCol(lngJ) Or (blnWriting And blnLive(lngJ)) Then
                rngResults.Columns(lngJ).Value2 = OutputColumn(varOut, lngJ)
            End If
        Next lngJ
    End If
    ' Esc or a failure after the check started: other sheets may still be calculated on the probe
    ' sentinel. Recalculate before the keys are handed back, so it cannot be interrupted.
    If blnTouched And Not blnWritten Then Application.Calculate
    If blnKeysSet Then
        Application.CalculationInterruptKey = lngInterruptKey
        Application.EnableCancelKey = xlInterrupt
    End If
    If blnInit Then
        Application.EnableEvents = blnEvents
        VBAFin
    End If
    If Len(strErr) > 0 Then
        Application.StatusBar = Left$("Run Data Table failed - " & strErr, m_MAX_STATUS)
    ElseIf Len(strMsg) > 0 Then
        Application.StatusBar = Left$(strMsg, m_MAX_STATUS)
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If lngI > 0 And lngI <= lngGames Then strErr = strErr & " (game " & lngI & " of " & lngGames & ")"
    If LogError(PROC_NAME, Err.Number, strErr) Then strErr = strErr & " - see " & LogLocation()
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Load Game Into Calculation
' Description:            Loads the selected data table game into the example calculation, or restores the example
' Macro Expression:       modDataTable.LoadGameIntoCalculation()
' Generated:              2026-09-14
'----------------------------------------------------------------------------------------------------
' Purpose: For edge-case testing and troubleshooting. With a game row of the data table selected
'          (its game number or its result), put that game number into the input cell and recalculate,
'          so the whole calc block shows that game. With the input cell or ANSWER row selected, put
'          the "Example#" label back. The label is remembered in a hidden sheet-level name, so Run Data
'          Table also restores it after a game was loaded.
Public Sub LoadGameIntoCalculation()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "LoadGameIntoCalculation"

    Dim ws As Worksheet
    Dim rngInput As Range
    Dim rngFormula As Range
    Dim rngGames As Range
    Dim varGame As Variant
    Dim lngGames As Long
    Dim lngRow As Long
    Dim strLabel As String
    Dim strMsg As String
    Dim strErr As String

    On Error GoTo ErrHandler

    If TypeName(ActiveSheet) <> "Worksheet" Then
        strMsg = "Load Game: select a cell in a data table first."
        GoTo Cleanup
    End If
    Set ws = ActiveSheet

    Set rngInput = FindInputCell(ws, ActiveCell, "Load Game", strMsg)
    If rngInput Is Nothing Then GoTo Cleanup

    lngGames = CountGames(rngInput.Offset(m_BODY_OFFSET_ROWS, 0))
    ' The block is the game column, the answer column and any SIDE columns.
    If Intersect(ActiveCell, rngInput.Resize(m_BODY_OFFSET_ROWS + lngGames, 2 + SideColumnCount(rngInput))) Is Nothing Then
        strMsg = "Load Game: select a game row inside the data table (game number or result)."
        GoTo Cleanup
    End If
    If ws.ProtectContents Then
        strMsg = "Load Game: " & ws.Name & " is protected - unprotect it first."
        GoTo Cleanup
    End If

    Set rngFormula = rngInput.Offset(m_FORMULA_OFFSET_ROWS, 1)
    Set rngGames = rngInput.Offset(m_BODY_OFFSET_ROWS, 0).Resize(lngGames, 1)
    strLabel = ExampleLabel(rngInput)
    lngRow = ActiveCell.Row

    If lngRow >= rngGames.Row Then
        varGame = ws.Cells(lngRow, rngGames.Column).Value2
        If Not RememberLabel(rngInput) Then
            strMsg = "Load Game: could not remember the example label - nothing changed."
            GoTo Cleanup
        End If
        rngInput.Value2 = varGame
        Application.Calculate
        strMsg = "Load Game (" & ws.Name & "): game " & varGame & " is in the calculation - output " & _
                 SafeText(rngFormula.Value2) & ". Select the input cell and run again to restore " & strLabel & "."
    Else
        If Len(strLabel) = 0 Then
            strMsg = "Load Game: the example label is not known - type it back into " & rngInput.Address(False, False) & "."
            GoTo Cleanup
        End If
        rngInput.Value2 = strLabel
        ForgetLabel rngInput
        Application.Calculate
        strMsg = "Load Game (" & ws.Name & "): " & strLabel & " restored - output " & SafeText(rngFormula.Value2) & "." & _
                 ExampleCheckText(rngInput)
    End If

Cleanup:
    On Error Resume Next
    If Len(strErr) > 0 Then
        Application.StatusBar = Left$("Load Game failed - " & strErr, m_MAX_STATUS)
    ElseIf Len(strMsg) > 0 Then
        Application.StatusBar = Left$(strMsg, m_MAX_STATUS)
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, strErr) Then strErr = strErr & " - see " & LogLocation()
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Data Table
' Description:            Builds (or finds) the data table block for the level you are in, beside its inputs
' Macro Expression:       modDataTable.CreateDataTable()
' Generated:              2026-09-24 (Solve Level and the old Create Data Table merged)
'----------------------------------------------------------------------------------------------------
' Purpose: Build the level's data table block, or find the one it already has, and land on the block's
'          LAST INPUT CELL, ready to build the example calculation beside it. Works on the sheet the
'          user is on - never moves them to another sheet and never creates level sheets (Jaq,
'          2026-09-14: some people solve on the Case sheet, some on the L## sheets).
'          Which level (first match wins):
'            - the active cell is the level's "Example#" label: that level, on any sheet, in any column
'              (the old Create Data Table rule);
'            - an L## sheet: the level on that sheet (any block on it is that level's);
'            - any other sheet: the level whose "Level #" heading in column B is the last one at or
'              above the active row (Contents lists ignored; the bonus section is refused).
'          Existing block: nothing is built; the block's last input is selected.
'          Level answers linking to another sheet that already has a block: refused, rather than
'          re-point the answers away from the sheet the level is being solved on.
'          New block: level with the example row, as close as possible to two columns right of the
'          level's last input, moved right only as far as it takes to overlap nothing (BlockTarget).
'          The view scrolls to the block if it is off screen.
'          The level finding is LocateLevel, shared with Go To Example.
'          Merged 2026-09-24 (Jaq): Solve Level's finding and placement, the old command's name, launch
'          code and end state. Recorded A-ZTraining replays (sheet B, Levels 6-7) start from the last
'          input cell, which is why the end state is that cell and not ANSWER.
Public Sub CreateDataTable()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "CreateDataTable"

    Dim wbCase As Workbook
    Dim wsWork As Worksheet
    Dim wsLinked As Worksheet
    Dim colGames As Collection
    Dim colIn As Collection
    Dim rngInput As Range
    Dim rngZone As Range
    Dim rngTarget As Range
    Dim rngLand As Range
    Dim varGameRow As Variant
    Dim blnLevelSheet As Boolean
    Dim blnOnLabel As Boolean
    Dim blnFailed As Boolean
    Dim lngTop As Long
    Dim lngBottom As Long
    Dim lngExRow As Long
    Dim lngGameCol As Long
    Dim lngHdRow As Long
    Dim lngACol As Long
    Dim lngAnsCol As Long
    Dim lngFirstCol As Long
    Dim lngWidth As Long
    Dim strLinked As String
    Dim strWhere As String
    Dim strFound As String
    Dim strMsg As String
    Dim strErr As String

    On Error GoTo ErrHandler

    If TypeName(ActiveSheet) <> "Worksheet" Then
        strMsg = "Create Data Table: select a cell in a level first."
        GoTo Cleanup
    End If
    Set wsWork = ActiveSheet
    Set wbCase = wsWork.Parent
    blnLevelSheet = IsLevelSheetName(wsWork.Name)
    blnOnLabel = IsExampleLabel(ActiveCell.Value2)

    ' 1. Which level, and which rows belong to it? (the rule Go To Example shares)
    If Not LocateLevel(wsWork, "Create Data Table", lngGameCol, lngExRow, lngTop, lngBottom, strWhere, strMsg) Then
        GoTo Cleanup
    End If

    ' The level's games - the same rule BuildDataTable uses.
    Set colGames = LevelGameRows(wsWork, lngGameCol, lngExRow, lngBottom)
    ' From the label, the level ends with its games: a later level's block must not be taken for this one.
    If blnOnLabel And colGames.Count > 0 Then lngBottom = colGames(colGames.Count)

    ' 2. This level already has a block: use it. On a level sheet any block is this level's; elsewhere
    '    only a block whose input cell lies in this level's rows.
    If blnLevelSheet Then
        Set rngZone = wsWork.Cells
    Else
        Set rngZone = wsWork.Range(wsWork.Cells(lngTop, 1), wsWork.Cells(lngBottom, wsWork.Columns.Count))
    End If
    Set rngInput = FindInputCell(wsWork, ActiveCell, "Create Data Table", strFound, rngZone, blnFailed)
    If blnFailed Then
        ' The search itself failed - do not build blind next to a block that may already exist.
        strMsg = strFound
        GoTo Cleanup
    End If
    If Not rngInput Is Nothing Then
        Set rngLand = LastInputCell(rngInput)
        ' Putting the user on the block is the purpose of the command, so Select is deliberate.
        rngLand.Select
        ShowBlock rngInput, rngLand
        strMsg = "Create Data Table (" & strWhere & "): block already at " & rngInput.Address(False, False) & _
                 ", ANSWER is " & rngInput.Offset(m_FORMULA_OFFSET_ROWS, 1).Address(False, False)
        If SideColumnCount(rngInput) > 0 Then strMsg = strMsg & ", " & SideColumnCount(rngInput) & " SIDE column(s)"
        strMsg = strMsg & "; run RDT after changing the solve." & ExampleCheckText(rngInput)
        GoTo Cleanup
    End If

    ' 3. If this level's answers link to another sheet that already has a block, the level is being
    '    solved there - building here would re-point its answers to an empty table.
    lngHdRow = HeaderRowFor(wsWork, lngExRow)
    If lngHdRow < lngTop Then lngHdRow = lngTop
    If Not blnLevelSheet And colGames.Count > 0 And lngHdRow >= 1 Then
        lngACol = AnswerColumn(wsWork, lngHdRow)
        If lngACol > 0 Then
            For Each varGameRow In colGames
                strLinked = LinkedSheetName(CStr(wsWork.Cells(varGameRow, lngACol).Formula))
                If Len(strLinked) > 0 Then Exit For
            Next varGameRow
            If Len(strLinked) > 0 Then
                Set wsLinked = Nothing
                ' Probe: the linked sheet may have been deleted.
                On Error Resume Next
                Set wsLinked = wbCase.Worksheets(strLinked)
                On Error GoTo ErrHandler
                If Not wsLinked Is Nothing Then
                    Set rngInput = FindInputCell(wsLinked, wsLinked.Cells(1, 1), "Create Data Table", strFound, wsLinked.Cells, blnFailed)
                    If Not rngInput Is Nothing Then
                        strMsg = "Create Data Table: " & strWhere & " is being solved on " & wsLinked.Name & _
                                 " (data table at " & rngInput.Address(False, False) & ") - solve it there, " & _
                                 "or clear that block to solve on " & wsWork.Name & "."
                        GoTo Cleanup
                    End If
                End If
            End If
        End If
    End If

    ' 4. Place the block beside the level's inputs (the same input rule BuildDataTable copies from).
    Set colIn = LevelInputColumns(wsWork, lngGameCol, lngExRow, lngAnsCol)
    If colIn Is Nothing Then
        strMsg = "Create Data Table: could not read the inputs on the example row - see " & LogLocation()
        GoTo Cleanup
    End If
    If colIn.Count > 0 Then
        lngFirstCol = colIn(colIn.Count) + 2
    Else
        lngFirstCol = lngGameCol + 2
    End If
    lngWidth = colIn.Count + 1
    If lngWidth < 2 Then lngWidth = 2
    Set rngTarget = BlockTarget(wsWork, lngExRow, lngFirstCol, lngWidth, colGames.Count)
    If rngTarget Is Nothing Then
        strMsg = "Create Data Table: found no free space for the block to the right of the inputs on " & wsWork.Name & "."
        GoTo Cleanup
    End If

    ' 5. Build it. BuildDataTable reports its own reason when it fails.
    Set rngInput = BuildDataTable(rngTarget, wsWork.Cells(lngExRow, lngGameCol))
    If rngInput Is Nothing Then GoTo Cleanup
    Set rngLand = LastInputCell(rngInput)
    rngLand.Select
    ShowBlock rngInput, rngLand
    strMsg = "Create Data Table (" & strWhere & "): block built at " & rngInput.Address(False, False) & _
             "; build the example calc beside " & rngLand.Address(False, False) & _
             ", point ANSWER (" & rngInput.Offset(m_FORMULA_OFFSET_ROWS, 1).Address(False, False) & _
             ") at your output, then run RDT."

Cleanup:
    On Error Resume Next
    If Len(strErr) > 0 Then
        Application.StatusBar = Left$("Create Data Table failed - " & strErr, m_MAX_STATUS)
    ElseIf Len(strMsg) > 0 Then
        Application.StatusBar = Left$(strMsg, m_MAX_STATUS)
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, strErr) Then strErr = strErr & " - see " & LogLocation()
    Resume Cleanup
End Sub

' Purpose: Which level the active cell is in, the shared rule of Create Data Table and Go To Example
'          (first match wins):
'            - the active cell is the level's "Example#" label: that level, on any sheet, in any column;
'            - an L## sheet: the level on that sheet;
'            - any other sheet: the level whose "Level #" heading in column B is the last one at or
'              above the active row (Contents lists ignored; the bonus section is refused).
'          Returns True with the game column, example row, the level's row span and a "sheet, level"
'          description; False with strMsg saying why (strCmd heads the message).
Private Function LocateLevel(ByVal wsWork As Worksheet, ByVal strCmd As String, ByRef lngGameCol As Long, _
                             ByRef lngExRow As Long, ByRef lngTop As Long, ByRef lngBottom As Long, _
                             ByRef strWhere As String, ByRef strMsg As String) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const MAX_SCAN_ROWS As Long = 1000

    Dim colRows As Collection
    Dim lngLevel As Long
    Dim lngI As Long
    Dim lngLastRow As Long
    Dim lngBonusRow As Long

    On Error GoTo ErrHandler

    If IsExampleLabel(ActiveCell.Value2) Then
        lngGameCol = ActiveCell.Column
        lngExRow = ActiveCell.Row
        lngLastRow = wsWork.Cells(wsWork.Rows.Count, lngGameCol).End(xlUp).Row
        If lngLastRow > lngExRow + MAX_SCAN_ROWS Then lngLastRow = lngExRow + MAX_SCAN_ROWS
        lngTop = lngExRow
        lngBottom = lngLastRow
        strWhere = wsWork.Name & ", " & SafeText(ActiveCell.Value2)
        LocateLevel = True
        GoTo Cleanup
    End If

    lngGameCol = 2
    lngLastRow = wsWork.Cells(wsWork.Rows.Count, 2).End(xlUp).Row
    If lngLastRow < 2 Then
        strMsg = strCmd & ": " & wsWork.Name & " has nothing in column B."
        GoTo Cleanup
    End If
    If IsLevelSheetName(wsWork.Name) Then
        lngTop = 1
        lngBottom = lngLastRow
        strWhere = wsWork.Name
    Else
        Set colRows = LevelMarkerRows(wsWork)
        If colRows Is Nothing Then
            strMsg = strCmd & ": could not read the level headings on " & wsWork.Name & " - see " & LogLocation()
            GoTo Cleanup
        End If
        For lngI = 1 To colRows.Count
            If colRows(lngI) <= ActiveCell.Row Then lngLevel = lngI
        Next lngI
        If lngLevel = 0 Then
            strMsg = strCmd & ": select a cell inside a level (at or below its 'Level #' heading), or its 'Example#' cell."
            GoTo Cleanup
        End If
        lngTop = colRows(lngLevel)
        If lngLevel < colRows.Count Then
            lngBottom = colRows(lngLevel + 1) - 1
        Else
            ' The last level ends where a bonus section starts.
            lngBottom = lngLastRow
            lngBonusRow = BonusStartRow(wsWork, lngTop)
            If lngBonusRow > 0 Then
                If ActiveCell.Row >= lngBonusRow Then
                    strMsg = strCmd & ": that is the bonus section - select a cell inside a level."
                    GoTo Cleanup
                End If
                lngBottom = lngBonusRow - 1
            End If
        End If
        strWhere = wsWork.Name & ", " & SafeText(wsWork.Cells(lngTop, 2).Value2)
    End If

    ' The level's example row ("Example3" / "Example 3", preferring one with an Answer header).
    lngExRow = FindExampleRow(wsWork, lngTop, lngBottom)
    If lngExRow = 0 Then
        strMsg = strCmd & ": no 'Example#' row found in column B for " & strWhere & "."
        GoTo Cleanup
    End If
    LocateLevel = True

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strMsg = strCmd & " failed: " & Err.Description
    If LogError("LocateLevel", Err.Number, Err.Description) Then strMsg = strMsg & " - see " & LogLocation()
    LocateLevel = False
    Resume Cleanup
End Function

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Go To Example
' Description:            Goes to the level's example row, two cells right of its last input, ready to solve
' Macro Expression:       modDataTable.GoToExample()
' Generated:              2026-09-24 (replaces Select Example And Questions, Jaq)
'----------------------------------------------------------------------------------------------------
' Purpose: From anywhere in a level, select the cell two columns right of the level's last input on its
'          example row: where Jaq starts a live solve. Finds the level exactly as Create Data Table does
'          (LocateLevel) and the inputs by the one input rule (LevelInputColumns). Scrolls to it if it is
'          off screen. Selecting is the purpose of the command.
Public Sub GoToExample()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "GoToExample"
    Const CMD_TITLE As String = "Go To Example"

    Dim wsWork      As Worksheet
    Dim colIn       As Collection
    Dim rngLand     As Range
    Dim lngGameCol  As Long
    Dim lngExRow    As Long
    Dim lngTop      As Long
    Dim lngBottom   As Long
    Dim lngAnsCol   As Long
    Dim lngCol      As Long
    Dim strWhere    As String
    Dim strMsg      As String
    Dim strErr      As String

    On Error GoTo ErrHandler

    If TypeName(ActiveSheet) <> "Worksheet" Then
        strMsg = CMD_TITLE & ": select a cell in a level first."
        GoTo Cleanup
    End If
    Set wsWork = ActiveSheet
    If Not LocateLevel(wsWork, CMD_TITLE, lngGameCol, lngExRow, lngTop, lngBottom, strWhere, strMsg) Then GoTo Cleanup

    Set colIn = LevelInputColumns(wsWork, lngGameCol, lngExRow, lngAnsCol)
    If colIn Is Nothing Then
        strMsg = CMD_TITLE & ": could not read the inputs on the example row - see " & LogLocation()
        GoTo Cleanup
    End If
    If colIn.Count > 0 Then
        lngCol = colIn(colIn.Count) + 2
    Else
        lngCol = lngGameCol + 2
    End If

    Set rngLand = wsWork.Cells(lngExRow, lngCol)
    rngLand.Select
    ShowBlock rngLand, rngLand
    strMsg = CMD_TITLE & " (" & strWhere & "): " & rngLand.Address(False, False) & ", two right of the last input on " & _
             SafeText(wsWork.Cells(lngExRow, lngGameCol).Value2) & "."

Cleanup:
    On Error Resume Next
    If Len(strErr) > 0 Then
        Application.StatusBar = Left$(CMD_TITLE & " failed - " & strErr, m_MAX_STATUS)
    ElseIf Len(strMsg) > 0 Then
        Application.StatusBar = Left$(strMsg, m_MAX_STATUS)
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, strErr) Then strErr = strErr & " - see " & LogLocation()
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Link Answer Cell
' Description:            Points the level's data table ANSWER cell at the active cell (your calculation's output); then run RDT
' Macro Expression:       modDataTable.LinkAnswerCell()
' Generated:              2026-09-24 (CompBot's own take on A-ZTraining's ANS, Jaq)
'----------------------------------------------------------------------------------------------------
' Purpose: The step between Create Data Table and Run Data Table: stand on the cell holding the example
'          calculation's OUTPUT and run this; the ANSWER cell of THIS level's block (3 rows below the
'          block's input cell, 1 column right) becomes =that cell. The level is found exactly as Create
'          Data Table finds it (LocateLevel), so on a Case sheet with several blocks the right one is
'          used; on an L## sheet it is the sheet's block. The active cell stays where it is.
'          A-ZTraining's ANS searched the sheet for the first green ANSWER cell (an Excel What-If table)
'          and left the results copied; this one works on CompBot's own block and copies nothing.
Public Sub LinkAnswerCell()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "LinkAnswerCell"
    Const CMD_TITLE As String = "Link Answer Cell"

    Dim wsWork      As Worksheet
    Dim rngOutput   As Range
    Dim rngZone     As Range
    Dim rngInput    As Range
    Dim rngAnswer   As Range
    Dim blnFailed   As Boolean
    Dim lngGameCol  As Long
    Dim lngExRow    As Long
    Dim lngTop      As Long
    Dim lngBottom   As Long
    Dim strWhere    As String
    Dim strFound    As String
    Dim strMsg      As String
    Dim strErr      As String

    On Error GoTo ErrHandler

    If TypeName(ActiveSheet) <> "Worksheet" Then
        strMsg = CMD_TITLE & ": stand on your calculation's output cell first."
        GoTo Cleanup
    End If
    Set wsWork = ActiveSheet
    Set rngOutput = ActiveCell
    If Not LocateLevel(wsWork, CMD_TITLE, lngGameCol, lngExRow, lngTop, lngBottom, strWhere, strMsg) Then GoTo Cleanup

    ' This level's block: on a level sheet any block is the level's; elsewhere only one whose input
    ' cell lies in the level's rows (the same zone rule as Create Data Table).
    If IsLevelSheetName(wsWork.Name) Then
        Set rngZone = wsWork.Cells
    Else
        Set rngZone = wsWork.Range(wsWork.Cells(lngTop, 1), wsWork.Cells(lngBottom, wsWork.Columns.Count))
    End If
    Set rngInput = FindInputCell(wsWork, rngOutput, CMD_TITLE, strFound, rngZone, blnFailed)
    If blnFailed Then
        strMsg = strFound
        GoTo Cleanup
    End If
    If rngInput Is Nothing Then
        strMsg = CMD_TITLE & " (" & strWhere & "): this level has no data table block yet. Run Create Data Table (CDT) first."
        GoTo Cleanup
    End If

    Set rngAnswer = rngInput.Offset(m_FORMULA_OFFSET_ROWS, 1)
    ' The whole block: input row down to the last game, game, answer and SIDE columns (as Link Side Cell).
    If Not Application.Intersect(rngOutput, rngInput.Resize(m_BODY_OFFSET_ROWS + CountGames(rngInput.Offset(m_BODY_OFFSET_ROWS, 0)), _
                                                            2 + SideColumnCount(rngInput))) Is Nothing Then
        strMsg = CMD_TITLE & ": stand on your calculation's OUTPUT cell, not on the block itself."
        GoTo Cleanup
    End If
    rngAnswer.Formula = "=" & rngOutput.Address(True, True)
    strMsg = CMD_TITLE & " (" & strWhere & "): ANSWER " & rngAnswer.Address(False, False) & " = " & _
             rngOutput.Address(False, False) & ". Now run RDT."

Cleanup:
    On Error Resume Next
    If Len(strErr) > 0 Then
        Application.StatusBar = Left$(CMD_TITLE & " failed - " & strErr, m_MAX_STATUS)
    ElseIf Len(strMsg) > 0 Then
        Application.StatusBar = Left$(strMsg, m_MAX_STATUS)
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, strErr) Then strErr = strErr & " - see " & LogLocation()
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Link Side Cell
' Description:            Adds a SIDE column to the level's data table, pointed at the active cell, so RDT also records a second per-game value for a bonus
' Macro Expression:       modDataTable.LinkSideCell()
' Generated:              2026-09-24 (Jaq; spec commandtools\SPEC-side-bonus-column-260924.md)
'----------------------------------------------------------------------------------------------------
' Purpose: Some bonuses need a DIFFERENT per-game number from the same calculation (a count, a per-player
'          score), then a SUM / MAX / COUNTIF over the games. Stand on the cell holding that number and
'          run this: the next SIDE column of THIS level's block (found as Link Answer Cell finds it) is
'          added in the first column right of the block's last column:
'              Check row (input row + 2)   "SIDE k"
'              ANSWER row (input row + 3)  =the active cell, light blue
'              game rows                   #N/A until Run Data Table fills them, alongside the answer
'          Up to m_MAX_SIDES columns. The column must be empty from the Check row to the last game: if it
'          is not, nothing is written in it; an orange note goes in the nearest empty cell ABOVE it
'          instead, saying which cell is in the way. The next successful run clears that note.
'          Never moves the block, never overwrites. The active cell stays where it is.
Public Sub LinkSideCell()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "LinkSideCell"
    Const CMD_TITLE As String = "Link Side Cell"

    Dim wsWork      As Worksheet
    Dim rngOutput   As Range
    Dim rngZone     As Range
    Dim rngInput    As Range
    Dim rngSide     As Range
    Dim rngHead     As Range
    Dim rngBlocker  As Range
    Dim varVals     As Variant
    Dim varForm     As Variant
    Dim blnFailed   As Boolean
    Dim lngGameCol  As Long
    Dim lngExRow    As Long
    Dim lngTop      As Long
    Dim lngBottom   As Long
    Dim lngGames    As Long
    Dim lngSides    As Long
    Dim lngK        As Long
    Dim lngCol      As Long
    Dim lngR        As Long
    Dim strCol      As String
    Dim strNote     As String
    Dim strWhere    As String
    Dim strFound    As String
    Dim strMsg      As String
    Dim strErr      As String

    On Error GoTo ErrHandler

    If TypeName(ActiveSheet) <> "Worksheet" Then
        strMsg = CMD_TITLE & ": stand on the cell holding the side value first."
        GoTo Cleanup
    End If
    Set wsWork = ActiveSheet
    Set rngOutput = ActiveCell
    If Not LocateLevel(wsWork, CMD_TITLE, lngGameCol, lngExRow, lngTop, lngBottom, strWhere, strMsg) Then GoTo Cleanup

    ' This level's block - the same zone rule as Link Answer Cell.
    If IsLevelSheetName(wsWork.Name) Then
        Set rngZone = wsWork.Cells
    Else
        Set rngZone = wsWork.Range(wsWork.Cells(lngTop, 1), wsWork.Cells(lngBottom, wsWork.Columns.Count))
    End If
    Set rngInput = FindInputCell(wsWork, rngOutput, CMD_TITLE, strFound, rngZone, blnFailed)
    If blnFailed Then
        strMsg = strFound
        GoTo Cleanup
    End If
    If rngInput Is Nothing Then
        strMsg = CMD_TITLE & " (" & strWhere & "): this level has no data table block yet. Run Create Data Table (CDT) first."
        GoTo Cleanup
    End If

    lngGames = CountGames(rngInput.Offset(m_BODY_OFFSET_ROWS, 0))
    If lngGames = 0 Then
        strMsg = CMD_TITLE & ": no game numbers below " & rngInput.Address(False, False) & "."
        GoTo Cleanup
    End If
    lngSides = SideColumnCount(rngInput)

    ' At the cap there is no new column, so say that first.
    If lngSides >= m_MAX_SIDES Then
        strMsg = CMD_TITLE & ": this block already has " & m_MAX_SIDES & " SIDE columns, the most it takes."
        GoTo Cleanup
    End If
    ' The block, INCLUDING the column the new SIDE would take: a head pointing into its own column
    ' would be circular, or overwritten by the label / #N/A body.
    If Not Application.Intersect(rngOutput, rngInput.Resize(m_BODY_OFFSET_ROWS + lngGames, 3 + lngSides)) Is Nothing Then
        strMsg = CMD_TITLE & ": stand on the cell in YOUR calculation that holds the side value, not on the block itself."
        GoTo Cleanup
    End If
    If wsWork.ProtectContents Then
        strMsg = CMD_TITLE & ": " & wsWork.Name & " is protected - unprotect it first."
        GoTo Cleanup
    End If

    lngK = lngSides + 1
    lngCol = rngInput.Column + 1 + lngK
    If lngCol > wsWork.Columns.Count Then
        strMsg = CMD_TITLE & ": no column left on the sheet for SIDE " & lngK & "."
        GoTo Cleanup
    End If
    strCol = ColumnLetters(wsWork, lngCol)

    ' The new column, Check row down to the last game, must be empty: one read of values and formulas.
    Set rngSide = wsWork.Range(wsWork.Cells(rngInput.Row + 2, lngCol), _
                               wsWork.Cells(rngInput.Row + m_FORMULA_OFFSET_ROWS + lngGames, lngCol))
    varVals = rngSide.Value2
    varForm = rngSide.Formula
    For lngR = 1 To UBound(varVals, 1)
        If Not IsEmpty(varVals(lngR, 1)) Or Len(CStr(varForm(lngR, 1))) > 0 Then
            Set rngBlocker = rngSide.Cells(lngR, 1)
            Exit For
        End If
    Next lngR
    If rngBlocker Is Nothing Then
        If IsNull(rngSide.MergeCells) Then
            Set rngBlocker = rngSide.Cells(1, 1)
        ElseIf rngSide.MergeCells Then
            Set rngBlocker = rngSide.Cells(1, 1)
        End If
    End If

    If Not rngBlocker Is Nothing Then
        ' Refuse, and say so on the sheet above the column (Jaq, 2026-09-24).
        strNote = m_NOTE_PREFIX & lngK & ": " & rngBlocker.Address(False, False) & " is in use. Clear " & _
                  rngSide.Address(False, False) & ", or move your calculation, then run LSC again."
        If WriteNoRoomNote(wsWork, lngCol, rngInput.Row + 1, rngInput.Row - 1, strNote) Then
            strMsg = CMD_TITLE & ": no room for SIDE " & lngK & " in column " & strCol & " (" & _
                     rngBlocker.Address(False, False) & " is in use) - nothing linked, see the orange note."
        Else
            strMsg = CMD_TITLE & ": no room for SIDE " & lngK & " in column " & strCol & " (" & _
                     rngBlocker.Address(False, False) & " is in use) - nothing linked. Clear " & _
                     rngSide.Address(False, False) & " or move your calculation."
        End If
        GoTo Cleanup
    End If

    ' Write it: label, head, #N/A body. Then clear any old no-room note above the column.
    rngSide.Cells(1, 1).Value2 = m_SIDE_PREFIX & CStr(lngK)
    Set rngHead = rngSide.Cells(2, 1)
    rngHead.Formula = "=" & rngOutput.Address(True, True)
    rngHead.Interior.Color = m_SIDE_FILL
    rngSide.Offset(2, 0).Resize(lngGames, 1).Value2 = CVErr(xlErrNA)
    ClearSideNotes wsWork, lngCol, rngInput.Row + 1, rngInput.Row - 1

    strMsg = CMD_TITLE & " (" & strWhere & "): SIDE " & lngK & " " & rngHead.Address(False, False) & " = " & _
             rngOutput.Address(False, False) & ". Run RDT: it fills the answer and every SIDE column together."

Cleanup:
    On Error Resume Next
    If Len(strErr) > 0 Then
        Application.StatusBar = Left$(CMD_TITLE & " failed - " & strErr, m_MAX_STATUS)
    ElseIf Len(strMsg) > 0 Then
        Application.StatusBar = Left$(strMsg, m_MAX_STATUS)
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, strErr) Then strErr = strErr & " - see " & LogLocation()
    Resume Cleanup
End Sub

' Purpose: How many SIDE columns a block has: the run of columns right of the answer column whose Check-row
'          cell (input row + 2) reads "SIDE 1", "SIDE 2", ... in order. 0 if none, or if it cannot be read.
Private Function SideColumnCount(ByVal rngInput As Range) As Long

    Dim ws As Worksheet
    Dim varLabel As Variant
    Dim lngK As Long
    Dim lngCol As Long

    On Error GoTo ErrHandler

    Set ws = rngInput.Parent
    For lngK = 1 To m_MAX_SIDES
        lngCol = rngInput.Column + 1 + lngK
        If lngCol > ws.Columns.Count Then Exit For
        varLabel = ws.Cells(rngInput.Row + 2, lngCol).Value2
        If VarType(varLabel) <> vbString Then Exit For
        If StrComp(Trim$(varLabel), m_SIDE_PREFIX & CStr(lngK), vbTextCompare) <> 0 Then Exit For
        SideColumnCount = lngK
    Next lngK

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    SideColumnCount = 0
    Resume Cleanup
End Function

' Purpose: The name of output column lngJ of a block, for messages: "ANSWER" or "SIDE k".
Private Function OutputTitle(ByVal lngJ As Long) As String
    On Error GoTo ErrHandler
    If lngJ <= 1 Then
        OutputTitle = "ANSWER"
    Else
        OutputTitle = m_SIDE_PREFIX & CStr(lngJ - 1)
    End If
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    OutputTitle = "column " & lngJ
End Function

' Purpose: Column lngJ of a games x columns result array, as a games x 1 array ready to write.
'          The caller's own handler deals with a failure.
Private Function OutputColumn(ByRef varOut() As Variant, ByVal lngJ As Long) As Variant

    Dim varCol() As Variant
    Dim lngI As Long

    On Error GoTo ErrHandler

    ReDim varCol(1 To UBound(varOut, 1), 1 To 1)
    For lngI = 1 To UBound(varOut, 1)
        varCol(lngI, 1) = varOut(lngI, lngJ)
    Next lngI
    OutputColumn = varCol
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Err.Raise Err.Number, "OutputColumn", Err.Description
End Function

' Purpose: A column's letters ("M") from its number.
Private Function ColumnLetters(ByVal ws As Worksheet, ByVal lngCol As Long) As String
    On Error GoTo ErrHandler
    ColumnLetters = Split(ws.Cells(1, lngCol).Address(True, False), "$")(0)
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ColumnLetters = CStr(lngCol)
End Function

' Purpose: Write Link Side Cell's refusal note in burnt orange in the nearest EMPTY, unmerged cell of column
'          lngCol from lngFromRow up to lngFloorRow (the block's own rows above the SIDE range: its Expected,
'          input and header rows), so a note never lands in another level's area. Any older note there is
'          cleared first, so there is only ever one. Returns False if no cell was free or the write failed.
Private Function WriteNoRoomNote(ByVal ws As Worksheet, ByVal lngCol As Long, ByVal lngFromRow As Long, _
                                 ByVal lngFloorRow As Long, ByVal strNote As String) As Boolean

    Dim rngAbove As Range
    Dim varVals As Variant
    Dim varForm As Variant
    Dim lngR As Long

    On Error GoTo ErrHandler

    If lngFloorRow < 1 Then lngFloorRow = 1
    If lngFromRow < lngFloorRow Then GoTo Cleanup
    ClearSideNotes ws, lngCol, lngFromRow, lngFloorRow

    Set rngAbove = ws.Range(ws.Cells(lngFloorRow, lngCol), ws.Cells(lngFromRow, lngCol))
    If lngFromRow = lngFloorRow Then
        ReDim varVals(1 To 1, 1 To 1)
        ReDim varForm(1 To 1, 1 To 1)
        varVals(1, 1) = rngAbove.Value2
        varForm(1, 1) = rngAbove.Formula
    Else
        varVals = rngAbove.Value2
        varForm = rngAbove.Formula
    End If
    For lngR = lngFromRow - lngFloorRow + 1 To 1 Step -1
        If IsEmpty(varVals(lngR, 1)) And Len(CStr(varForm(lngR, 1))) = 0 Then
            If Not ws.Cells(lngFloorRow + lngR - 1, lngCol).MergeCells Then
                With ws.Cells(lngFloorRow + lngR - 1, lngCol)
                    .Value2 = strNote
                    .Font.Color = m_NOTE_COLOUR
                End With
                WriteNoRoomNote = True
                GoTo Cleanup
            End If
        End If
    Next lngR

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "WriteNoRoomNote", Err.Number, Err.Description
    WriteNoRoomNote = False
    Resume Cleanup
End Function

' Purpose: Clear every Link Side Cell refusal note in column lngCol from lngFloorRow to lngToRow (text
'          starting m_NOTE_PREFIX), contents and font colour. Only this block's rows, so another level's
'          note in the same column is left alone. Cosmetic: a failure is logged and otherwise ignored.
Private Sub ClearSideNotes(ByVal ws As Worksheet, ByVal lngCol As Long, ByVal lngToRow As Long, _
                           ByVal lngFloorRow As Long)

    Dim rngAbove As Range
    Dim varVals As Variant
    Dim lngR As Long

    On Error GoTo ErrHandler

    If lngFloorRow < 1 Then lngFloorRow = 1
    If lngToRow < lngFloorRow Then GoTo Cleanup
    Set rngAbove = ws.Range(ws.Cells(lngFloorRow, lngCol), ws.Cells(lngToRow, lngCol))
    If lngToRow = lngFloorRow Then
        ReDim varVals(1 To 1, 1 To 1)
        varVals(1, 1) = rngAbove.Value2
    Else
        varVals = rngAbove.Value2
    End If
    For lngR = 1 To lngToRow - lngFloorRow + 1
        If VarType(varVals(lngR, 1)) = vbString Then
            If Left$(varVals(lngR, 1), Len(m_NOTE_PREFIX)) = m_NOTE_PREFIX Then
                With ws.Cells(lngFloorRow + lngR - 1, lngCol)
                    .ClearContents
                    .Font.ColorIndex = xlColorIndexAutomatic
                End With
            End If
        End If
    Next lngR

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "ClearSideNotes", Err.Number, Err.Description
    Resume Cleanup
End Sub

' Purpose: Where a new data table block goes: level with the example row, at lngFirstCol (two columns
'          right of the last input) if the block's whole footprint is empty there, else the nearest
'          column to the right where it is. The footprint is the block's rows (the header row above the
'          input cell down to the last game) across its lngWidth columns, plus one column to its left,
'          so the block never sits flush against other content. A cell counts as used if it holds a
'          formula or shows a value (so spilled arrays count). Returns Nothing if the sheet has no room
'          or the read fails (logged).
Private Function BlockTarget(ByVal ws As Worksheet, ByVal lngExRow As Long, ByVal lngFirstCol As Long, _
                             ByVal lngWidth As Long, ByVal lngGames As Long) As Range

    Dim rngBand As Range
    Dim rngLast As Range
    Dim varVals As Variant
    Dim varForm As Variant
    Dim blnUsed() As Boolean
    Dim lngTopRow As Long
    Dim lngBotRow As Long
    Dim lngLeft As Long
    Dim lngLastUsed As Long
    Dim lngR As Long
    Dim lngC As Long
    Dim lngCol As Long
    Dim blnClear As Boolean

    On Error GoTo ErrHandler

    lngTopRow = lngExRow - 1
    If lngTopRow < 1 Then lngTopRow = 1
    lngBotRow = lngExRow + m_BODY_OFFSET_ROWS - 1 + lngGames
    If lngBotRow > ws.Rows.Count Then lngBotRow = ws.Rows.Count
    lngLeft = lngFirstCol - 1
    If lngLeft < 1 Then lngLeft = 1

    ' Last used column in the band, from the gap column rightwards.
    Set rngBand = ws.Range(ws.Cells(lngTopRow, lngLeft), ws.Cells(lngBotRow, ws.Columns.Count))
    lngLastUsed = 0
    Set rngLast = rngBand.Find(What:="*", After:=rngBand.Cells(1, 1), LookIn:=xlFormulas, _
                               SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not rngLast Is Nothing Then lngLastUsed = rngLast.Column
    Set rngLast = rngBand.Find(What:="*", After:=rngBand.Cells(1, 1), LookIn:=xlValues, _
                               SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If Not rngLast Is Nothing Then
        If rngLast.Column > lngLastUsed Then lngLastUsed = rngLast.Column
    End If
    If lngLastUsed < lngLeft Then
        Set BlockTarget = ws.Cells(lngExRow, lngFirstCol)
        GoTo Cleanup
    End If

    ' One read of the band, then a used / unused flag per column.
    Set rngBand = ws.Range(ws.Cells(lngTopRow, lngLeft), ws.Cells(lngBotRow, lngLastUsed))
    varVals = rngBand.Value2
    varForm = rngBand.Formula
    If Not IsArray(varVals) Then
        ReDim varVals(1 To 1, 1 To 1)
        ReDim varForm(1 To 1, 1 To 1)
        varVals(1, 1) = rngBand.Value2
        varForm(1, 1) = rngBand.Formula
    End If
    ReDim blnUsed(lngLeft To lngLastUsed)
    For lngC = 1 To UBound(varVals, 2)
        For lngR = 1 To UBound(varVals, 1)
            If Not IsEmpty(varVals(lngR, lngC)) Or Len(CStr(varForm(lngR, lngC))) > 0 Then
                blnUsed(lngLeft + lngC - 1) = True
                Exit For
            End If
        Next lngR
    Next lngC

    ' The nearest start column whose gap column and block columns are all unused.
    For lngCol = lngFirstCol To lngLastUsed + 2
        If lngCol + lngWidth - 1 > ws.Columns.Count Then Exit For
        blnClear = True
        For lngC = lngCol - 1 To lngCol + lngWidth - 1
            If lngC >= lngLeft And lngC <= lngLastUsed Then
                If blnUsed(lngC) Then
                    blnClear = False
                    Exit For
                End If
            End If
        Next lngC
        If blnClear Then
            Set BlockTarget = ws.Cells(lngExRow, lngCol)
            GoTo Cleanup
        End If
    Next lngCol

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "BlockTarget", Err.Number, Err.Description
    Set BlockTarget = Nothing
    Resume Cleanup
End Function

' Purpose: The last input cell of a block: the run of XLOOKUP cells right of the input cell that read
'          the game from it. With no inputs, the cell right of the input cell (the old end state).
Private Function LastInputCell(ByVal rngInput As Range) As Range

    Dim ws As Worksheet
    Dim strAddr As String
    Dim strForm As String
    Dim lngCol As Long

    On Error GoTo ErrHandler

    Set ws = rngInput.Parent
    Set LastInputCell = rngInput.Offset(0, 1)
    strAddr = rngInput.Address(True, True)
    lngCol = rngInput.Column + 1
    Do While lngCol <= ws.Columns.Count
        strForm = CStr(ws.Cells(rngInput.Row, lngCol).Formula)
        If InStr(1, strForm, "XLOOKUP(", vbTextCompare) = 0 Then Exit Do
        If Not RefersToCell(strForm, strAddr) Then Exit Do
        Set LastInputCell = ws.Cells(rngInput.Row, lngCol)
        lngCol = lngCol + 1
    Loop

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "LastInputCell", Err.Number, Err.Description
    Set LastInputCell = rngInput.Offset(0, 1)
    Resume Cleanup
End Function

' Purpose: Scroll the active window so the block (its header row down to ANSWER, from the input cell
'          to the landing cell and one column beyond) is on screen. Only scrolls when part of it is
'          off screen, so a block already in view does not jump. Cosmetic: a failure is logged and
'          otherwise ignored.
Private Sub ShowBlock(ByVal rngInput As Range, ByVal rngLand As Range)

    Dim ws As Worksheet
    Dim rngShow As Range
    Dim rngSeen As Range
    Dim lngCol As Long
    Dim lngRow As Long

    On Error GoTo ErrHandler

    Set ws = rngInput.Parent
    Set rngShow = ws.Range(ws.Cells(rngInput.Row - 1, rngInput.Column), _
                           ws.Cells(rngInput.Row + m_FORMULA_OFFSET_ROWS, rngLand.Column + 1))
    With ActiveWindow
        Set rngSeen = Intersect(.VisibleRange, rngShow)
        If Not rngSeen Is Nothing Then
            If rngSeen.Address = rngShow.Address Then GoTo Cleanup
        End If
        ' Left edge of the block two columns in; header row two rows down.
        lngCol = rngInput.Column - 2
        If lngCol < 1 Then lngCol = 1
        lngRow = rngInput.Row - 3
        If lngRow < 1 Then lngRow = 1
        If rngShow.Column < .VisibleRange.Column Or _
           rngShow.Column + rngShow.Columns.Count > .VisibleRange.Column + .VisibleRange.Columns.Count Then
            .ScrollColumn = lngCol
        End If
        If rngShow.Row < .VisibleRange.Row Or _
           rngShow.Row + rngShow.Rows.Count > .VisibleRange.Row + .VisibleRange.Rows.Count Then
            .ScrollRow = lngRow
        End If
    End With

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "ShowBlock", Err.Number, Err.Description
    Resume Cleanup
End Sub

' Purpose: The sheet a link formula such as ='L01'!$E$14 or =L01!E14 points at, or "".
Private Function LinkedSheetName(ByVal strFormula As String) As String

    Dim lngBang As Long
    Dim strSheet As String

    On Error GoTo ErrHandler

    If Left$(strFormula, 1) <> "=" Then GoTo Cleanup
    lngBang = InStr(1, strFormula, "!")
    If lngBang = 0 Then GoTo Cleanup
    strSheet = Mid$(strFormula, 2, lngBang - 2)
    If Left$(strSheet, 1) = "'" And Right$(strSheet, 1) = "'" And Len(strSheet) >= 2 Then
        strSheet = Replace(Mid$(strSheet, 2, Len(strSheet) - 2), "''", "'")
    End If
    ' A plain link only - a formula with functions or operators before the "!" is not a link.
    If strSheet Like "*[(,+*/&=<>]*" Then GoTo Cleanup
    LinkedSheetName = strSheet

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LinkedSheetName = vbNullString
    Resume Cleanup
End Function

' Purpose: The hidden sheet-level name that remembers this input cell's "Example#" label, or Nothing.
'          The name REFERS TO the input cell (so it follows the cell when rows or columns move) and
'          keeps the label in its Comment. Found by where it points, not by its own text, so a stale
'          name left at the same address by another block can never be picked up.
Private Function LabelNameFor(ByVal rngInput As Range) As Name

    Dim nmAny As Name
    Dim rngRef As Range

    On Error GoTo ErrHandler

    For Each nmAny In rngInput.Parent.Names
        If InStr(1, nmAny.Name, m_LABEL_PREFIX, vbTextCompare) > 0 Then
            Set rngRef = Nothing
            ' Probe: a name whose target was deleted has no RefersToRange.
            On Error Resume Next
            Set rngRef = nmAny.RefersToRange
            On Error GoTo ErrHandler
            If Not rngRef Is Nothing Then
                If rngRef.Parent Is rngInput.Parent And rngRef.Address = rngInput.Address Then
                    Set LabelNameFor = nmAny
                    Exit Function
                End If
            End If
        End If
    Next nmAny
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Set LabelNameFor = Nothing
End Function

' Purpose: The block's "Example#" label: the input cell's text if it still holds it, else the
'          remembered label, else "".
Private Function ExampleLabel(ByVal rngInput As Range) As String

    Dim nmLabel As Name

    On Error GoTo ErrHandler

    If VarType(rngInput.Value2) = vbString Then
        ExampleLabel = rngInput.Value2
        Exit Function
    End If

    Set nmLabel = LabelNameFor(rngInput)
    If Not nmLabel Is Nothing Then ExampleLabel = nmLabel.Comment
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ExampleLabel = vbNullString
End Function

' Purpose: Make sure the block's "Example#" label is remembered before a game overwrites it.
'          Returns True when the label is safely remembered (already, or now); False if not - the
'          caller must then leave the input cell alone.
Private Function RememberLabel(ByVal rngInput As Range) As Boolean

    Dim nmLabel As Name
    Dim strSheet As String
    Dim strBase As String
    Dim strName As String
    Dim lngN As Long

    On Error GoTo ErrHandler

    If VarType(rngInput.Value2) <> vbString Then
        ' A game is already loaded: fine only if the label was remembered earlier.
        RememberLabel = Not (LabelNameFor(rngInput) Is Nothing)
        Exit Function
    End If

    Set nmLabel = LabelNameFor(rngInput)
    If nmLabel Is Nothing Then
        ' A name with this text may already exist for another block that has since moved (rows were
        ' inserted) - Names.Add would silently re-point it. Pick an unused name instead.
        strBase = m_LABEL_PREFIX & Replace(rngInput.Address(False, False), "$", "")
        strName = strBase
        lngN = 1
        Do While SheetNameExists(rngInput.Parent, strName)
            lngN = lngN + 1
            strName = strBase & "_" & lngN
        Loop
        strSheet = Replace(rngInput.Parent.Name, "'", "''")
        Set nmLabel = rngInput.Parent.Names.Add(Name:=strName, _
                                                RefersTo:="='" & strSheet & "'!" & rngInput.Address(True, True), _
                                                Visible:=False)
    End If
    nmLabel.Comment = CStr(rngInput.Value2)
    RememberLabel = (nmLabel.Comment = CStr(rngInput.Value2))
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "RememberLabel", Err.Number, Err.Description
    RememberLabel = False
End Function

' Purpose: True if WS has a sheet-level name with exactly this text (Worksheet.Names reports them as
'          "'Sheet'!Name" or "Sheet!Name").
Private Function SheetNameExists(ByVal ws As Worksheet, ByVal strName As String) As Boolean

    Dim nmAny As Name
    Dim strFull As String

    On Error GoTo ErrHandler

    For Each nmAny In ws.Names
        strFull = nmAny.Name
        If InStr(1, strFull, "!") > 0 Then strFull = Mid$(strFull, InStrRev(strFull, "!") + 1)
        If StrComp(strFull, strName, vbTextCompare) = 0 Then
            SheetNameExists = True
            GoTo Cleanup
        End If
    Next nmAny

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    SheetNameExists = False
    Resume Cleanup
End Function

' Purpose: Remove the remembered label once the input cell holds it again, so no hidden names pile up.
Private Sub ForgetLabel(ByVal rngInput As Range)

    Dim nmLabel As Name

    On Error GoTo ErrHandler

    Set nmLabel = LabelNameFor(rngInput)
    If Not nmLabel Is Nothing Then nmLabel.Delete
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
End Sub

' Purpose: " Example check: MATCH." from the Check row Create Data Table puts above the table,
'          or "" for a block built before the check existed.
Private Function ExampleCheckText(ByVal rngInput As Range) As String

    Dim rngLabel As Range

    On Error GoTo ErrHandler

    Set rngLabel = rngInput.Offset(2, 0)
    If VarType(rngLabel.Value2) = vbString Then
        If rngLabel.Value2 = "Check" Then
            ExampleCheckText = " Example check: " & SafeText(rngLabel.Offset(0, 1).Value2) & "."
        End If
    End If
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ExampleCheckText = vbNullString
End Function

' Purpose: A value as text, with error values shown as "#error".
Private Function SafeText(ByVal varValue As Variant) As String
    On Error GoTo ErrHandler
    If IsError(varValue) Then
        SafeText = "#error"
    Else
        SafeText = CStr(varValue)
    End If
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    SafeText = vbNullString
End Function

' Purpose: Decide Run Data Table's per-game recalc scope. Returns m_SCOPE_SHEET when recalculating only
'          wsCalc is proven to give the same answers as a full recalculation AND is worth it, else
'          m_SCOPE_FULL. strWhy says which, and why (shown on the StatusBar).
'          rngFormula is the whole output row (ANSWER and any SIDE heads): the probes compare every cell
'          of it and the structural scan follows the precedents of all of them, because a side output
'          can depend on a sheet the answer does not.
'            1. First probe, and worth it? - game 1 is calculated in full (timed: the loop's per-game full
'               cost), the input is parked on a sentinel value and calculated in full, then game 1 is
'               read again with a sheet-only recalculation (timed). The two answers must agree. What the
'               loop would save, less the second probe, is the budget for step 2; if it does not beat
'               SCAN_ALLOWANCE the full recalc is used - a light level gains nothing.
'            2. Structural - CrossSheetFeedback: nothing the solve depends on (on other sheets, followed
'               sheet to sheet) reads back from wsCalc, and nothing on that path is out of the scan's
'               sight (INDIRECT, another workbook, a sheet reference it cannot resolve). This catches a
'               helper that only matters for some games, which the two probes would miss. It gives up,
'               and gives full, once it runs past the budget.
'            3. Second probe - the last game, the same way.
'          The sentinel parking means any stale cell elsewhere holds the same known state during the
'          sheet-only probes as during the loop (the sheet-only steps never touch other sheets).
'          Esc (error 18) is passed up to Run Data Table; any other error is logged and gives full.
'          Leaves the input cell holding the game last probed and other sheets parked - the loop overwrites
'          the input, and Run Data Table recalculates in full on every way out.
Private Function ChooseRecalcScope(ByVal wsCalc As Worksheet, ByVal rngInput As Range, _
                                   ByVal rngFormula As Range, ByVal varGames As Variant, _
                                   ByRef strWhy As String) As Long

    ' --- CONSTANTS (local to function) ---
    Const PROBE_SENTINEL As Double = -987654321#
    Const MIN_GAMES As Long = 4             ' below this the check can never pay for itself
    Const SCAN_ALLOWANCE As Double = 0.25   ' seconds the structural scan may need on a large sheet

    Dim lngCount As Long
    Dim lngPick As Long
    Dim lngI As Long
    Dim varSheet As Variant
    Dim varFull As Variant
    Dim strVia As String
    Dim dblT0 As Double
    Dim dblT1 As Double
    Dim dblSheetSecs As Double
    Dim dblFullSecs As Double
    Dim dblBudget As Double

    On Error GoTo ErrHandler

    ChooseRecalcScope = m_SCOPE_FULL
    lngCount = UBound(varGames, 1)
    If lngCount < MIN_GAMES Then
        strWhy = "full - too few games to gain from a sheet-only recalc"
        GoTo Cleanup
    End If

    For lngPick = 1 To 2
        If lngPick = 1 Then lngI = 1 Else lngI = lngCount

        ' The true answer, from a full recalculation.
        rngInput.Value2 = varGames(lngI, 1)
        dblT0 = HighResSeconds()
        Application.Calculate
        dblT1 = HighResSeconds()
        If dblT0 < 0 Or dblT1 < 0 Then
            strWhy = "full - no high-resolution timer to size the check"
            GoTo Cleanup
        End If
        If lngPick = 1 Then dblFullSecs = dblT1 - dblT0
        If Application.CalculationState <> xlDone Then
            strWhy = "full - calculation did not finish"
            GoTo Cleanup
        End If
        varFull = rngFormula.Value2

        ' Park everything on the sentinel, then read the same game with a sheet-only recalculation.
        If Not ParkOnSentinel(rngInput, PROBE_SENTINEL) Then
            strWhy = "full - calculation did not finish"
            GoTo Cleanup
        End If
        rngInput.Value2 = varGames(lngI, 1)
        ' No CalculationState test after the sheet-only calculation: it reads "pending" whenever
        ' anything outside wsCalc is dirty. Interrupt keys are off, so it always runs to the end.
        dblT0 = HighResSeconds()
        wsCalc.Calculate
        dblT1 = HighResSeconds()
        If dblT0 < 0 Or dblT1 < 0 Then
            strWhy = "full - no high-resolution timer to size the check"
            GoTo Cleanup
        End If
        If lngPick = 1 Then dblSheetSecs = dblT1 - dblT0
        varSheet = rngFormula.Value2

        If Not SameValues(varSheet, varFull) Then
            strWhy = "full - game " & SafeText(varGames(lngI, 1)) & " differs on a sheet-only recalc"
            GoTo Cleanup
        End If

        If lngPick = 1 Then
            ' Worth it? What the loop would save, less the second probe, is the scan's budget.
            dblBudget = lngCount * (dblFullSecs - dblSheetSecs) - (2 * dblFullSecs + dblSheetSecs)
            If dblBudget <= SCAN_ALLOWANCE Then
                strWhy = "full - this solve recalculates too fast to gain from a sheet-only recalc"
                GoTo Cleanup
            End If

            ' Structural check.
            strVia = CrossSheetFeedback(wsCalc, rngFormula, dblBudget)
            If Len(strVia) > 0 Then
                strWhy = "full - " & strVia
                GoTo Cleanup
            End If
        End If
    Next lngPick

    ChooseRecalcScope = m_SCOPE_SHEET
    strWhy = "sheet only"

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Esc: raised again from inside the handler, so it reaches Run Data Table's handler and cancels the run.
    If Err.Number = 18 Then Err.Raise 18
    If LogError("ChooseRecalcScope", Err.Number, Err.Description) Then
        strWhy = "full - the recalc check failed, see " & LogLocation()
    Else
        strWhy = "full - the recalc check failed"
    End If
    ChooseRecalcScope = m_SCOPE_FULL
    Resume Cleanup
End Function

' Purpose: Put the sentinel in the input cell and calculate everything, so every cell that depends on the
'          input holds one known state. Returns False if the calculation did not finish.
Private Function ParkOnSentinel(ByVal rngInput As Range, ByVal dblSentinel As Double) As Boolean

    On Error GoTo ErrHandler

    rngInput.Value2 = dblSentinel
    Application.Calculate
    ParkOnSentinel = (Application.CalculationState = xlDone)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If Err.Number = 18 Then Err.Raise 18
    ParkOnSentinel = False
    Resume Cleanup
End Function

' Purpose: Seconds from the high-resolution performance counter (Timer is too coarse to time one recalc).
'          Returns -1 when the counter is unavailable (e.g. Excel for Mac) - callers must check. Never mixes
'          in Timer, which runs on a different clock.
Private Function HighResSeconds() As Double

    Dim curCount As Currency
    Dim curFreq As Currency

    On Error GoTo ErrHandler

    HighResSeconds = -1
    If QueryPerformanceCounter(curCount) = 0 Then GoTo Cleanup
    If QueryPerformanceFrequency(curFreq) = 0 Then GoTo Cleanup
    If curFreq > 0 Then HighResSeconds = CDbl(curCount) / CDbl(curFreq)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If Err.Number = 18 Then Err.Raise 18
    HighResSeconds = -1
    Resume Cleanup
End Function

' Purpose: Why a sheet-only recalc of wsCalc cannot be trusted for the solve, or "" if it can.
'          The solve is rngFormula and its precedents on wsCalc. The sheets those formulas reference are
'          followed sheet to sheet (a whole sheet's formulas count once it is reached); if any reached
'          sheet references wsCalc, the solve's inputs can change during a game on another sheet, and a
'          sheet-only recalc would read them stale. Links from other sheets that the solve never reaches
'          (answer and score links on the Case or level sheets) do not count.
'          Defined names count in every scope (a sheet-scoped and a global "Rate" are merged - an over-count
'          is safe), including names built on other names; tables count as names; a 3D reference counts
'          every sheet in its span.
'          Anything the scan cannot follow returns a reason instead of "": INDIRECT, another workbook, a
'          sheet reference that does not resolve, a block too large to read, running past dblBudgetSecs,
'          or an error. Esc (error 18) is passed up.
Private Function CrossSheetFeedback(ByVal wsCalc As Worksheet, ByVal rngFormula As Range, _
                                    ByVal dblBudgetSecs As Double) As String

    Dim wbCalc As Workbook
    Dim reRef As Object
    Dim reText As Object
    Dim dicNameText As Object       ' UPPER name -> its RefersTo text, all scopes joined
    Dim dicNameSheets As Object     ' UPPER name or table -> "|KEY|KEY|" sheet keys behind it
    Dim dicReached As Object        ' UPPER sheet name -> True (sheets the solve depends on)
    Dim dicKeys As Object
    Dim dicSheets As Object
    Dim colQueue As Collection
    Dim rngSolve As Range
    Dim rngPrec As Range
    Dim nmAny As Name
    Dim wsAny As Worksheet
    Dim loAny As ListObject
    Dim wsNext As Worksheet
    Dim varKey As Variant
    Dim strName As String
    Dim strCalcUpper As String
    Dim strProblem As String
    Dim dblStart As Double
    Dim lngPass As Long
    Dim blnChanged As Boolean

    On Error GoTo ErrHandler

    dblStart = HighResSeconds()
    Set wbCalc = wsCalc.Parent
    strCalcUpper = UCase$(wsCalc.Name)

    ' 'Sheet name'! or SheetName! - group 1 the quoted form, group 2 the plain form. Either may be a 3D
    ' span (First:Last) or carry a [Book] prefix. The plain form stops at any operator or separator, so
    ' unquoted non-ASCII names are read whole.
    Set reRef = CreateObject("VBScript.RegExp")
    reRef.Global = True
    reRef.Pattern = "(?:'((?:[^']|'')+)'|([^\s'!""(),;:=+\-*/&^<>{}%#$@]+(?::[^\s'!""(),;:=+\-*/&^<>{}%#$@]+)?))!"

    ' String literals, removed before matching ("Total!" in a text is not a reference).
    Set reText = CreateObject("VBScript.RegExp")
    reText.Global = True
    reText.Pattern = """(?:[^""]|"""")*"""

    ' Defined names, every scope, joined by name.
    Set dicNameText = CreateObject("Scripting.Dictionary")
    For Each nmAny In wbCalc.Names
        strName = nmAny.Name
        If InStr(1, strName, "!") > 0 Then strName = Mid$(strName, InStrRev(strName, "!") + 1)
        strName = UCase$(strName)
        If Left$(strName, 6) <> "_XLNM." Then
            If dicNameText.Exists(strName) Then
                dicNameText(strName) = dicNameText(strName) & " " & CStr(nmAny.RefersTo)
            Else
                dicNameText(strName) = CStr(nmAny.RefersTo)
            End If
        End If
    Next nmAny

    ' The sheets each name refers to directly; tables are used like names.
    Set dicNameSheets = CreateObject("Scripting.Dictionary")
    For Each varKey In dicNameText.Keys
        Set dicKeys = CreateObject("Scripting.Dictionary")
        AddSheetKeys CStr(dicNameText(varKey)), reRef, reText, Nothing, dicKeys
        AddNameSheets dicNameSheets, CStr(varKey), dicKeys
    Next varKey
    For Each wsAny In wbCalc.Worksheets
        For Each loAny In wsAny.ListObjects
            dicNameSheets(UCase$(loAny.Name)) = "|" & UCase$(wsAny.Name) & "|"
        Next loAny
    Next wsAny

    ' Names built on names: fold each name's sheets into the names that use it, until nothing changes.
    Do
        blnChanged = False
        lngPass = lngPass + 1
        For Each varKey In dicNameText.Keys
            Set dicKeys = CreateObject("Scripting.Dictionary")
            AddSheetKeys CStr(dicNameText(varKey)), reRef, reText, dicNameSheets, dicKeys
            If AddNameSheets(dicNameSheets, CStr(varKey), dicKeys) Then blnChanged = True
        Next varKey
        If OverBudget(dblStart, dblBudgetSecs) Then
            CrossSheetFeedback = "the recalc check ran out of time"
            GoTo Cleanup
        End If
    Loop While blnChanged And lngPass <= dicNameText.Count

    ' The solve: the ANSWER cell and its precedents on wsCalc. Precedents raises when there are none.
    Set rngSolve = rngFormula
    Set rngPrec = Nothing
    On Error Resume Next
    Set rngPrec = rngFormula.Precedents
    On Error GoTo ErrHandler
    If Not rngPrec Is Nothing Then Set rngSolve = Application.Union(rngFormula, rngPrec)

    ' Sheets the solve references directly.
    Set dicKeys = CreateObject("Scripting.Dictionary")
    strProblem = SolveSheetKeys(rngSolve, reRef, reText, dicNameSheets, dicKeys)
    Set dicSheets = CreateObject("Scripting.Dictionary")
    If Len(strProblem) = 0 Then strProblem = ExpandSheetKeys(wbCalc, wsCalc.Name, dicKeys, dicSheets)
    If Len(strProblem) = 0 And OverBudget(dblStart, dblBudgetSecs) Then strProblem = "the recalc check ran out of time"
    If Len(strProblem) > 0 Then
        CrossSheetFeedback = strProblem
        GoTo Cleanup
    End If

    Set dicReached = CreateObject("Scripting.Dictionary")
    Set colQueue = New Collection
    For Each varKey In dicSheets.Keys
        If varKey <> strCalcUpper And Not dicReached.Exists(varKey) Then
            dicReached(varKey) = True
            colQueue.Add varKey
        End If
    Next varKey

    ' Follow sheet to sheet. Any reached sheet that references wsCalc closes the loop.
    Do While colQueue.Count > 0
        If OverBudget(dblStart, dblBudgetSecs) Then
            CrossSheetFeedback = "the recalc check ran out of time"
            GoTo Cleanup
        End If
        strName = CStr(colQueue(1))
        colQueue.Remove 1
        Set wsNext = FindWorksheet(wbCalc, strName)     ' resolved by ExpandSheetKeys, so it exists

        Set dicKeys = CreateObject("Scripting.Dictionary")
        strProblem = UsedRangeSheetKeys(wsNext, reRef, reText, dicNameSheets, dicKeys)
        Set dicSheets = CreateObject("Scripting.Dictionary")
        If Len(strProblem) = 0 Then strProblem = ExpandSheetKeys(wbCalc, wsNext.Name, dicKeys, dicSheets)
        If Len(strProblem) > 0 Then
            CrossSheetFeedback = strProblem
            GoTo Cleanup
        End If
        If dicSheets.Exists(strCalcUpper) Then
            CrossSheetFeedback = "the solve links through " & wsNext.Name
            GoTo Cleanup
        End If
        For Each varKey In dicSheets.Keys
            If varKey <> strCalcUpper And Not dicReached.Exists(varKey) Then
                dicReached(varKey) = True
                colQueue.Add varKey
            End If
        Next varKey
    Loop

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If Err.Number = 18 Then Err.Raise 18
    LogError "CrossSheetFeedback", Err.Number, Err.Description
    CrossSheetFeedback = "the recalc check hit an unreadable reference"
    Resume Cleanup
End Function

' Purpose: True once dblBudgetSecs have passed since dblStart (a HighResSeconds reading), or if the
'          counter cannot be read.
Private Function OverBudget(ByVal dblStart As Double, ByVal dblBudgetSecs As Double) As Boolean

    Dim dblNow As Double

    On Error GoTo ErrHandler

    dblNow = HighResSeconds()
    OverBudget = (dblNow < 0 Or dblStart < 0 Or dblNow - dblStart > dblBudgetSecs)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If Err.Number = 18 Then Err.Raise 18
    OverBudget = True
    Resume Cleanup
End Function

' Purpose: Adds the keys of dicKeys to the "|KEY|KEY|" entry for strName in dicNameSheets. Returns True if
'          anything was new. A name with no keys gets no entry (keeps the per-formula name search short).
Private Function AddNameSheets(ByVal dicNameSheets As Object, ByVal strName As String, _
                               ByVal dicKeys As Object) As Boolean

    Dim varKey As Variant
    Dim strList As String

    On Error GoTo ErrHandler

    If dicNameSheets.Exists(strName) Then strList = dicNameSheets(strName) Else strList = "|"
    For Each varKey In dicKeys.Keys
        If InStr(1, strList, "|" & varKey & "|", vbBinaryCompare) = 0 Then
            strList = strList & varKey & "|"
            AddNameSheets = True
        End If
    Next varKey
    If AddNameSheets Then dicNameSheets(strName) = strList

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Err.Raise Err.Number, "AddNameSheets", Err.Description             ' the caller decides (full recalc)
End Function

' Purpose: Adds to dicKeys the sheet keys of the formulas in rngSolve (cells on one sheet). The areas are
'          read in ONE formula read of their bounding block, clipped to the used range, with only the areas'
'          own cells scanned - a solve with thousands of precedent areas still costs one read. Returns ""
'          or the reason the block could not be scanned.
Private Function SolveSheetKeys(ByVal rngSolve As Range, ByVal reRef As Object, ByVal reText As Object, _
                                ByVal dicNameSheets As Object, ByVal dicKeys As Object) As String

    Dim wsSolve As Worksheet
    Dim rngUsed As Range
    Dim rngArea As Range
    Dim varForm As Variant
    Dim blnMark() As Boolean
    Dim lngAreaTop() As Long
    Dim lngAreaLeft() As Long
    Dim lngAreaBottom() As Long
    Dim lngAreaRight() As Long
    Dim lngAreas As Long
    Dim lngA As Long
    Dim lngLastRow As Long
    Dim lngLastCol As Long
    Dim lngTop As Long
    Dim lngLeft As Long
    Dim lngBottom As Long
    Dim lngRight As Long
    Dim lngR As Long
    Dim lngC As Long

    On Error GoTo ErrHandler

    Set wsSolve = rngSolve.Worksheet
    Set rngUsed = wsSolve.UsedRange
    lngLastRow = rngUsed.Row + rngUsed.Rows.Count - 1
    lngLastCol = rngUsed.Column + rngUsed.Columns.Count - 1

    ' Each area clipped to the used range (whole-column precedents such as B:B stop there), and the
    ' bounding block of them all.
    lngAreas = rngSolve.Areas.Count
    ReDim lngAreaTop(1 To lngAreas)
    ReDim lngAreaLeft(1 To lngAreas)
    ReDim lngAreaBottom(1 To lngAreas)
    ReDim lngAreaRight(1 To lngAreas)
    lngTop = lngLastRow + 1
    lngLeft = lngLastCol + 1
    lngA = 0
    For Each rngArea In rngSolve.Areas
        lngA = lngA + 1
        lngAreaTop(lngA) = rngArea.Row
        lngAreaLeft(lngA) = rngArea.Column
        lngAreaBottom(lngA) = rngArea.Row + rngArea.Rows.Count - 1
        lngAreaRight(lngA) = rngArea.Column + rngArea.Columns.Count - 1
        If lngAreaBottom(lngA) > lngLastRow Then lngAreaBottom(lngA) = lngLastRow
        If lngAreaRight(lngA) > lngLastCol Then lngAreaRight(lngA) = lngLastCol
        If lngAreaTop(lngA) <= lngAreaBottom(lngA) And lngAreaLeft(lngA) <= lngAreaRight(lngA) Then
            If lngAreaTop(lngA) < lngTop Then lngTop = lngAreaTop(lngA)
            If lngAreaLeft(lngA) < lngLeft Then lngLeft = lngAreaLeft(lngA)
            If lngAreaBottom(lngA) > lngBottom Then lngBottom = lngAreaBottom(lngA)
            If lngAreaRight(lngA) > lngRight Then lngRight = lngAreaRight(lngA)
        End If
    Next rngArea
    If lngBottom = 0 Then GoTo Cleanup      ' nothing inside the used range

    If CDbl(lngBottom - lngTop + 1) * CDbl(lngRight - lngLeft + 1) > m_SCAN_CELL_LIMIT Then
        SolveSheetKeys = "the solve on " & wsSolve.Name & " spans too many cells to check"
        GoTo Cleanup
    End If

    If lngBottom = lngTop And lngRight = lngLeft Then
        ReDim varForm(1 To 1, 1 To 1)
        varForm(1, 1) = wsSolve.Cells(lngTop, lngLeft).Formula
    Else
        varForm = wsSolve.Range(wsSolve.Cells(lngTop, lngLeft), wsSolve.Cells(lngBottom, lngRight)).Formula
    End If

    ' Mark the areas' own cells, then scan only those.
    ReDim blnMark(1 To lngBottom - lngTop + 1, 1 To lngRight - lngLeft + 1)
    For lngA = 1 To lngAreas
        For lngR = lngAreaTop(lngA) To lngAreaBottom(lngA)
            For lngC = lngAreaLeft(lngA) To lngAreaRight(lngA)
                blnMark(lngR - lngTop + 1, lngC - lngLeft + 1) = True
            Next lngC
        Next lngR
    Next lngA

    For lngR = 1 To UBound(varForm, 1)
        For lngC = 1 To UBound(varForm, 2)
            If blnMark(lngR, lngC) Then
                If Left$(CStr(varForm(lngR, lngC)), 1) = "=" Then
                    AddSheetKeys CStr(varForm(lngR, lngC)), reRef, reText, dicNameSheets, dicKeys
                End If
            End If
        Next lngC
    Next lngR

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Err.Raise Err.Number, "SolveSheetKeys", Err.Description            ' the caller decides (full recalc)
End Function

' Purpose: Adds to dicKeys the sheet keys of every formula in wsScan's used range, in one formula read.
'          Returns "" or the reason the sheet could not be scanned.
Private Function UsedRangeSheetKeys(ByVal wsScan As Worksheet, ByVal reRef As Object, ByVal reText As Object, _
                                    ByVal dicNameSheets As Object, ByVal dicKeys As Object) As String

    Dim rngUsed As Range
    Dim varForm As Variant
    Dim lngR As Long
    Dim lngC As Long

    On Error GoTo ErrHandler

    Set rngUsed = wsScan.UsedRange
    If rngUsed.CountLarge > m_SCAN_CELL_LIMIT Then
        UsedRangeSheetKeys = "the solve reaches " & wsScan.Name & ", too large to check"
        GoTo Cleanup
    End If

    If rngUsed.CountLarge = 1 Then
        ReDim varForm(1 To 1, 1 To 1)
        varForm(1, 1) = rngUsed.Formula
    Else
        varForm = rngUsed.Formula
    End If

    ' Constants come back as their text and are skipped by the "=" test.
    For lngR = 1 To UBound(varForm, 1)
        For lngC = 1 To UBound(varForm, 2)
            If Left$(CStr(varForm(lngR, lngC)), 1) = "=" Then
                AddSheetKeys CStr(varForm(lngR, lngC)), reRef, reText, dicNameSheets, dicKeys
            End If
        Next lngC
    Next lngR

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Err.Raise Err.Number, "UsedRangeSheetKeys", Err.Description        ' the caller decides (full recalc)
End Function

' Purpose: Adds to dicOut the sheet keys (UPPER) one formula or name definition references: each sheet
'          reference as written ("SHEET", "FIRST:LAST" for a 3D span, text with [ ] for another
'          workbook), the keys behind any defined name or table it uses (dicNameSheets may be Nothing), and
'          m_KEY_INDIRECT if it calls INDIRECT. String literals are ignored. ExpandSheetKeys resolves them.
Private Sub AddSheetKeys(ByVal strFormula As String, ByVal reRef As Object, ByVal reText As Object, _
                         ByVal dicNameSheets As Object, ByVal dicOut As Object)

    Dim objMatch As Object
    Dim strText As String
    Dim strUpper As String
    Dim strSheet As String
    Dim strBefore As String
    Dim strAfter As String
    Dim varKey As Variant
    Dim varParts As Variant
    Dim lngPos As Long
    Dim lngP As Long

    On Error GoTo ErrHandler

    strText = reText.Replace(strFormula, """""")
    strUpper = UCase$(strText)

    If InStr(1, strUpper, "INDIRECT(", vbBinaryCompare) > 0 Then dicOut(m_KEY_INDIRECT) = True

    For Each objMatch In reRef.Execute(strText)
        If Len(objMatch.SubMatches(0)) > 0 Then
            strSheet = Replace(objMatch.SubMatches(0), "''", "'")
        Else
            strSheet = objMatch.SubMatches(1)
        End If
        If Len(strSheet) > 0 Then dicOut(UCase$(strSheet)) = True
    Next objMatch

    If Not dicNameSheets Is Nothing Then
        For Each varKey In dicNameSheets.Keys
            lngPos = InStr(1, strUpper, varKey, vbBinaryCompare)
            Do While lngPos > 0
                strBefore = vbNullString
                strAfter = vbNullString
                If lngPos > 1 Then strBefore = Mid$(strUpper, lngPos - 1, 1)
                If lngPos + Len(varKey) <= Len(strUpper) Then strAfter = Mid$(strUpper, lngPos + Len(varKey), 1)
                If Not (strBefore Like "[A-Z0-9_.]") And Not (strAfter Like "[A-Z0-9_.]") Then
                    varParts = Split(Mid$(dicNameSheets(varKey), 2), "|")
                    For lngP = LBound(varParts) To UBound(varParts)
                        If Len(varParts(lngP)) > 0 Then dicOut(varParts(lngP)) = True
                    Next lngP
                    Exit Do
                End If
                lngPos = InStr(lngPos + 1, strUpper, varKey, vbBinaryCompare)
            Loop
        Next varKey
    End If

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Err.Raise Err.Number, "AddSheetKeys", Err.Description              ' the caller decides (full recalc)
End Sub

' Purpose: Resolves the sheet keys found on strWhere (see AddSheetKeys) into dicSheets as UPPER worksheet
'          names - a 3D span adds every worksheet from its first to its last sheet. Returns "" or the reason a
'          key cannot be trusted: INDIRECT, another workbook, or a name that is not a worksheet here.
Private Function ExpandSheetKeys(ByVal wbCalc As Workbook, ByVal strWhere As String, _
                                 ByVal dicKeys As Object, ByVal dicSheets As Object) As String

    Dim varKey As Variant
    Dim strKey As String
    Dim wsFirst As Worksheet
    Dim wsLast As Worksheet
    Dim lngColon As Long
    Dim lngFrom As Long
    Dim lngTo As Long
    Dim lngS As Long

    On Error GoTo ErrHandler

    For Each varKey In dicKeys.Keys
        strKey = CStr(varKey)
        If strKey = m_KEY_INDIRECT Then
            ExpandSheetKeys = "INDIRECT on " & strWhere & " cannot be followed"
            GoTo Cleanup
        End If
        If InStr(1, strKey, "[") > 0 Or InStr(1, strKey, "]") > 0 Then
            ExpandSheetKeys = strWhere & " links to another workbook"
            GoTo Cleanup
        End If

        lngColon = InStr(1, strKey, ":")
        If lngColon > 0 Then
            Set wsFirst = FindWorksheet(wbCalc, Left$(strKey, lngColon - 1))
            Set wsLast = FindWorksheet(wbCalc, Mid$(strKey, lngColon + 1))
        Else
            Set wsFirst = FindWorksheet(wbCalc, strKey)
            Set wsLast = wsFirst
        End If
        If wsFirst Is Nothing Or wsLast Is Nothing Then
            ExpandSheetKeys = strWhere & " has a sheet reference that cannot be resolved (" & strKey & ")"
            GoTo Cleanup
        End If

        If wsFirst.Index <= wsLast.Index Then
            lngFrom = wsFirst.Index
            lngTo = wsLast.Index
        Else
            lngFrom = wsLast.Index
            lngTo = wsFirst.Index
        End If
        For lngS = lngFrom To lngTo
            If TypeName(wbCalc.Sheets(lngS)) = "Worksheet" Then dicSheets(UCase$(wbCalc.Sheets(lngS).Name)) = True
        Next lngS
    Next varKey

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Err.Raise Err.Number, "ExpandSheetKeys", Err.Description           ' the caller decides (full recalc)
End Function

' Purpose: The worksheet called strName in wbCalc (not case sensitive), or Nothing.
Private Function FindWorksheet(ByVal wbCalc As Workbook, ByVal strName As String) As Worksheet

    On Error GoTo ErrHandler

    ' Probe: Worksheets() raises when there is no such sheet.
    On Error Resume Next
    Set FindWorksheet = wbCalc.Worksheets(strName)
    On Error GoTo ErrHandler

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Set FindWorksheet = Nothing
    Resume Cleanup
End Function

' Purpose: True if two reads of the same output row match: a single cell (a scalar) or a row read as a
'          1 x n array, compared cell by cell with SameResult.
Private Function SameValues(ByVal varA As Variant, ByVal varB As Variant) As Boolean

    Dim lngC As Long

    On Error GoTo ErrHandler

    If IsArray(varA) And IsArray(varB) Then
        If UBound(varA, 2) <> UBound(varB, 2) Then GoTo Cleanup
        For lngC = 1 To UBound(varA, 2)
            If Not SameResult(varA(1, lngC), varB(1, lngC)) Then GoTo Cleanup
        Next lngC
        SameValues = True
    ElseIf Not IsArray(varA) And Not IsArray(varB) Then
        SameValues = SameResult(varA, varB)
    End If

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If Err.Number = 18 Then Err.Raise 18
    SameValues = False
    Resume Cleanup
End Function

' Purpose: True if two cell values are the same - numbers, text, booleans, or the same error value.
Private Function SameResult(ByVal varA As Variant, ByVal varB As Variant) As Boolean

    On Error GoTo ErrHandler

    If IsError(varA) Or IsError(varB) Then
        ' CStr of an error value gives e.g. "Error 2042", so equal text means the same error.
        If IsError(varA) And IsError(varB) Then SameResult = (CStr(varA) = CStr(varB))
    ElseIf VarType(varA) = VarType(varB) Then
        SameResult = (varA = varB)
    End If

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If Err.Number = 18 Then Err.Raise 18
    SameResult = False
    Resume Cleanup
End Function

' Purpose: Find the input cell of a data table block on WS. A candidate is a cell outside column B
'          with a number 4 rows below it (the first game), a formula or "ANSWER" 3 rows below and
'          1 column right (the formula cell), and EITHER text starting "Example" OR a loaded game
'          number that the lookup cell to its right refers to (e.g. =XLOOKUP($J$2,...)).
'          The block holding rngActive wins; with exactly one candidate it is used; with several and
'          the active cell in none of them, nothing is guessed - strMsg asks for a cell inside a block.
'          With rngZone given, only a block whose INPUT CELL lies inside rngZone is returned (the first
'          one found) - used by Solve Level to pick one level's block on a sheet holding several.
'          blnSearchFailed is set True when the search itself raised, so callers need not parse strMsg.
Private Function FindInputCell(ByVal ws As Worksheet, ByVal rngActive As Range, ByVal strCaller As String, _
                               ByRef strMsg As String, Optional ByVal rngZone As Range = Nothing, _
                               Optional ByRef blnSearchFailed As Boolean) As Range

    Dim rngLast As Range
    Dim rngArea As Range
    Dim rngCand As Range
    Dim rngOnly As Range
    Dim varVals As Variant
    Dim varForm As Variant
    Dim lngRows As Long
    Dim lngCols As Long
    Dim lngR As Long
    Dim lngC As Long
    Dim lngGames As Long
    Dim lngCandidates As Long
    Dim blnInput As Boolean
    Dim strAnswer As String

    On Error GoTo ErrHandler

    blnSearchFailed = False

    ' Bound the scan by real content, not UsedRange - whole-row formatting can stretch UsedRange
    ' to column XFD. Anchored at A1 so array indices are sheet row / column numbers.
    Set rngLast = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If rngLast Is Nothing Then GoTo NotFound
    lngRows = rngLast.Row
    Set rngLast = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    lngCols = rngLast.Column
    If lngRows < m_BODY_OFFSET_ROWS + 1 Or lngCols < 4 Then GoTo NotFound

    ' Two reads for the whole area: values for types, formulas for the structure checks.
    Set rngArea = ws.Range(ws.Cells(1, 1), ws.Cells(lngRows, lngCols))
    varVals = rngArea.Value2
    varForm = rngArea.Formula

    For lngR = 1 To lngRows - m_BODY_OFFSET_ROWS
        For lngC = 3 To lngCols - 1
            If IsNumericCell(varVals(lngR + m_BODY_OFFSET_ROWS, lngC)) Then
                blnInput = False
                If VarType(varVals(lngR, lngC)) = vbString Then
                    ' Same matching as Create Data Table: spaces ignored, so " Example 3" counts.
                    ' Shared example-word list (English / French / Brazilian Portuguese) rather than
                    ' a hardcoded "Example" - see modCaseSetup.m_EXAMPLE_WORDS.
                    blnInput = (modCaseSetup.ExamplePrefixLen(Replace(varVals(lngR, lngC), " ", "")) > 0)
                ElseIf IsNumericCell(varVals(lngR, lngC)) Then
                    ' A game loaded into the calculation - accepted only in a block Create Data Table
                    ' built (Expected / Check labels below the input) whose lookup refers to exactly
                    ' this cell. Ordinary number grids with absolute references must never match.
                    If VarType(varVals(lngR + 1, lngC)) = vbString And VarType(varVals(lngR + 2, lngC)) = vbString Then
                        If varVals(lngR + 1, lngC) = "Expected" And varVals(lngR + 2, lngC) = "Check" Then
                            blnInput = RefersToCell(CStr(varForm(lngR, lngC + 1)), ws.Cells(lngR, lngC).Address(True, True))
                        End If
                    End If
                End If

                If blnInput Then
                    strAnswer = CStr(varForm(lngR + m_FORMULA_OFFSET_ROWS, lngC + 1))
                    If Left$(strAnswer, 1) = "=" Or UCase$(strAnswer) = "ANSWER" Then
                        Set rngCand = ws.Cells(lngR, lngC)
                        If Not rngZone Is Nothing Then
                            If Not Intersect(rngCand, rngZone) Is Nothing Then
                                Set FindInputCell = rngCand
                                Exit Function
                            End If
                            GoTo NextCell
                        End If
                        lngCandidates = lngCandidates + 1
                        Set rngOnly = rngCand
                        lngGames = CountGames(rngCand.Offset(m_BODY_OFFSET_ROWS, 0))
                        ' Inside this block (input row down to the last result; the game and answer
                        ' columns plus any SIDE columns)?
                        If Not Intersect(rngActive, rngCand.Resize(m_BODY_OFFSET_ROWS + lngGames, _
                                                                   2 + SideColumnCount(rngCand))) Is Nothing Then
                            Set FindInputCell = rngCand
                            Exit Function
                        End If
                    End If
                End If
            End If
NextCell:
        Next lngC
    Next lngR

    If Not rngZone Is Nothing Then
        strMsg = strCaller & ": no data table block for this level on " & ws.Name & "."
        Exit Function
    End If

    If lngCandidates = 1 Then
        Set FindInputCell = rngOnly
        Exit Function
    End If
    If lngCandidates > 1 Then
        strMsg = strCaller & ": " & lngCandidates & " data table blocks on " & ws.Name & " - select a cell inside the one to use."
        Exit Function
    End If

NotFound:
    strMsg = strCaller & ": no data table block found on " & ws.Name & " - run Create Data Table first."
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strMsg = strCaller & ": could not search " & ws.Name & " - " & Err.Number & ": " & Err.Description
    blnSearchFailed = True
    Set FindInputCell = Nothing
End Function

' Purpose: Count contiguous numeric cells from rngFirst downwards (the game numbers).
'          The bottom limit comes from the real last used row, not End(xlUp), which stops short
'          over filtered or hidden rows.
Private Function CountGames(ByVal rngFirst As Range) As Long

    ' --- CONSTANTS (local to function) ---
    Const MAX_GAMES As Long = 100000

    Dim ws As Worksheet
    Dim rngLast As Range
    Dim lngLast As Long
    Dim varVals As Variant
    Dim lngI As Long

    On Error GoTo ErrHandler

    Set ws = rngFirst.Parent
    If Not IsNumericCell(rngFirst.Value2) Then Exit Function

    Set rngLast = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If rngLast Is Nothing Then Exit Function
    lngLast = rngLast.Row
    If lngLast <= rngFirst.Row Then
        CountGames = 1
        Exit Function
    End If
    If lngLast - rngFirst.Row + 1 > MAX_GAMES Then lngLast = rngFirst.Row + MAX_GAMES - 1

    varVals = rngFirst.Resize(lngLast - rngFirst.Row + 1, 1).Value2
    For lngI = 1 To UBound(varVals, 1)
        If Not IsNumericCell(varVals(lngI, 1)) Then Exit For
        CountGames = lngI
    Next lngI
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    CountGames = 0
End Function

' Purpose: True for a real number (not blank, not text that looks numeric, not an error).
Private Function IsNumericCell(ByVal varValue As Variant) As Boolean
    On Error GoTo ErrHandler
    IsNumericCell = (VarType(varValue) = vbDouble)
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    IsNumericCell = False
End Function

' Purpose: True if formula text refers to exactly strAddress (e.g. "$J$2"): not $J$20 or $J$200,
'          and not the same address on another sheet ('L01'!$J$2).
Private Function RefersToCell(ByVal strFormula As String, ByVal strAddress As String) As Boolean

    Dim lngPos As Long
    Dim strBefore As String
    Dim strAfter As String

    On Error GoTo ErrHandler

    lngPos = InStr(1, strFormula, strAddress, vbTextCompare)
    Do While lngPos > 0
        strBefore = vbNullString
        strAfter = vbNullString
        If lngPos > 1 Then strBefore = Mid$(strFormula, lngPos - 1, 1)
        If lngPos + Len(strAddress) <= Len(strFormula) Then strAfter = Mid$(strFormula, lngPos + Len(strAddress), 1)
        If Not (strAfter Like "[0-9]") And Not (strBefore Like "[A-Za-z0-9!_'.]") Then
            RefersToCell = True
            Exit Function
        End If
        lngPos = InStr(lngPos + 1, strFormula, strAddress, vbTextCompare)
    Loop
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    RefersToCell = False
End Function

' Purpose: True if any cell in rngCheck holds a What-If TABLE() formula. Public so Create Data Table
'          can refuse to build over an old What-If table.
Public Function HasTableFormulas(ByVal rngCheck As Range) As Boolean

    Dim varForm As Variant
    Dim lngI As Long

    On Error GoTo ErrHandler

    If rngCheck.Cells.Count = 1 Then
        HasTableFormulas = (UCase$(Left$(CStr(rngCheck.Formula), 7)) = "=TABLE(")
        Exit Function
    End If

    varForm = rngCheck.Formula
    For lngI = 1 To UBound(varForm, 1)
        If UCase$(Left$(CStr(varForm(lngI, 1)), 7)) = "=TABLE(" Then
            HasTableFormulas = True
            Exit Function
        End If
    Next lngI
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    HasTableFormulas = False
End Function

' Purpose: What will slow Run Data Table down in this workbook, as short text for the StatusBar:
'          live What-If data tables and volatile functions, by sheet, e.g.
'          "data tables on L01, L02; INDIRECT on Calc". "" if nothing found.
'          Every pass recalculates the whole workbook's dirty and volatile cells, so every sheet counts.
Private Function SlowdownsIn(ByVal wbCase As Workbook) As String

    Dim wsAny As Worksheet
    Dim rngForm As Range
    Dim rngArea As Range
    Dim varForm As Variant
    Dim varNames As Variant
    Dim lngI As Long
    Dim lngR As Long
    Dim lngC As Long
    Dim strF As String
    Dim strTables As String
    Dim strVolatile As String
    Dim strFound As String
    Dim blnTable As Boolean

    On Error GoTo ErrHandler

    varNames = Split(m_VOLATILE_LIST, "|")

    For Each wsAny In wbCase.Worksheets
        Set rngForm = Nothing
        ' Probe: SpecialCells raises when the sheet has no formulas.
        On Error Resume Next
        Set rngForm = wsAny.Cells.SpecialCells(xlCellTypeFormulas)
        On Error GoTo ErrHandler

        If Not rngForm Is Nothing Then
            blnTable = False
            strFound = vbNullString
            For Each rngArea In rngForm.Areas
                If rngArea.Cells.Count = 1 Then
                    ReDim varForm(1 To 1, 1 To 1)
                    varForm(1, 1) = rngArea.Formula
                Else
                    varForm = rngArea.Formula
                End If
                For lngR = 1 To UBound(varForm, 1)
                    For lngC = 1 To UBound(varForm, 2)
                        strF = UCase$(CStr(varForm(lngR, lngC)))
                        If Left$(strF, 7) = "=TABLE(" Then blnTable = True
                        For lngI = LBound(varNames) To UBound(varNames)
                            If InStr(1, strF, varNames(lngI), vbBinaryCompare) > 0 Then
                                If InStr(1, "," & strFound & ",", "," & varNames(lngI) & ",", vbBinaryCompare) = 0 Then
                                    If Len(strFound) > 0 Then strFound = strFound & ","
                                    strFound = strFound & varNames(lngI)
                                End If
                            End If
                        Next lngI
                    Next lngC
                Next lngR
            Next rngArea

            If blnTable Then
                If Len(strTables) > 0 Then strTables = strTables & ", "
                strTables = strTables & wsAny.Name
            End If
            If Len(strFound) > 0 Then
                If Len(strVolatile) > 0 Then strVolatile = strVolatile & "; "
                strVolatile = strVolatile & Replace(strFound, "(", "") & " on " & wsAny.Name
            End If
        End If
    Next wsAny

    If Len(strTables) > 0 Then SlowdownsIn = "data tables on " & strTables
    If Len(strVolatile) > 0 Then
        If Len(SlowdownsIn) > 0 Then SlowdownsIn = SlowdownsIn & "; "
        SlowdownsIn = SlowdownsIn & strVolatile
    End If
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    SlowdownsIn = vbNullString
End Function












