Attribute VB_Name = "modLevelSummary"
Option Explicit

' Purpose: Level summary (GitHub #19; Jaq, 2026-10-23). One table per case: each level's difficulty,
'          its points and the bonus points that relate to it, so a level that is easier than an earlier one
'          (cases do that, and say so) is seen before the solve starts.
'            - Difficulty: the case's own label, column C of the "Level n" row (1,354 of 1,413 levels in the
'              case library, 2026-10-23).
'            - Points: the level's Points column summed over its games.
'            - Bonus Pts: the bonuses whose "Relates to" column names the level (modBonusDock reads them).
'          Full Setup Case writes the table beside the case intro, and below the bonuses on sheet B (long
'          questions make sheet B's columns too wide to see anything placed to their right). The bonus dock
'          lists the same summary as text at the bottom.
'
'          Difficulty is ranked like ski runs: green easy, blue medium, red hard, black very hard and beyond.
'          English, German, French, Spanish, Portuguese and Italian labels are ranked; any other label is shown
'          as the case writes it, uncoloured, and never counts as easier. A ranked level easier than an earlier
'          one is bold, with a down arrow.
'
'          THE LEVEL CELL IS A NUMBER. It reads "Level 3" through its number format only, so nothing in CompBot
'          that looks for "Level n" text (Create Level Sheets, Create Data Table, Solve Level) can take the
'          table for a real level. The table never uses column B, which the bonus dock reads on every sheet.

' --- MODULE CONSTANTS ---
Private Const m_DEBUG_MODE          As Boolean = False
Private Const m_COLS                As Long = 4         ' Level | Difficulty | Points | Bonus Pts
Private Const m_MIN_COL             As Long = 3         ' never column A (spacer) or B (labels the dock reads)
Private Const m_MAX_SHIFT           As Long = 30        ' columns tried to the right when the first spot is taken
Private Const m_MAX_LEVEL_ROWS      As Long = 400       ' the last level's block, when nothing below ends it
Private Const m_MAX_HEADER_ROWS     As Long = 60        ' a level's header row is never further below its heading
Private Const m_MAX_LABEL           As Long = 24        ' longest difficulty label the dock shows in full
Private Const m_MIN_WIDTHS          As String = "11|14|8|10"    ' per table column; a column is widened, never narrowed
Private Const m_BONUS_SHEET         As String = "B"
Private Const m_TITLE               As String = "Dr. Kazza Level Table"     ' named for the requester (Jaq, 2026-10-23)
Private Const m_INK                 As Long = 2500134   ' RGB(38, 38, 38): near-black text, softer than pure black
Private Const m_RANK_VERY_EASY      As Long = 1
Private Const m_RANK_EASY           As Long = 2
Private Const m_RANK_EASY_MEDIUM    As Long = 3
Private Const m_RANK_MEDIUM         As Long = 4
Private Const m_RANK_MID_HARD       As Long = 5
Private Const m_RANK_HARD           As Long = 6
Private Const m_RANK_VERY_HARD      As Long = 7
Private Const m_RANK_EXPERT         As Long = 8
' Columns of the level array ReadLevels fills.
Private Const m_F_LEVEL             As Long = 1
Private Const m_F_LABEL             As Long = 2
Private Const m_F_RANK              As Long = 3
Private Const m_F_POINTS            As Long = 4         ' Empty when the level has no Points column
Private Const m_F_EASIER            As Long = 5         ' the earlier level this one is easier than, or 0
Private Const m_FIELDS              As Long = 5


' Purpose: Full Setup Case's LevelSummary step: write the table beside the intro of wsCase and, when sheet B
'          exists, below its bonuses. Returns False on failure, with the reason in strNote (and blnLogged True
'          when the log has it). On success strNote is "" or says what was skipped on purpose.
Public Function AddLevelSummary(ByVal wbCase As Workbook, ByVal wsCase As Worksheet, _
                                ByRef strNote As String, ByRef blnLogged As Boolean) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "AddLevelSummary"

    Dim varLevels   As Variant
    Dim dicBonus    As Object
    Dim wsBonus     As Worksheet
    Dim dblOther    As Double
    Dim lngCount    As Long
    Dim lngRows     As Long
    Dim lngTop      As Long
    Dim lngLeft     As Long
    Dim blnInit     As Boolean

    On Error GoTo ErrHandler

    strNote = vbNullString
    blnLogged = False
    VBAInit
    blnInit = True

    lngCount = ReadLevels(wsCase, varLevels)
    If lngCount < 0 Then
        strNote = "could not read the levels on " & wsCase.Name
        blnLogged = True                                ' ReadLevels logged the cause
        GoTo Cleanup
    End If
    If lngCount = 0 Then
        strNote = "no levels on " & wsCase.Name
        AddLevelSummary = True
        GoTo Cleanup
    End If

    ' A bonus that cannot be read leaves the Bonus Pts column empty; the table is still worth having.
    Set dicBonus = CreateObject("Scripting.Dictionary")
    If Not modBonusDock.BonusPointsByLevel(wbCase, dicBonus, dblOther) Then
        dicBonus.RemoveAll
        dblOther = 0
        strNote = "bonus points not read (see " & LogLocation() & ")"
    End If
    dblOther = dblOther + OrphanBonus(varLevels, lngCount, dicBonus)

    ' A table from an earlier Setup run is replaced where it stands, never written twice.
    lngRows = TableRows(varLevels, lngCount, dblOther)
    If ReuseSpot(wsCase, lngRows, lngTop, lngLeft) Then
        WriteLevelTable wsCase, lngTop, lngLeft, varLevels, lngCount, dicBonus, dblOther
    ElseIf CaseTableSpot(wsCase, lngRows, lngTop, lngLeft) Then
        WriteLevelTable wsCase, lngTop, lngLeft, varLevels, lngCount, dicBonus, dblOther
    Else
        strNote = Trim$(strNote & " no room beside the intro on " & wsCase.Name)
    End If

    ' Probe: sheet B exists only when Create Bonus Sheet ran.
    On Error Resume Next
    Set wsBonus = wbCase.Worksheets(m_BONUS_SHEET)
    On Error GoTo ErrHandler
    If Not wsBonus Is Nothing Then
        If ReuseSpot(wsBonus, lngRows, lngTop, lngLeft) Then
            WriteLevelTable wsBonus, lngTop, lngLeft, varLevels, lngCount, dicBonus, dblOther
        ElseIf BelowContentSpot(wsBonus, lngRows, lngTop, lngLeft) Then
            WriteLevelTable wsBonus, lngTop, lngLeft, varLevels, lngCount, dicBonus, dblOther
        Else
            strNote = Trim$(strNote & " no free space on sheet " & m_BONUS_SHEET)
        End If
    End If

    AddLevelSummary = True

Cleanup:
    On Error Resume Next
    If blnInit Then VBAFin
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strNote = Err.Number & ": " & Err.Description
    blnLogged = LogError(PROC_NAME, Err.Number, Err.Description)
    AddLevelSummary = False
    Resume Cleanup
End Function


' Purpose: The bonus dock's level summary for wbCase, as plain lines (the task pane shows Text only: HTML
'          renders in a tiny box, Markdown blank). dicBonus / dblOther are the bonus points the dock has
'          already read. Bold is Unicode bold letters, the one bold plain text can carry. Returns "" when the
'          case has no levels. Never raises.
Public Function LevelSummaryText(ByVal wbCase As Workbook, ByVal dicBonus As Object, ByVal dblOther As Double) As String

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "LevelSummaryText"

    Dim wsLevels    As Worksheet
    Dim varLevels   As Variant
    Dim lngCount    As Long
    Dim lngIdx      As Long
    Dim strLevel    As String
    Dim strLine     As String
    Dim strOut      As String
    Dim blnEasier   As Boolean

    On Error GoTo ErrHandler

    Set wsLevels = LevelSourceSheet(wbCase)
    If wsLevels Is Nothing Then GoTo Cleanup
    lngCount = ReadLevels(wsLevels, varLevels)
    If lngCount < 0 Then
        strOut = UCase$(m_TITLE) & vbLf & "Could not read the levels - see " & LogLocation() & "."
        GoTo Cleanup
    End If
    If lngCount = 0 Then GoTo Cleanup
    If Not dicBonus Is Nothing Then dblOther = dblOther + OrphanBonus(varLevels, lngCount, dicBonus)

    strOut = UCase$(m_TITLE) & MidDot() & "sheet " & wsLevels.Name & vbLf
    For lngIdx = 1 To lngCount
        strLevel = "L" & varLevels(lngIdx, m_F_LEVEL) & " " & ShortLabel(varLevels(lngIdx, m_F_LABEL))
        If varLevels(lngIdx, m_F_EASIER) > 0 Then
            blnEasier = True
            strLine = BoldText(strLevel) & " " & ChrW$(&H25BC) & " " & BoldText("easier than L" & varLevels(lngIdx, m_F_EASIER))
        Else
            strLine = strLevel
        End If
        strLine = RankMarker(varLevels(lngIdx, m_F_RANK)) & " " & strLine
        If Not IsEmpty(varLevels(lngIdx, m_F_POINTS)) Then
            strLine = strLine & MidDot() & Format$(varLevels(lngIdx, m_F_POINTS), "0") & " pts"
        End If
        If Not dicBonus Is Nothing Then
            If dicBonus.Exists(CLng(varLevels(lngIdx, m_F_LEVEL))) Then
                strLine = strLine & MidDot() & "+" & Format$(dicBonus(CLng(varLevels(lngIdx, m_F_LEVEL))), "0") & " bonus"
            End If
        End If
        strOut = strOut & strLine & vbLf
    Next lngIdx
    If dblOther > 0 Then strOut = strOut & "Other bonuses: " & Format$(dblOther, "0") & " pts" & vbLf

    strOut = strOut & RankMarker(m_RANK_EASY) & " easy  " & RankMarker(m_RANK_MEDIUM) & " medium  " & _
             RankMarker(m_RANK_HARD) & " hard  " & RankMarker(m_RANK_VERY_HARD) & " very hard  ? unknown tag"
    If blnEasier Then strOut = strOut & vbLf & ChrW$(&H25BC) & " = easier than an earlier level"

Cleanup:
    LevelSummaryText = strOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strOut = UCase$(m_TITLE) & vbLf & "Failed - see " & LogLocation() & "."
    Resume Cleanup
End Function


' Purpose: The sheet holding the case's levels, without asking: Case, Case-Varsity or Task when it has levels,
'          else the visible sheet with the most level headings that is not a level sheet (L01...), sheet B, a
'          backup or an answer key. Nothing when there is none. For the dock and for Setup's landing, which
'          must never pop the "What is the case sheet called?" box.
Public Function LevelSourceSheet(ByVal wbCase As Workbook) As Worksheet

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "LevelSourceSheet"

    Dim varName     As Variant
    Dim wsEach      As Worksheet
    Dim wsBest      As Worksheet
    Dim colMarks    As Collection
    Dim lngBest     As Long

    On Error GoTo ErrHandler

    For Each varName In Array("Case", "Case-Varsity", "Task")
        Set wsEach = Nothing
        ' Probe: the sheet may simply not exist.
        On Error Resume Next
        Set wsEach = wbCase.Worksheets(CStr(varName))
        On Error GoTo ErrHandler
        If Not wsEach Is Nothing Then
            Set colMarks = modCaseSetup.LevelMarkerRows(wsEach)
            If Not colMarks Is Nothing Then
                If colMarks.Count > 0 Then
                    Set LevelSourceSheet = wsEach
                    GoTo Cleanup
                End If
            End If
        End If
    Next varName

    For Each wsEach In wbCase.Worksheets
        If wsEach.Visible = xlSheetVisible _
           And Not modCaseSetup.IsLevelSheetName(wsEach.Name) _
           And StrComp(wsEach.Name, m_BONUS_SHEET, vbTextCompare) <> 0 _
           And Not modBonusDock.IsAnswerKeySheet(wsEach.Name) Then
            Set colMarks = modCaseSetup.LevelMarkerRows(wsEach)
            If Not colMarks Is Nothing Then
                If colMarks.Count > lngBest Then
                    lngBest = colMarks.Count
                    Set wsBest = wsEach
                End If
            End If
        End If
    Next wsEach
    Set LevelSourceSheet = wsBest

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    Set LevelSourceSheet = Nothing
    Resume Cleanup
End Function


' ================================== helpers ==================================

' Purpose: Every level on ws, top to bottom, into varLevels (1 To n, 1 To m_FIELDS). Returns n, 0 when there
'          are no levels, -1 when the read failed (logged).
Private Function ReadLevels(ByVal ws As Worksheet, ByRef varLevels As Variant) As Long

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ReadLevels"

    Dim colMarks    As Collection
    Dim lngCount    As Long
    Dim lngIdx      As Long
    Dim lngRow      As Long
    Dim lngEnd      As Long
    Dim lngBonus    As Long
    Dim lngLastB    As Long
    Dim lngRank     As Long
    Dim lngHardest  As Long
    Dim lngHardLvl  As Long

    On Error GoTo ErrHandler

    ReadLevels = -1
    Set colMarks = modCaseSetup.LevelMarkerRows(ws)
    If colMarks Is Nothing Then GoTo Cleanup
    lngCount = colMarks.Count
    If lngCount = 0 Then
        ReadLevels = 0
        GoTo Cleanup
    End If

    ReDim varLevels(1 To lngCount, 1 To m_FIELDS)
    lngLastB = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For lngIdx = 1 To lngCount
        lngRow = colMarks(lngIdx)
        varLevels(lngIdx, m_F_LEVEL) = MarkerNumber(SafeText(ws.Cells(lngRow, 2).Value2))
        varLevels(lngIdx, m_F_LABEL) = FirstLine(SafeText(ws.Cells(lngRow, 3).Value2))
        lngRank = DifficultyRank(varLevels(lngIdx, m_F_LABEL))
        varLevels(lngIdx, m_F_RANK) = lngRank

        ' The level runs to the row before the next heading; the last one to a Bonus heading below it, the end
        ' of column B, or m_MAX_LEVEL_ROWS, whichever comes first.
        If lngIdx < lngCount Then
            lngEnd = colMarks(lngIdx + 1) - 1
        Else
            lngEnd = lngRow + m_MAX_LEVEL_ROWS
            lngBonus = modCaseSetup.BonusStartRow(ws, lngRow)
            If lngBonus > 0 And lngBonus - 1 < lngEnd Then lngEnd = lngBonus - 1
            If lngLastB < lngEnd Then lngEnd = lngLastB
        End If
        varLevels(lngIdx, m_F_POINTS) = LevelPoints(ws, lngRow, lngEnd)

        ' Easier than an earlier level: ranked lower than the hardest ranked level so far.
        varLevels(lngIdx, m_F_EASIER) = 0
        If lngRank > 0 Then
            If lngRank < lngHardest Then
                varLevels(lngIdx, m_F_EASIER) = lngHardLvl
            ElseIf lngRank > lngHardest Then
                lngHardest = lngRank
                lngHardLvl = varLevels(lngIdx, m_F_LEVEL)
            End If
        End If
    Next lngIdx
    ReadLevels = lngCount

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description & " (sheet " & ws.Name & ")"
    ReadLevels = -1
    Resume Cleanup
End Function


' Purpose: A level's points: every number in its Points column (the first row below the heading with a short
'          Points / Punkte / Puntos / Pontos / Punti cell; found by that cell, not by "Answer", so a German
'          "Antwort" header works too), down to the next section break once games have started. Not only game
'          rows: the 2022 Number 1 Fan scores a "Level Total" row (450 of a section's 750). Example rows score 0.
'          Empty when the level has no Points column or no number in it.
Private Function LevelPoints(ByVal ws As Worksheet, ByVal lngTop As Long, ByVal lngEnd As Long) As Variant

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "LevelPoints"

    Dim varHeader   As Variant
    Dim varBlock    As Variant
    Dim lngHdr      As Long
    Dim lngRow      As Long
    Dim lngPtsCol   As Long
    Dim lngCol      As Long
    Dim lngLastCol  As Long
    Dim lngGames    As Long
    Dim dblSum      As Double
    Dim blnAny      As Boolean
    Dim strGame     As String

    On Error GoTo ErrHandler

    LevelPoints = Empty
    ' The header row: one read of the rows below the heading, scanned in memory.
    lngHdr = Application.WorksheetFunction.Min(lngEnd, lngTop + m_MAX_HEADER_ROWS)
    lngLastCol = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If lngLastCol < 3 Or lngHdr <= lngTop Then GoTo Cleanup
    varHeader = ws.Range(ws.Cells(lngTop, 1), ws.Cells(lngHdr, lngLastCol)).Value2
    lngHdr = 0
    For lngRow = 1 To UBound(varHeader, 1)
        For lngCol = 3 To lngLastCol
            If IsPointsHeader(SafeText(varHeader(lngRow, lngCol))) Then
                lngHdr = lngTop + lngRow - 1
                lngPtsCol = lngCol
                Exit For
            End If
        Next lngCol
        If lngHdr > 0 Then Exit For
    Next lngRow
    If lngHdr = 0 Or lngHdr >= lngEnd Then GoTo Cleanup

    ' Column B to the Points column, one read; at least two rows so the result is always an array.
    varBlock = ws.Range(ws.Cells(lngHdr, 2), ws.Cells(Application.WorksheetFunction.Max(lngEnd, lngHdr + 1), lngPtsCol)).Value2
    For lngRow = 2 To UBound(varBlock, 1)
        strGame = Trim$(SafeText(varBlock(lngRow, 1)))
        If Len(strGame) > 0 Then
            If IsNumeric(strGame) Then
                lngGames = lngGames + 1
            ElseIf lngGames > 0 And modCaseSetup.IsSectionBreak(strGame) Then
                Exit For
            End If
        End If
        If VarType(varBlock(lngRow, lngPtsCol - 1)) = vbDouble Then
            dblSum = dblSum + varBlock(lngRow, lngPtsCol - 1)
            blnAny = True
        End If
    Next lngRow
    If blnAny Then LevelPoints = dblSum

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description & " (sheet " & ws.Name & ", row " & lngTop & ")"
    LevelPoints = Empty
    Resume Cleanup
End Function


' Purpose: Where the table goes beside the intro: level with the intro's first row (the first text or picture
'          above the first level, or above the bonus block when that comes first), one blank column right of the
'          furthest text or picture IN THE ROWS THE TABLE TAKES (a side board lower down, as in Collect Them All,
'          does not push it off screen), never left of column C, moved right until all its cells are empty and
'          unmerged. False if no such spot within m_MAX_SHIFT columns.
Private Function CaseTableSpot(ByVal ws As Worksheet, ByVal lngRows As Long, _
                               ByRef lngTop As Long, ByRef lngLeft As Long) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "CaseTableSpot"

    Dim colMarks    As Collection
    Dim shp         As Shape
    Dim rngLast     As Range
    Dim lngIntroEnd As Long
    Dim lngBonus    As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim lngRight    As Long
    Dim lngBottom   As Long

    On Error GoTo ErrHandler

    Set colMarks = modCaseSetup.LevelMarkerRows(ws)
    If colMarks Is Nothing Then GoTo Cleanup
    If colMarks.Count = 0 Then GoTo Cleanup
    lngIntroEnd = colMarks(1) - 1
    lngBonus = modCaseSetup.BonusStartRow(ws, 0)
    If lngBonus > 0 And lngBonus - 1 < lngIntroEnd Then lngIntroEnd = lngBonus - 1
    If lngIntroEnd < 1 Then lngIntroEnd = 1

    ' The intro's first row: the first with text, or the top of the highest picture.
    lngTop = 0
    For lngRow = 1 To lngIntroEnd
        If Application.WorksheetFunction.CountA(ws.Rows(lngRow)) > 0 Then
            lngTop = lngRow
            Exit For
        End If
    Next lngRow
    For Each shp In ws.Shapes
        If shp.TopLeftCell.Row <= lngIntroEnd Then
            If lngTop = 0 Or shp.TopLeftCell.Row < lngTop Then lngTop = shp.TopLeftCell.Row
        End If
    Next shp
    If lngTop = 0 Then lngTop = 2
    ' No intro above the first level: no room, rather than a table on Level 1's own heading row (where its
    ' "Points" header would be read as the level's own).
    If lngTop >= colMarks(1) Then GoTo Cleanup

    ' The furthest text or picture in the rows the table (and its key row) will take.
    lngBottom = lngTop + lngRows
    For lngRow = lngTop To lngBottom
        Set rngLast = ws.Cells(lngRow, ws.Columns.Count).End(xlToLeft)
        If Not IsEmpty(rngLast.Value2) Or rngLast.MergeCells Then
            lngCol = rngLast.Column + rngLast.MergeArea.Columns.Count - 1
            If lngCol > lngRight Then lngRight = lngCol
        End If
    Next lngRow
    For Each shp In ws.Shapes
        If shp.TopLeftCell.Row <= lngBottom And shp.BottomRightCell.Row >= lngTop Then
            If shp.BottomRightCell.Column > lngRight Then lngRight = shp.BottomRightCell.Column
        End If
    Next shp

    ' One blank column beside the intro; two when a short intro lets the table reach Level 1's rows, so the
    ' level's own working (Save From Example reads up to two blank columns past it) never runs into it.
    If lngBottom >= colMarks(1) Then
        lngLeft = Application.WorksheetFunction.Max(lngRight + 3, m_MIN_COL)
    Else
        lngLeft = Application.WorksheetFunction.Max(lngRight + 2, m_MIN_COL)
    End If
    CaseTableSpot = FreeSpot(ws, lngTop, lngLeft, lngRows)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Raised, not read as "no room": AddLevelSummary logs it and the step reports a failure.
    Err.Raise Err.Number, PROC_NAME, Err.Description & " (sheet " & ws.Name & ")"
End Function


' Purpose: A table from an earlier Setup run (its header row: Level | Difficulty | Points | Bonus Pts, never in
'          column B) is cleared, and True returned with its top-left when the new table fits in the same spot.
'          False when there was none, or the new one does not fit there (it then goes where a first run would).
Private Function ReuseSpot(ByVal ws As Worksheet, ByVal lngRows As Long, _
                           ByRef lngTop As Long, ByRef lngLeft As Long) As Boolean

    Dim varUsed     As Variant
    Dim rngUsed     As Range
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim lngEnd      As Long
    Dim lngOldLeft  As Long
    Dim strKey      As String

    On Error GoTo ErrHandler

    Set rngUsed = ws.UsedRange
    If rngUsed.Rows.Count < 2 Or rngUsed.Columns.Count < m_COLS Then GoTo Cleanup
    varUsed = rngUsed.Value2
    For lngRow = 1 To UBound(varUsed, 1)
        For lngCol = 1 To UBound(varUsed, 2) - m_COLS + 1
            If SafeText(varUsed(lngRow, lngCol)) = "Level" Then
                If SafeText(varUsed(lngRow, lngCol + 1)) = "Difficulty" And SafeText(varUsed(lngRow, lngCol + 2)) = "Points" _
                   And SafeText(varUsed(lngRow, lngCol + 3)) = "Bonus Pts" And rngUsed.Column + lngCol - 1 > 2 Then
                    lngTop = rngUsed.Row + lngRow - 1
                    lngLeft = rngUsed.Column + lngCol - 1
                    GoTo Found
                End If
            End If
        Next lngCol
    Next lngRow
    GoTo Cleanup

Found:
    ' The old table runs down while its Level column is filled; a key row ("down arrow ...") may follow.
    lngEnd = lngTop
    Do While Len(SafeText(ws.Cells(lngEnd + 1, lngLeft).Value2)) > 0
        lngEnd = lngEnd + 1
        strKey = SafeText(ws.Cells(lngEnd, lngLeft).Value2)
        If Left$(strKey, 1) = ChrW$(&H25BC) Then Exit Do
    Loop
    ' The title sits in the row above the header; the table's spot starts there.
    If lngTop > 1 Then
        If SafeText(ws.Cells(lngTop - 1, lngLeft).Value2) = m_TITLE Then lngTop = lngTop - 1
    End If
    With ws.Range(ws.Cells(lngTop, lngLeft), ws.Cells(lngEnd, lngLeft + m_COLS - 1))
        .ClearContents
        .ClearFormats
    End With
    ' FreeSpot may move right; only the old spot itself counts as reuse.
    lngOldLeft = lngLeft
    ReuseSpot = FreeSpot(ws, lngTop, lngLeft, lngRows)
    If ReuseSpot Then ReuseSpot = (lngLeft = lngOldLeft)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Err.Raise Err.Number, "ReuseSpot", Err.Description
End Function


' Purpose: Bonus points filed under a level the case does not have ("relates to 7" in a 6-level case). They go
'          to the Other row, so Bonus Pts always adds up to every bonus. Removed from dicBonus.
Private Function OrphanBonus(ByVal varLevels As Variant, ByVal lngCount As Long, ByVal dicBonus As Object) As Double

    Dim varKey  As Variant
    Dim lngIdx  As Long
    Dim blnSeen As Boolean

    For Each varKey In dicBonus.Keys
        blnSeen = False
        For lngIdx = 1 To lngCount
            If CLng(varLevels(lngIdx, m_F_LEVEL)) = CLng(varKey) Then
                blnSeen = True
                Exit For
            End If
        Next lngIdx
        If Not blnSeen Then
            OrphanBonus = OrphanBonus + dicBonus(varKey)
            dicBonus.Remove varKey
        End If
    Next varKey
End Function


' Purpose: Where the table goes on sheet B: two blank rows below the last content or picture, from column C.
Private Function BelowContentSpot(ByVal ws As Worksheet, ByVal lngRows As Long, _
                                  ByRef lngTop As Long, ByRef lngLeft As Long) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "BelowContentSpot"

    Dim shp         As Shape
    Dim lngLast     As Long

    On Error GoTo ErrHandler

    ' UsedRange can over-report at the bottom; that only moves the table further down.
    lngLast = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    For Each shp In ws.Shapes
        If shp.BottomRightCell.Row > lngLast Then lngLast = shp.BottomRightCell.Row
    Next shp

    lngTop = lngLast + 3
    lngLeft = m_MIN_COL
    BelowContentSpot = FreeSpot(ws, lngTop, lngLeft, lngRows)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Raised, not read as "no room": AddLevelSummary logs it and the step reports a failure.
    Err.Raise Err.Number, PROC_NAME, Err.Description & " (sheet " & ws.Name & ")"
End Function


' Purpose: Move lngLeft right until lngRows x m_COLS cells from (lngTop, lngLeft) are empty and unmerged.
'          False if none within m_MAX_SHIFT columns.
Private Function FreeSpot(ByVal ws As Worksheet, ByVal lngTop As Long, ByRef lngLeft As Long, _
                          ByVal lngRows As Long) As Boolean

    Dim rngSpot     As Range
    Dim varMerged   As Variant
    Dim lngShift    As Long

    On Error GoTo ErrHandler

    For lngShift = 0 To m_MAX_SHIFT
        Set rngSpot = ws.Cells(lngTop, lngLeft).Resize(lngRows, m_COLS)
        varMerged = rngSpot.MergeCells                  ' Null when some cells are merged
        If Application.WorksheetFunction.CountA(rngSpot) = 0 And Not IsNull(varMerged) Then
            If Not varMerged Then
                FreeSpot = True
                GoTo Cleanup
            End If
        End If
        lngLeft = lngLeft + 1
    Next lngShift

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Raised, not read as "no room": AddLevelSummary logs it and the step reports a failure.
    Err.Raise Err.Number, "FreeSpot", Err.Description & " (sheet " & ws.Name & ")"
End Function


' Purpose: Rows the table takes: title, header, one per level, an "Other bonus" row when some bonus relates to no
'          level, and a key row when a level is easier than an earlier one.
Private Function TableRows(ByVal varLevels As Variant, ByVal lngCount As Long, ByVal dblOther As Double) As Long

    Dim lngIdx As Long

    TableRows = 2 + lngCount
    If dblOther > 0 Then TableRows = TableRows + 1
    For lngIdx = 1 To lngCount
        If varLevels(lngIdx, m_F_EASIER) > 0 Then
            TableRows = TableRows + 1
            Exit For
        End If
    Next lngIdx
End Function


' Purpose: Write the title at (lngTop, lngLeft) and the table below it in one array assignment. Look (Jaq,
'          2026-10-23): clean and neutral - white rows, a light grey header, thin grey lines, dark text - with
'          colour ONLY in the Difficulty cell, a soft pastel per ski step, so it reads at a glance without
'          glaring. A level easier than an earlier one is bold with a down arrow. A table column narrower than
'          m_MIN_WIDTHS is widened (never narrowed: the columns may hold the case's own data further down, and
'          wider never hides anything); a label still too long shrinks to fit.
Private Sub WriteLevelTable(ByVal ws As Worksheet, ByVal lngTop As Long, ByVal lngLeft As Long, _
                            ByVal varLevels As Variant, ByVal lngCount As Long, _
                            ByVal dicBonus As Object, ByVal dblOther As Double)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "WriteLevelTable"

    Dim varOut      As Variant
    Dim rngTable    As Range
    Dim rngRow      As Range
    Dim lngRows     As Long
    Dim lngIdx      As Long
    Dim lngLevel    As Long
    Dim lngRank     As Long
    Dim blnEasier   As Boolean
    Dim strFmt      As String
    Dim strFmtDown  As String
    Dim varWidths   As Variant

    On Error GoTo ErrHandler

    lngRows = 1 + lngCount
    If dblOther > 0 Then lngRows = lngRows + 1
    ReDim varOut(1 To lngRows, 1 To m_COLS)
    varOut(1, 1) = "Level"
    varOut(1, 2) = "Difficulty"
    varOut(1, 3) = "Points"
    varOut(1, 4) = "Bonus Pts"
    For lngIdx = 1 To lngCount
        lngLevel = varLevels(lngIdx, m_F_LEVEL)
        varOut(lngIdx + 1, 1) = lngLevel
        varOut(lngIdx + 1, 2) = varLevels(lngIdx, m_F_LABEL)
        varOut(lngIdx + 1, 3) = varLevels(lngIdx, m_F_POINTS)
        If dicBonus.Exists(lngLevel) Then varOut(lngIdx + 1, 4) = dicBonus(lngLevel)
    Next lngIdx
    If dblOther > 0 Then
        varOut(lngRows, 1) = "Other bonus"
        varOut(lngRows, 4) = dblOther
    End If

    With ws.Cells(lngTop, lngLeft)
        .Value2 = m_TITLE
        .Font.Bold = True
        .Font.Color = m_INK
    End With

    Set rngTable = ws.Cells(lngTop + 1, lngLeft).Resize(lngRows, m_COLS)
    ' Labels are text, whatever they look like ("1/2" must not become a date).
    rngTable.Columns(2).NumberFormat = "@"
    rngTable.Value2 = varOut

    ' Number formats with quotes: "Level "0 and the same with a down arrow (ChrW, so the source stays ASCII).
    strFmt = """Level ""0"
    strFmtDown = """Level ""0"" " & ChrW$(&H25BC) & """"
    rngTable.Interior.Color = RGB(255, 255, 255)
    rngTable.Font.Color = m_INK
    rngTable.Borders.LineStyle = xlContinuous
    rngTable.Borders.Weight = xlThin
    rngTable.Borders.Color = RGB(208, 208, 208)
    varWidths = Split(m_MIN_WIDTHS, "|")
    For lngIdx = 1 To m_COLS
        If rngTable.Columns(lngIdx).ColumnWidth < CDbl(varWidths(lngIdx - 1)) Then
            rngTable.Columns(lngIdx).ColumnWidth = CDbl(varWidths(lngIdx - 1))
        End If
    Next lngIdx
    rngTable.Columns(2).ShrinkToFit = True
    rngTable.Columns(1).HorizontalAlignment = xlLeft
    With rngTable.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
        .Borders(xlEdgeBottom).Color = RGB(128, 128, 128)
    End With

    For lngIdx = 1 To lngCount
        Set rngRow = rngTable.Rows(lngIdx + 1)
        lngRank = varLevels(lngIdx, m_F_RANK)
        If lngRank > 0 Then rngRow.Cells(1, 2).Interior.Color = RankFill(lngRank)
        If varLevels(lngIdx, m_F_EASIER) > 0 Then
            blnEasier = True
            rngRow.Font.Bold = True
            rngRow.Cells(1, 1).NumberFormat = strFmtDown
        Else
            rngRow.Cells(1, 1).NumberFormat = strFmt
        End If
    Next lngIdx

    If blnEasier Then
        With ws.Cells(lngTop + 1 + rngTable.Rows.Count, lngLeft)
            .Value2 = ChrW$(&H25BC) & " bold: easier than an earlier level"
            .Font.Italic = True
            .Font.Color = m_INK
        End With
    End If

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Raised on: Setup's step reports it, and the half-written table stays visible as the failure flag.
    Err.Raise Err.Number, PROC_NAME, Err.Description
End Sub


' Purpose: 0 for an unknown label; else 1 very easy .. 8 expert. Accents, case, spaces, dashes and slashes
'          ignored, so "VeryHard", "Mid-Hard", "Medium/Hard", "Sehr Schwer", "Tres difficile" and "Muito facil"
'          all rank. Every tag in the case library was checked (tools\label_inventory.py, 2026-10-23):
'            - between two steps: "Easy-Medium", "Mid-Easy", "Moderately Easy" -> easy-medium;
'              "Medium-Hard", "Medium/Hard", "Moderately Hard" -> medium-hard;
'            - unknown, never guessed: "Not too easy", "Easy/Hard" (easy AND hard), "40 points", jokes, blank.
'          Hard is tested before easy, and easy before expert, so "Extremely hard" / "Extremely easy" are not
'          "expert".
Private Function DifficultyRank(ByVal strLabel As String) As Long

    ' --- CONSTANTS (local to function) ---
    Const EASY_WORDS As String = "EASY|EINFACH|LEICHT|FACIL|SIMPLE|BEGINNER|BASIC"
    Const HARD_WORDS As String = "HARD|DIFFICULT|SCHWER|DIFICIL|DIFFICILE"

    Dim strKey  As String
    Dim blnVery As Boolean
    Dim blnEasy As Boolean
    Dim blnHard As Boolean

    strKey = NormaliseLabel(strLabel)
    If Len(strKey) = 0 Then Exit Function
    ' "Not too easy" (2026 Claim Treasure) says what a level is not: left unranked rather than guessed.
    If Left$(strKey, 3) = "NOT" Then Exit Function
    blnEasy = ContainsAny(strKey, EASY_WORDS)
    blnHard = ContainsAny(strKey, HARD_WORDS)
    ' "Easy/Hard" (2026 UK) contradicts itself: unknown.
    If blnEasy And blnHard Then Exit Function

    blnVery = StartsWithAny(strKey, "VERY|SEHR|MUY|MUITO|TRES|MOLTO|EXTREMELY|SUPER|ULTRA|REALLY")

    If ContainsAny(strKey, "MIDHARD|MEDIUMHARD|MEDHARD|MODERATELYHARD|MODERATELYDIFFICULT|MITTELSCHWER") Then
        DifficultyRank = m_RANK_MID_HARD
    ElseIf ContainsAny(strKey, "EASYMEDIUM|MEDIUMEASY|MIDEASY|EASYMED|EASYMID|MODERATELYEASY|EASYMODERATE") Then
        DifficultyRank = m_RANK_EASY_MEDIUM
    ElseIf blnHard Then
        If blnVery Then DifficultyRank = m_RANK_VERY_HARD Else DifficultyRank = m_RANK_HARD
    ElseIf blnEasy Then
        If blnVery Then DifficultyRank = m_RANK_VERY_EASY Else DifficultyRank = m_RANK_EASY
    ElseIf ContainsAny(strKey, "MEDI|MODERATE|MITTEL|MOYEN|INTERMEDI|NORMAL|AVERAGE|MID") Then
        DifficultyRank = m_RANK_MEDIUM
    ElseIf ContainsAny(strKey, "EXPERT|EXTREME|INSANE|NIGHTMARE|IMPOSSIBLE|BRUTAL") Then
        DifficultyRank = m_RANK_EXPERT
    End If
End Function


' Purpose: Upper case, accents folded to plain letters, spaces / dashes / underscores / dots / slashes dropped.
'          Accented letters are built with ChrW so the source stays ASCII on any code page.
Private Function NormaliseLabel(ByVal strLabel As String) As String

    Dim strKey  As String
    Dim varPair As Variant
    Dim varCode As Variant

    strKey = UCase$(Trim$(Replace(strLabel, Chr$(160), " ")))
    For Each varPair In Array(Array("A", &HC0, &HC1, &HC2, &HC3, &HC4), Array("E", &HC8, &HC9, &HCA, &HCB), _
                              Array("I", &HCC, &HCD, &HCE, &HCF), Array("O", &HD2, &HD3, &HD4, &HD5, &HD6), _
                              Array("U", &HD9, &HDA, &HDB, &HDC), Array("C", &HC7))
        For Each varCode In varPair
            If Not VarType(varCode) = vbString Then strKey = Replace(strKey, ChrW$(varCode), varPair(0))
        Next varCode
    Next varPair
    strKey = Replace(Replace(Replace(Replace(Replace(strKey, " ", ""), "-", ""), "_", ""), ".", ""), "/", "")
    NormaliseLabel = strKey
End Function


Private Function ContainsAny(ByVal strText As String, ByVal strWords As String) As Boolean
    Dim varWord As Variant
    For Each varWord In Split(strWords, "|")
        If InStr(1, strText, varWord, vbBinaryCompare) > 0 Then
            ContainsAny = True
            Exit Function
        End If
    Next varWord
End Function


Private Function StartsWithAny(ByVal strText As String, ByVal strWords As String) As Boolean
    Dim varWord As Variant
    For Each varWord In Split(strWords, "|")
        If Left$(strText, Len(varWord)) = varWord Then
            StartsWithAny = True
            Exit Function
        End If
    Next varWord
End Function


' Purpose: The Difficulty cell's fill: ski steps as soft pastels (green, blue, red, grey for "black"), all light
'          enough for dark text, so nothing glares.
Private Function RankFill(ByVal lngRank As Long) As Long
    Select Case lngRank
        Case m_RANK_VERY_EASY:   RankFill = RGB(232, 244, 226)     ' palest green
        Case m_RANK_EASY:        RankFill = RGB(214, 236, 205)     ' soft green
        Case m_RANK_EASY_MEDIUM: RankFill = RGB(213, 234, 232)     ' between the green and the blue
        Case m_RANK_MEDIUM:      RankFill = RGB(214, 226, 243)     ' soft blue
        Case m_RANK_MID_HARD:    RankFill = RGB(246, 226, 214)     ' soft peach
        Case m_RANK_HARD:        RankFill = RGB(243, 212, 212)     ' soft red
        Case m_RANK_VERY_HARD:   RankFill = RGB(214, 214, 214)     ' soft grey: the "black run", without glare
        Case Else:               RankFill = RGB(190, 190, 190)     ' expert: a shade darker, text still dark
    End Select
End Function


' Purpose: The dock's ski marker (the pane shows no colour): circle easy, square medium, diamond hard, two
'          diamonds very hard, three expert, half-filled circle / half-filled diamond between two steps, and a
'          question mark for a tag CompBot does not know (so an unknown tag is never mistaken for a ranked one).
Private Function RankMarker(ByVal lngRank As Long) As String
    Select Case lngRank
        Case m_RANK_VERY_EASY, m_RANK_EASY:     RankMarker = ChrW$(&H25CF)
        Case m_RANK_EASY_MEDIUM:                RankMarker = ChrW$(&H25D1)
        Case m_RANK_MEDIUM:                     RankMarker = ChrW$(&H25A0)
        Case m_RANK_MID_HARD:                   RankMarker = ChrW$(&H25C8)
        Case m_RANK_HARD:                       RankMarker = ChrW$(&H25C6)
        Case m_RANK_VERY_HARD:                  RankMarker = ChrW$(&H25C6) & ChrW$(&H25C6)
        Case m_RANK_EXPERT:                     RankMarker = ChrW$(&H25C6) & ChrW$(&H25C6) & ChrW$(&H25C6)
        Case Else:                              RankMarker = "?"
    End Select
End Function


' Purpose: ASCII letters and digits as Unicode sans-serif bold (the pane's plain text has no other bold).
'          Anything else is kept as it is.
Private Function BoldText(ByVal strText As String) As String

    Dim lngIdx  As Long
    Dim lngCode As Long
    Dim strOut  As String

    For lngIdx = 1 To Len(strText)
        lngCode = AscW(Mid$(strText, lngIdx, 1))
        If lngCode >= 65 And lngCode <= 90 Then
            strOut = strOut & Astral(&H1D5D4 + lngCode - 65)
        ElseIf lngCode >= 97 And lngCode <= 122 Then
            strOut = strOut & Astral(&H1D5EE + lngCode - 97)
        ElseIf lngCode >= 48 And lngCode <= 57 Then
            strOut = strOut & Astral(&H1D7EC + lngCode - 48)
        Else
            strOut = strOut & Mid$(strText, lngIdx, 1)
        End If
    Next lngIdx
    BoldText = strOut
End Function


' Purpose: A code point above U+FFFF as its UTF-16 surrogate pair.
Private Function Astral(ByVal lngCode As Long) As String
    Astral = ChrW$(&HD800& + ((lngCode - &H10000) \ &H400&)) & ChrW$(&HDC00& + ((lngCode - &H10000) Mod &H400&))
End Function


' Purpose: "Level 3", "SECTION 3:" -> 3. The heading has already passed modCaseSetup.IsLevelMarker.
Private Function MarkerNumber(ByVal strMarker As String) As Long
    Dim strNum As String
    strNum = UCase$(Replace(strMarker, Chr$(160), " "))
    strNum = Trim$(Replace(Replace(Replace(strNum, "LEVEL", ""), "SECTION", ""), ":", ""))
    If IsNumeric(strNum) Then MarkerNumber = CLng(strNum)
End Function


Private Function FirstLine(ByVal strText As String) As String
    strText = Replace(strText, vbCr, vbLf)
    If InStr(1, strText, vbLf) > 0 Then strText = Left$(strText, InStr(1, strText, vbLf) - 1)
    FirstLine = Trim$(strText)
End Function


' Purpose: The label for the dock: "no label" when blank, cut with an ellipsis past m_MAX_LABEL characters.
Private Function ShortLabel(ByVal strLabel As String) As String
    If Len(strLabel) = 0 Then
        ShortLabel = "no label"
    ElseIf Len(strLabel) > m_MAX_LABEL Then
        ShortLabel = Left$(strLabel, m_MAX_LABEL - 1) & ChrW$(&H2026)
    Else
        ShortLabel = strLabel
    End If
End Function


' Purpose: A Points header in the languages of the library: Points, Pts, Punkte, Puntos, Pontos, Punti.
'          Short cells only, so a sentence that starts "Points are..." is never taken for the header.
Private Function IsPointsHeader(ByVal strText As String) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const MAX_HEADER_LEN As Long = 25

    Dim strKey As String
    strKey = UCase$(Trim$(Replace(strText, Chr$(160), " ")))
    If Len(strKey) > MAX_HEADER_LEN Then Exit Function
    IsPointsHeader = (Left$(strKey, 5) = "POINT" Or Left$(strKey, 5) = "PUNKT" Or Left$(strKey, 5) = "PUNTO" _
                      Or Left$(strKey, 5) = "PONTO" Or Left$(strKey, 5) = "PUNTI" Or strKey = "PTS")
End Function


Private Function MidDot() As String
    MidDot = " " & ChrW$(&HB7) & " "
End Function


' Purpose: Any cell value as text; errors become "".
Private Function SafeText(ByVal varValue As Variant) As String
    If IsError(varValue) Or IsNull(varValue) Or IsObject(varValue) Then
        SafeText = vbNullString
    Else
        SafeText = CStr(varValue)
    End If
End Function

