Attribute VB_Name = "modBonusDock"
Option Explicit

' Purpose: Bonus questions dock. Lists the active case's open bonus questions in OA Robot's output task pane
'          (docked on the right), drops them as they are answered, lets a bonus be cleared by hand, and steps
'          through them on the StatusBar as a fallback. Refresh = run Show Bonus Dock again (or chain it with
'          CommandAfter behind any command that saves a bonus answer).
'
' Commands (BonusDock collection):
'   Show Bonus Dock                 BonusDockText()            -> Excel Task Pane Output (Text)
'   Clear Bonus From Dock           ClearBonusFromDock(bonus)  -> then Show Bonus Dock
'   Restore Cleared Bonuses         RestoreClearedBonuses()    -> then Show Bonus Dock
'   Next Bonus In Status Bar        StepBonusStatus(1)
'   Previous Bonus In Status Bar    StepBonusStatus(-1)
'   Clear Bonus Status Bar          ClearBonusStatusBar()
'
' No Application events and no Application.OnTime, by design: an event-driven auto-refresh was built and
' parked (bonusdock\parked-autorefresh) after Excel crashed repeatedly during testing on 2026-09-16 - one of
' those crashes an OA Robot .NET exception while it was switched on. Competition tooling must not risk that.
'
' Layout rule (measured 2026-09-16 over the 152-case library, 222 bonus blocks): the bonus labels
' ("Bonus 1", "Bonus A") sit in column B (219 of 222), under a header row carrying "Answer" (E),
' "Points" (D), a relates-to column (C) and "Question" (G, sometimes H). Answer sheets are skipped.
'
' State kept in the case workbook: one hidden workbook-level name, BonusDock_Cleared, holding the keys
' cleared by hand ("|2|B|"). Nothing else is written to the case file.

' --- MODULE CONSTANTS ---
Private Const m_DEBUG_MODE          As Boolean = False
Private Const m_CLEARED_NAME        As String = "BonusDock_Cleared"
Private Const m_DOCK_COMMAND        As String = "Show Bonus Dock"
Private Const m_LABEL_COL           As Long = 2
Private Const m_MAX_COLS            As Long = 60
Private Const m_HEADER_LOOKBACK     As Long = 12
Private Const m_MAX_LINK_HOPS       As Long = 10
Private Const m_MAX_STATUS          As Long = 255
Private Const m_DEFAULT_ANSWER_COL  As Long = 5
Private Const m_DEFAULT_POINTS_COL  As Long = 4
Private Const m_RELATES_COL         As Long = 3
Private Const m_STATE_OPEN          As Long = 0
Private Const m_STATE_ANSWERED      As Long = 1
Private Const m_STATE_ERROR         As Long = 2
Private Const m_RESULT_FAILED       As Long = -1
Private Const m_WRAP_WIDTH          As Long = 55    ' OA Robot's Text pane does not wrap - lines are broken here
Private Const m_LABEL_PATTERN       As String = "*bonus*"
Private Const m_MAX_MATCH_HOPS      As Long = 50    ' more "bonus" mentions than this: one array read instead

' --- MODULE VARIABLES ---
Private m_objLinkRegex        As Object
Private m_lngStatusPos        As Long
Private m_strStatusBook       As String

' --- TYPES ---
Private Type tBonus
    strKey      As String
    strQuestion As String
    varPoints   As Variant
    strRelates  As String
    rngAnswer   As Range
    lngState    As Long
    blnCleared  As Boolean
End Type

' Handler pattern: every procedure that touches the object model uses On Error GoTo ErrHandler / Cleanup.
' The pure string helpers (BonusKey, NormaliseKey, IsAnswerKeySheet, BonusCaption, RelatesText, LongestText,
' KeyListText, WrapText, MidDot, SafeText) are deliberately exempt: they only run under a caller's handler.


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Show Bonus Dock
' Macro Expression:       modBonusDock.BonusDockText()
'----------------------------------------------------------------------------------------------------
' Purpose: The dock's text for the active workbook: open bonuses first (points, relates-to, question),
'          then a footer of answered and cleared ones. Never raises; on failure returns a short
'          message pointing at the log, so the pane still says something useful.
Public Function BonusDockText() As String

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "BonusDockText"

    Dim wbCase       As Workbook
    Dim udtBonuses() As tBonus
    Dim lngCount     As Long
    Dim lngIdx       As Long
    Dim lngOpen      As Long
    Dim dblOpenPts   As Double
    Dim strSheet     As String
    Dim strCards     As String
    Dim strAnswered  As String
    Dim strCleared   As String
    Dim strRule      As String
    Dim strOut       As String
    Dim lngErrNumber As Long
    Dim strErrText   As String

    On Error GoTo ErrHandler

    strRule = String$(30, ChrW$(&H2501))
    Set wbCase = ActiveWorkbook
    If wbCase Is Nothing Then
        strOut = "No workbook is active. Open a case file and run " & m_DOCK_COMMAND & " again."
        GoTo Cleanup
    End If

    lngCount = CollectBonuses(wbCase, udtBonuses, strSheet)
    If lngCount = m_RESULT_FAILED Then
        strOut = "BONUS QUESTIONS" & vbLf & WrapText(wbCase.Name) & vbLf & strRule & vbLf & _
                 WrapText("Could not read the bonus questions - see " & LogLocation() & ".")
        GoTo Cleanup
    End If
    If lngCount = 0 Then
        strOut = "BONUS QUESTIONS" & vbLf & WrapText(wbCase.Name) & vbLf & strRule & vbLf & _
                 WrapText("No bonus questions found. The dock looks for 'Bonus 1' / 'Bonus A' labels in " & _
                          "column B of every sheet except answer sheets.")
        GoTo Cleanup
    End If

    For lngIdx = 1 To lngCount
        If udtBonuses(lngIdx).blnCleared Then
            strCleared = strCleared & ", " & udtBonuses(lngIdx).strKey
        ElseIf udtBonuses(lngIdx).lngState = m_STATE_ANSWERED Then
            strAnswered = strAnswered & ", " & udtBonuses(lngIdx).strKey
        Else
            lngOpen = lngOpen + 1
            If IsNumeric(udtBonuses(lngIdx).varPoints) And Not IsEmpty(udtBonuses(lngIdx).varPoints) Then
                dblOpenPts = dblOpenPts + CDbl(udtBonuses(lngIdx).varPoints)
            End If
            strCards = strCards & WrapText(ChrW$(&H25B6) & " " & BonusCaption(udtBonuses(lngIdx))) & vbLf
            If udtBonuses(lngIdx).lngState = m_STATE_ERROR Then
                strCards = strCards & WrapText(ChrW$(&H26A0) & " the answer cell (" & _
                           udtBonuses(lngIdx).rngAnswer.Address(False, False) & ") shows an error") & vbLf
            End If
            strCards = strCards & WrapText(udtBonuses(lngIdx).strQuestion) & vbLf & vbLf
        End If
    Next lngIdx

    strOut = "BONUS QUESTIONS" & vbLf & _
             WrapText(wbCase.Name & " - sheet " & strSheet) & vbLf & _
             WrapText(lngOpen & " open of " & lngCount & MidDot() & Format$(dblOpenPts, "0") & " pts still to get" & _
                      MidDot() & "updated " & Format$(Now, "hh:nn:ss")) & vbLf & _
             strRule & vbLf & vbLf

    If lngOpen = 0 Then
        strOut = strOut & ChrW$(&H2714) & " Every bonus is answered or cleared." & vbLf & vbLf
    Else
        strOut = strOut & strCards
    End If

    strOut = strOut & strRule & vbLf & _
             WrapText(ChrW$(&H2714) & " Answered: " & KeyListText(strAnswered)) & vbLf & _
             WrapText(ChrW$(&H2716) & " Cleared: " & KeyListText(strCleared))
    If Len(strCleared) > 0 Then strOut = strOut & vbLf & "  (Restore Cleared Bonuses brings them back)"
    strOut = strOut & vbLf & "Refresh: re-run " & m_DOCK_COMMAND & " (BQ)"

Cleanup:
    BonusDockText = strOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Copied first: LogError's own handler clears Err.
    lngErrNumber = Err.Number
    strErrText = Err.Description
    LogError PROC_NAME, lngErrNumber, strErrText
    strOut = "Bonus dock failed (" & lngErrNumber & ": " & strErrText & ") - see " & LogLocation() & "."
    Resume Cleanup
End Function


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Clear Bonus From Dock
' Macro Expression:       modBonusDock.ClearBonusFromDock({{bonus_to_clear}})
'----------------------------------------------------------------------------------------------------
' Purpose: Hide one bonus from the dock by hand (skipped, given up, answered elsewhere). Accepts "2",
'          "B", "Bonus 2". Remembered in the case workbook's hidden BonusDock_Cleared name.
Public Sub ClearBonusFromDock(ByVal varBonus As Variant)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ClearBonusFromDock"

    Dim wbCase       As Workbook
    Dim udtBonuses() As tBonus
    Dim lngCount     As Long
    Dim lngIdx       As Long
    Dim strKey       As String
    Dim strCleared   As String
    Dim blnFound     As Boolean
    Dim strStatus    As String

    On Error GoTo ErrHandler

    Set wbCase = ActiveWorkbook
    If wbCase Is Nothing Then
        strStatus = "Clear Bonus From Dock: no workbook is active"
        GoTo Cleanup
    End If

    strKey = NormaliseKey(SafeText(varBonus))
    If Len(strKey) = 0 Then
        strStatus = "Clear Bonus From Dock: no bonus given (type 2, B or Bonus 2)"
        GoTo Cleanup
    End If

    lngCount = CollectBonuses(wbCase, udtBonuses, vbNullString)
    If lngCount = m_RESULT_FAILED Then
        strStatus = "Clear Bonus From Dock: could not read the bonus questions - see " & LogLocation()
        GoTo Cleanup
    End If
    For lngIdx = 1 To lngCount
        If udtBonuses(lngIdx).strKey = strKey Then blnFound = True
    Next lngIdx
    If Not blnFound Then
        strStatus = "Clear Bonus From Dock: there is no Bonus " & strKey & " in " & wbCase.Name
        GoTo Cleanup
    End If

    strCleared = ReadClearedKeys(wbCase)
    If InStr(1, strCleared, "|" & strKey & "|", vbTextCompare) = 0 Then
        If Len(strCleared) = 0 Then strCleared = "|"
        strCleared = strCleared & strKey & "|"
        If Not WriteClearedKeys(wbCase, strCleared) Then
            strStatus = "Clear Bonus From Dock: could not store the cleared bonus - see " & LogLocation()
            GoTo Cleanup
        End If
    End If
    strStatus = "Bonus " & strKey & " cleared from the dock"

Cleanup:
    If Len(strStatus) > 0 Then ShowStatus strStatus
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strStatus = "Clear Bonus From Dock failed - see " & LogLocation()
    Resume Cleanup
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Restore Cleared Bonuses
' Macro Expression:       modBonusDock.RestoreClearedBonuses()
'----------------------------------------------------------------------------------------------------
' Purpose: Bring back every bonus cleared by hand (removes the hidden BonusDock_Cleared name).
Public Sub RestoreClearedBonuses()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "RestoreClearedBonuses"

    Dim wbCase     As Workbook
    Dim strStatus  As String

    On Error GoTo ErrHandler

    Set wbCase = ActiveWorkbook
    If wbCase Is Nothing Then
        strStatus = "Restore Cleared Bonuses: no workbook is active"
        GoTo Cleanup
    End If

    If Len(ReadClearedKeys(wbCase)) = 0 Then
        strStatus = "Restore Cleared Bonuses: nothing was cleared in " & wbCase.Name
    ElseIf WriteClearedKeys(wbCase, vbNullString) Then
        strStatus = "Cleared bonuses restored to the dock"
    Else
        strStatus = "Restore Cleared Bonuses: could not remove the cleared list - see " & LogLocation()
    End If

Cleanup:
    If Len(strStatus) > 0 Then ShowStatus strStatus
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strStatus = "Restore Cleared Bonuses failed - see " & LogLocation()
    Resume Cleanup
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Next Bonus In Status Bar / Previous Bonus In Status Bar
' Macro Expression:       modBonusDock.StepBonusStatus(1) / modBonusDock.StepBonusStatus(-1)
'----------------------------------------------------------------------------------------------------
' Purpose: The no-dock fallback. Each run shows the next (or previous) open bonus on the StatusBar,
'          wrapping round. Position resets when the active workbook changes.
Public Sub StepBonusStatus(ByVal lngStep As Long)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "StepBonusStatus"

    Dim wbCase       As Workbook
    Dim udtBonuses() As tBonus
    Dim lngCount     As Long
    Dim lngIdx       As Long
    Dim lngOpen      As Long
    Dim lngOpenIdx() As Long
    Dim strStatus    As String

    On Error GoTo ErrHandler

    Set wbCase = ActiveWorkbook
    If wbCase Is Nothing Then
        strStatus = "Bonus status: no workbook is active"
        GoTo Cleanup
    End If

    lngCount = CollectBonuses(wbCase, udtBonuses, vbNullString)
    If lngCount = m_RESULT_FAILED Then
        strStatus = "Bonus status: could not read the bonus questions - see " & LogLocation()
        GoTo Cleanup
    End If
    If lngCount > 0 Then ReDim lngOpenIdx(1 To lngCount)
    For lngIdx = 1 To lngCount
        If Not udtBonuses(lngIdx).blnCleared And udtBonuses(lngIdx).lngState <> m_STATE_ANSWERED Then
            lngOpen = lngOpen + 1
            lngOpenIdx(lngOpen) = lngIdx
        End If
    Next lngIdx

    If lngOpen = 0 Then
        If lngCount = 0 Then
            strStatus = "No bonus questions found in " & wbCase.Name
        Else
            strStatus = "Every bonus in " & wbCase.Name & " is answered or cleared"
        End If
        m_lngStatusPos = 0
        GoTo Cleanup
    End If

    If m_strStatusBook <> wbCase.FullName Then
        m_strStatusBook = wbCase.FullName
        m_lngStatusPos = 0
    End If
    If lngStep = 0 Then lngStep = 1
    m_lngStatusPos = m_lngStatusPos + Sgn(lngStep)
    If m_lngStatusPos > lngOpen Or (m_lngStatusPos < 1 And lngStep > 0) Then m_lngStatusPos = 1
    If m_lngStatusPos < 1 Then m_lngStatusPos = lngOpen

    lngIdx = lngOpenIdx(m_lngStatusPos)
    strStatus = "[" & m_lngStatusPos & "/" & lngOpen & " open] " & BonusCaption(udtBonuses(lngIdx)) & ": " & _
                Replace(Replace(udtBonuses(lngIdx).strQuestion, vbCr, " "), vbLf, " ")

Cleanup:
    If Len(strStatus) > 0 Then ShowStatus strStatus
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strStatus = "Bonus status failed - see " & LogLocation()
    Resume Cleanup
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Clear Bonus Status Bar
' Macro Expression:       modBonusDock.ClearBonusStatusBar()
'----------------------------------------------------------------------------------------------------
' Purpose: Hand the StatusBar back to Excel.
Public Sub ClearBonusStatusBar()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ClearBonusStatusBar"

    On Error GoTo ErrHandler

    Application.StatusBar = False

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    Resume Cleanup
End Sub


' ================================== helpers ==================================

' Purpose: Every bonus of the case, in sheet order, with its answered / cleared state.
'          Returns the count, 0 when there is no bonus block, m_RESULT_FAILED on error.
'          Reads the primary sheet (B, else Case, else the active sheet, else the first with labels); a
'          bonus also counts as answered if its answer cell on any other bonus sheet is answered (Case
'          links to B after setup, and either may be typed into).
Private Function CollectBonuses(ByVal wbCase As Workbook, ByRef udtBonuses() As tBonus, _
                                ByRef strPrimarySheet As String) As Long

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "CollectBonuses"

    Dim colSheets   As Collection
    Dim colRowsBy   As Collection
    Dim colRows     As Collection
    Dim wsEach      As Worksheet
    Dim wsPrimary   As Worksheet
    Dim varSheet    As Variant
    Dim udtOther()  As tBonus
    Dim lngCount    As Long
    Dim lngOther    As Long
    Dim lngIdx      As Long
    Dim lngO        As Long
    Dim strCleared  As String

    On Error GoTo ErrHandler

    CollectBonuses = 0
    strPrimarySheet = vbNullString

    ' Each sheet is scanned once; its label rows are kept (keyed by sheet name) for the reads below.
    ' Hidden sheets are skipped - an answer key is often hidden, and a hidden key must never become primary.
    Set colSheets = New Collection
    Set colRowsBy = New Collection
    For Each wsEach In wbCase.Worksheets
        If wsEach.Visible = xlSheetVisible And Not IsAnswerKeySheet(wsEach.Name) Then
            Set colRows = LabelRows(wsEach)
            ' Nothing = unreadable (already logged): skip that sheet rather than sink the whole dock.
            If Not colRows Is Nothing Then
                If colRows.Count > 0 Then
                    colSheets.Add wsEach
                    colRowsBy.Add colRows, wsEach.Name
                End If
            End If
        End If
    Next wsEach
    If colSheets.Count = 0 Then GoTo Cleanup

    Set wsPrimary = PickPrimarySheet(wbCase, colSheets)
    If wsPrimary Is Nothing Then GoTo Failed
    lngCount = ReadBonusBlock(wsPrimary, colRowsBy(wsPrimary.Name), udtBonuses)
    If lngCount = m_RESULT_FAILED Then GoTo Failed
    If lngCount = 0 Then GoTo Cleanup
    strPrimarySheet = wsPrimary.Name

    For Each varSheet In colSheets
        Set wsEach = varSheet
        ' Only the Case / B pair Create Bonus Sheet makes: any other sheet with bonus labels (a key in a
        ' language IsAnswerKeySheet does not know, a stray copy) must never mark a bonus as answered.
        If wsEach.Name <> wsPrimary.Name And _
           (StrComp(wsEach.Name, "B", vbTextCompare) = 0 Or StrComp(wsEach.Name, "Case", vbTextCompare) = 0) Then
            lngOther = ReadBonusBlock(wsEach, colRowsBy(wsEach.Name), udtOther)
            If lngOther = m_RESULT_FAILED Then GoTo Failed
            For lngO = 1 To lngOther
                If udtOther(lngO).lngState = m_STATE_ANSWERED Then
                    For lngIdx = 1 To lngCount
                        If udtBonuses(lngIdx).strKey = udtOther(lngO).strKey Then
                            udtBonuses(lngIdx).lngState = m_STATE_ANSWERED
                        End If
                    Next lngIdx
                End If
            Next lngO
        End If
    Next varSheet

    strCleared = ReadClearedKeys(wbCase)
    For lngIdx = 1 To lngCount
        udtBonuses(lngIdx).blnCleared = (InStr(1, strCleared, "|" & udtBonuses(lngIdx).strKey & "|", vbTextCompare) > 0)
    Next lngIdx

    CollectBonuses = lngCount

Cleanup:
    Exit Function

Failed:
    ' A helper already logged the cause.
    CollectBonuses = m_RESULT_FAILED
    GoTo Cleanup

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description & " (" & wbCase.Name & ")"
    CollectBonuses = m_RESULT_FAILED
    Resume Cleanup
End Function


' Purpose: Read one sheet's bonus rows (colRows, from LabelRows) into udtBonuses (1-based). Returns the count,
'          m_RESULT_FAILED on error. A label repeated on the sheet (a summary list above the detail) is kept
'          once - the occurrence that carries question text.
Private Function ReadBonusBlock(ByVal wsBonus As Worksheet, ByVal colRows As Collection, _
                                ByRef udtBonuses() As tBonus) As Long

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ReadBonusBlock"

    Dim varLabelRow As Variant
    Dim varRow      As Variant
    Dim varHeader   As Variant
    Dim lngRow      As Long
    Dim lngHdrRow   As Long
    Dim lngCount    As Long
    Dim lngIdx      As Long
    Dim lngTarget   As Long
    Dim lngAnsCol   As Long
    Dim lngPtsCol   As Long
    Dim lngQCol     As Long
    Dim strKey      As String
    Dim strQuestion As String

    On Error GoTo ErrHandler

    ReadBonusBlock = 0
    If colRows Is Nothing Then
        ReadBonusBlock = m_RESULT_FAILED
        GoTo Cleanup
    End If
    If colRows.Count = 0 Then GoTo Cleanup
    ReDim udtBonuses(1 To 1)

    For Each varLabelRow In colRows
        lngRow = varLabelRow
        strKey = BonusKey(wsBonus.Cells(lngRow, m_LABEL_COL).Value2)
        If Len(strKey) > 0 Then

            ' Header row: the nearest row above carrying an Answer heading.
            lngHdrRow = 0
            lngAnsCol = m_DEFAULT_ANSWER_COL
            lngPtsCol = m_DEFAULT_POINTS_COL
            lngQCol = 0
            If lngRow > 1 Then
                varHeader = wsBonus.Range(wsBonus.Cells(Application.Max(1, lngRow - m_HEADER_LOOKBACK), 1), _
                                          wsBonus.Cells(lngRow - 1, m_MAX_COLS)).Value2
                lngHdrRow = FindHeaderRow(varHeader, lngAnsCol, lngPtsCol, lngQCol)
            End If

            varRow = wsBonus.Range(wsBonus.Cells(lngRow, 1), wsBonus.Cells(lngRow, m_MAX_COLS)).Value2
            strQuestion = vbNullString
            If lngQCol > 0 Then strQuestion = Trim$(SafeText(varRow(1, lngQCol)))
            If Len(strQuestion) = 0 Then strQuestion = LongestText(varRow, lngAnsCol + 1)

            ' Keep one entry per key.
            lngTarget = 0
            For lngIdx = 1 To lngCount
                If udtBonuses(lngIdx).strKey = strKey Then lngTarget = lngIdx
            Next lngIdx
            If lngTarget > 0 Then
                If Len(udtBonuses(lngTarget).strQuestion) > 0 Or Len(strQuestion) = 0 Then lngTarget = -1
            Else
                lngCount = lngCount + 1
                If lngCount > UBound(udtBonuses) Then ReDim Preserve udtBonuses(1 To lngCount)
                lngTarget = lngCount
            End If

            If lngTarget > 0 Then
                With udtBonuses(lngTarget)
                    .strKey = strKey
                    .strQuestion = strQuestion
                    .varPoints = varRow(1, lngPtsCol)
                    If lngHdrRow > 0 Then
                        .strRelates = RelatesText(varHeader(lngHdrRow, m_RELATES_COL), varRow(1, m_RELATES_COL))
                    Else
                        .strRelates = RelatesText(Empty, varRow(1, m_RELATES_COL))
                    End If
                    Set .rngAnswer = wsBonus.Cells(lngRow, lngAnsCol)
                    .lngState = AnswerState(.rngAnswer)
                    .blnCleared = False
                End With
            End If
        End If
    Next varLabelRow

    ReadBonusBlock = lngCount

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description & " (sheet " & wsBonus.Name & ")"
    ReadBonusBlock = m_RESULT_FAILED
    Resume Cleanup
End Function


' Purpose: Row (within varHeader) of the nearest header above a bonus label, scanning upwards, plus the
'          Answer, Points and Question columns it names. Returns 0 and leaves the defaults if none is found.
'          Header cells are short labels: a long cell that merely mentions "answer" (a question such as
'          "Provide your answer in hours") is ignored, and bonus label rows are never taken as the header.
Private Function FindHeaderRow(ByVal varHeader As Variant, ByRef lngAnsCol As Long, _
                               ByRef lngPtsCol As Long, ByRef lngQCol As Long) As Long

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "FindHeaderRow"
    Const MAX_HEADER_LEN As Long = 25

    Dim lngRow   As Long
    Dim lngCol   As Long
    Dim lngAns   As Long
    Dim lngPts   As Long
    Dim lngQ     As Long
    Dim strText  As String

    On Error GoTo ErrHandler

    FindHeaderRow = 0
    For lngRow = UBound(varHeader, 1) To 1 Step -1
        lngAns = 0
        lngPts = 0
        lngQ = 0
        If Len(BonusKey(varHeader(lngRow, m_LABEL_COL))) > 0 Then GoTo NextRow
        For lngCol = 1 To UBound(varHeader, 2)
            strText = UCase$(Trim$(Replace(SafeText(varHeader(lngRow, lngCol)), Chr$(160), " ")))
            If Len(strText) > 0 And Len(strText) <= MAX_HEADER_LEN Then
                If lngAns = 0 And InStr(1, strText, "ANSWER") > 0 And InStr(1, strText, "KEY") = 0 Then lngAns = lngCol
                If lngPts = 0 And Left$(strText, 5) = "POINT" Then lngPts = lngCol
                If lngQ = 0 And lngCol > m_LABEL_COL And InStr(1, strText, "QUESTION") > 0 And InStr(1, strText, "#") = 0 Then lngQ = lngCol
            End If
        Next lngCol
        If lngAns > 0 Then
            lngAnsCol = lngAns
            If lngPts > 0 Then lngPtsCol = lngPts
            lngQCol = lngQ
            FindHeaderRow = lngRow
            Exit For
        End If
NextRow:
    Next lngRow

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    FindHeaderRow = 0
    Resume Cleanup
End Function


' Purpose: 0 open, 1 answered, 2 error. A formula that is only a link to one other cell ("='B'!E5", "=$E$5")
'          is followed to the cell it reads, so a link to an empty answer cell (which displays 0) is still open.
'          Anything else non-empty - a constant, or a formula that is not a bare link - is answered.
Private Function AnswerState(ByVal rngAnswer As Range) As Long

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "AnswerState"

    Dim rngCurrent As Range
    Dim rngNext    As Range
    Dim varValue   As Variant
    Dim lngHop     As Long

    On Error GoTo ErrHandler

    Set rngCurrent = rngAnswer
    For lngHop = 1 To m_MAX_LINK_HOPS
        Set rngNext = Nothing
        If rngCurrent.HasFormula Then Set rngNext = LinkTarget(rngCurrent)
        If rngNext Is Nothing Then Exit For
        Set rngCurrent = rngNext
    Next lngHop

    varValue = rngCurrent.Value2
    If IsError(varValue) Then
        AnswerState = m_STATE_ERROR
    ElseIf Len(Trim$(SafeText(varValue))) = 0 Then
        AnswerState = m_STATE_OPEN
    Else
        AnswerState = m_STATE_ANSWERED
    End If

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    AnswerState = m_STATE_OPEN
    Resume Cleanup
End Function


' Purpose: The single cell a bare-link formula reads, or Nothing if the formula is anything else
'          (a calculation, a range, another workbook, a name).
Private Function LinkTarget(ByVal rngCell As Range) As Range

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "LinkTarget"
    Const LINK_PATTERN As String = "^=(?:(?:'((?:[^']|'')+)'|([^'!\[\]\s()+\-*/&^,=<>""]+))!)?(\$?[A-Z]{1,3}\$?[0-9]+)$"

    Dim objMatches As Object
    Dim strSheet   As String
    Dim wsTarget   As Worksheet

    On Error GoTo ErrHandler

    Set LinkTarget = Nothing
    ' One RegExp for the module's lifetime - this runs for every formula answer cell on every refresh.
    If m_objLinkRegex Is Nothing Then
        Set m_objLinkRegex = CreateObject("VBScript.RegExp")
        m_objLinkRegex.Pattern = LINK_PATTERN
        m_objLinkRegex.IgnoreCase = True
    End If

    Set objMatches = m_objLinkRegex.Execute(rngCell.Formula)
    If objMatches.Count = 0 Then GoTo Cleanup
    strSheet = objMatches(0).SubMatches(0)
    If Len(strSheet) > 0 Then
        strSheet = Replace(strSheet, "''", "'")
    Else
        strSheet = objMatches(0).SubMatches(1)
    End If

    If Len(strSheet) = 0 Then
        Set wsTarget = rngCell.Worksheet
    Else
        ' Probe: the text may look like a sheet name without being one in this workbook.
        On Error Resume Next
        Set wsTarget = rngCell.Worksheet.Parent.Worksheets(strSheet)
        On Error GoTo ErrHandler
    End If
    If wsTarget Is Nothing Then GoTo Cleanup

    Set LinkTarget = wsTarget.Range(objMatches(0).SubMatches(2))

Cleanup:
    Set objMatches = Nothing
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    Set LinkTarget = Nothing
    Resume Cleanup
End Function


' Purpose: "Bonus 2", "BONUS B:", "Bonus #3" -> "2" / "B" / "3". Anything else (headings such as
'          "Bonus Questions", "Bonuses") -> "".
Private Function BonusKey(ByVal varValue As Variant) As String

    Dim strText As String

    If IsError(varValue) Then Exit Function
    strText = UCase$(Trim$(Replace(SafeText(varValue), Chr$(160), " ")))
    If Left$(strText, 5) <> "BONUS" Then Exit Function
    BonusKey = NormaliseKey(strText)
End Function


' Purpose: User or sheet text to a bonus key: strips "Bonus", "Question", "#", ":" and "."; the rest must be
'          a number (1-99) or a single letter. Returns "" otherwise.
Private Function NormaliseKey(ByVal strText As String) As String

    Dim strKey As String

    strKey = UCase$(Trim$(Replace(strText, Chr$(160), " ")))
    If Left$(strKey, 5) = "BONUS" Then strKey = Trim$(Mid$(strKey, 6))
    ' "Question 2" / "Question #2" only - "Questions" (the heading) must not leave a stray "S" behind.
    If Left$(strKey, 9) = "QUESTION " Or Left$(strKey, 9) = "QUESTION#" Then strKey = Trim$(Mid$(strKey, 9))
    strKey = Trim$(Replace(Replace(Replace(strKey, "#", ""), ":", ""), ".", ""))
    ' "B4" shorthand (as in the AB1-AB5 launch codes). A lone "B" stays the letter key.
    If strKey Like "B#" Or strKey Like "B##" Then strKey = Mid$(strKey, 2)

    If strKey Like "#" Or strKey Like "##" Then
        NormaliseKey = CStr(CLng(strKey))
    ElseIf strKey Like "[A-Z]" Then
        NormaliseKey = strKey
    Else
        NormaliseKey = vbNullString
    End If
End Function


' Purpose: Rows of column B holding a bonus label ("Bonus 2", "Bonus A"), top to bottom; an empty collection when
'          there are none; Nothing on error. Speed: a native COUNTIF rules out a sheet in one call. With a few
'          "bonus" mentions, MATCH jumps between them and stops at the last one; with many (a data set that
'          mentions bonuses), or if MATCH stops early, column B is read once into an array and scanned in memory.
'          Bounded by the used range, which over-reports at worst, so filtered or hidden rows are still covered.
'          No Range.Find - it would overwrite the Find dialog.
Private Function LabelRows(ByVal wsRead As Worksheet) As Collection

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "LabelRows"

    Dim colRows     As Collection
    Dim varColumn   As Variant
    Dim lngMentions As Long
    Dim lngHits     As Long
    Dim lngLast     As Long
    Dim lngFrom     As Long
    Dim lngRow      As Long
    Dim varHit      As Variant

    On Error GoTo ErrHandler

    Set colRows = New Collection
    lngMentions = Application.WorksheetFunction.CountIf(wsRead.Columns(m_LABEL_COL), m_LABEL_PATTERN)
    If lngMentions = 0 Then GoTo Done
    lngLast = wsRead.UsedRange.Row + wsRead.UsedRange.Rows.Count - 1

    If lngMentions <= m_MAX_MATCH_HOPS Then
        lngFrom = 1
        Do While lngFrom <= lngLast And lngHits < lngMentions
            varHit = Application.Match(m_LABEL_PATTERN, _
                                       wsRead.Range(wsRead.Cells(lngFrom, m_LABEL_COL), wsRead.Cells(lngLast, m_LABEL_COL)), 0)
            If IsError(varHit) Then Exit Do
            lngHits = lngHits + 1
            lngRow = lngFrom + CLng(varHit) - 1
            If Len(BonusKey(wsRead.Cells(lngRow, m_LABEL_COL).Value2)) > 0 Then colRows.Add lngRow
            lngFrom = lngRow + 1
        Loop
        If lngHits = lngMentions Then GoTo Done
        ' MATCH found fewer than COUNTIF counted: start again with the full scan below.
        Set colRows = New Collection
    End If

    If lngLast = 1 Then
        If Len(BonusKey(wsRead.Cells(1, m_LABEL_COL).Value2)) > 0 Then colRows.Add 1
    Else
        varColumn = wsRead.Range(wsRead.Cells(1, m_LABEL_COL), wsRead.Cells(lngLast, m_LABEL_COL)).Value2
        For lngRow = 1 To lngLast
            If Len(BonusKey(varColumn(lngRow, 1))) > 0 Then colRows.Add lngRow
        Next lngRow
    End If

Done:
    Set LabelRows = colRows

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description & " (sheet " & wsRead.Name & ")"
    Set LabelRows = Nothing
    Resume Cleanup
End Function


' Purpose: Which bonus sheet the dock describes: B (Create Bonus Sheet's copy), else Case, else the active
'          sheet if it has bonuses, else one with "Case" in its name, else the first that has bonuses. Nothing on error. Tab names are used
'          deliberately: these are other people's case files, so CodeNames are unknown.
Private Function PickPrimarySheet(ByVal wbCase As Workbook, ByVal colSheets As Collection) As Worksheet

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "PickPrimarySheet"

    Dim varSheet  As Variant
    Dim wsEach    As Worksheet
    Dim strActive As String

    On Error GoTo ErrHandler

    Set PickPrimarySheet = Nothing
    If Not wbCase.ActiveSheet Is Nothing Then strActive = wbCase.ActiveSheet.Name

    For Each varSheet In colSheets
        Set wsEach = varSheet
        If StrComp(wsEach.Name, "B", vbTextCompare) = 0 Then Set PickPrimarySheet = wsEach: GoTo Cleanup
    Next varSheet
    For Each varSheet In colSheets
        Set wsEach = varSheet
        If StrComp(wsEach.Name, "Case", vbTextCompare) = 0 Then Set PickPrimarySheet = wsEach: GoTo Cleanup
    Next varSheet
    For Each varSheet In colSheets
        Set wsEach = varSheet
        If wsEach.Name = strActive Then Set PickPrimarySheet = wsEach: GoTo Cleanup
    Next varSheet
    ' "Case File", "Case (English)" ...
    For Each varSheet In colSheets
        Set wsEach = varSheet
        If InStr(1, wsEach.Name, "Case", vbTextCompare) > 0 Then Set PickPrimarySheet = wsEach: GoTo Cleanup
    Next varSheet
    Set PickPrimarySheet = colSheets(1)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    Set PickPrimarySheet = Nothing
    Resume Cleanup
End Function


' Purpose: Answer keys, solution sheets and backups are never read - the dock must not surface answers, and a
'          key must never mark a bonus as answered. Names seen in the library: Answers, Answer, HiddenAnswers,
'          AnswersBU, Antworten (DACH); the other languages are the competitions' own (ES / PT / FR / IT).
'          Backups: CompBot's Backup makes BU_<name>; the library also has CaseBU / AnswersBU.
Private Function IsAnswerKeySheet(ByVal strName As String) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const KEY_WORDS As String = "ANSWER|ANTWORT|RESPUESTA|RESPOSTA|REPONSE|RISPOST|GABARITO|SOLUTION|SOLUCI|SOLUZION|LOESUNG|LOSUNG"

    Dim strUpper As String
    Dim strWords As String
    Dim varWord  As Variant

    strUpper = UCase$(Trim$(strName))
    If strUpper Like "KEY*" Or strUpper Like "BU[_]*" Or strUpper Like "*BU" Then
        IsAnswerKeySheet = True
        Exit Function
    End If
    ' Accented forms (French reponse with E-acute, Portuguese solucao with C-cedilla) are built at run time so
    ' the source stays ASCII on any code page.
    strWords = KEY_WORDS & "|R" & ChrW$(&HC9) & "PONSE|SOLU" & ChrW$(&HC7) & "|L" & ChrW$(&HD6) & "SUNG"
    For Each varWord In Split(strWords, "|")
        If InStr(1, strUpper, varWord, vbBinaryCompare) > 0 Then
            IsAnswerKeySheet = True
            Exit Function
        End If
    Next varWord
End Function


' Purpose: "Bonus 2 <middle dot> 20 pts <middle dot> Relates to L3" - the one-line caption for the dock and StatusBar.
Private Function BonusCaption(ByRef udtBonus As tBonus) As String

    Dim strCaption As String

    strCaption = "Bonus " & udtBonus.strKey
    If IsNumeric(udtBonus.varPoints) And Not IsEmpty(udtBonus.varPoints) Then
        strCaption = strCaption & MidDot() & Format$(udtBonus.varPoints, "0") & " pts"
    End If
    If Len(udtBonus.strRelates) > 0 Then strCaption = strCaption & MidDot() & udtBonus.strRelates
    BonusCaption = strCaption
End Function


' Purpose: The relates-to column as readable text, led by its own header: "Relates to L" + 3 -> "Relates to L3";
'          "Level" + 2 -> "Level 2"; "Relates to L" + "L03" -> "Relates to L03"; "Available to" + "All levels" ->
'          "Available to All levels"; no header -> the bare value.
Private Function RelatesText(ByVal varHeader As Variant, ByVal varValue As Variant) As String

    Dim strValue  As String
    Dim strHeader As String
    Dim lngSpace  As Long

    strValue = Trim$(Replace(Replace(SafeText(varValue), vbCr, " "), vbLf, " "))
    ' "N/A" / "-" say the bonus relates to nothing in particular: leave the caption without it.
    Select Case UCase$(Replace(strValue, " ", vbNullString))
        Case vbNullString, "N/A", "NA", "-"
            Exit Function
    End Select

    strHeader = Trim$(Replace(Replace(Replace(SafeText(varHeader), vbCr, " "), vbLf, " "), Chr$(160), " "))
    If Right$(strHeader, 1) = "#" Then strHeader = Trim$(Left$(strHeader, Len(strHeader) - 1))
    Do While InStr(1, strHeader, "  ") > 0
        strHeader = Replace(strHeader, "  ", " ")
    Loop
    If Len(strHeader) = 0 Then
        RelatesText = strValue
        Exit Function
    End If

    lngSpace = InStrRev(strHeader, " ")
    If Len(strHeader) - lngSpace = 1 Then
        ' Header ends in a one-letter unit ("Relates to L"): glue a number on, let a text value replace the letter.
        If IsNumeric(strValue) Then
            RelatesText = strHeader & strValue
        Else
            RelatesText = Left$(strHeader, lngSpace) & strValue
        End If
    Else
        RelatesText = strHeader & " " & strValue
    End If
End Function


' Purpose: The longest text cell in a row from column lngFromCol on - the fallback for a missing Question header.
Private Function LongestText(ByVal varRow As Variant, ByVal lngFromCol As Long) As String

    Dim lngCol  As Long
    Dim strText As String

    For lngCol = lngFromCol To UBound(varRow, 2)
        strText = Trim$(SafeText(varRow(1, lngCol)))
        If Len(strText) > Len(LongestText) And Not IsNumeric(strText) Then LongestText = strText
    Next lngCol
End Function


' Purpose: The hand-cleared keys stored in the case workbook, as "|2|B|", or "" when none (or unreadable).
Private Function ReadClearedKeys(ByVal wbCase As Workbook) As String

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ReadClearedKeys"

    Dim strRefersTo As String

    On Error GoTo ErrHandler

    ' Probe: the name only exists once something has been cleared.
    On Error Resume Next
    strRefersTo = wbCase.Names(m_CLEARED_NAME).RefersTo
    On Error GoTo ErrHandler

    If Left$(strRefersTo, 2) = "=""" And Right$(strRefersTo, 1) = """" Then
        ReadClearedKeys = Mid$(strRefersTo, 3, Len(strRefersTo) - 3)
    End If

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    ReadClearedKeys = vbNullString
    Resume Cleanup
End Function


' Purpose: Store the hand-cleared keys as a hidden workbook-level name; "" removes the name. True on success.
Private Function WriteClearedKeys(ByVal wbCase As Workbook, ByVal strKeys As String) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "WriteClearedKeys"

    On Error GoTo ErrHandler

    WriteClearedKeys = False
    If Len(strKeys) = 0 Then
        ' Probe: nothing to delete if the name was never created.
        On Error Resume Next
        wbCase.Names(m_CLEARED_NAME).Delete
        On Error GoTo ErrHandler
    Else
        wbCase.Names.Add Name:=m_CLEARED_NAME, RefersTo:="=""" & strKeys & """", Visible:=False
    End If
    WriteClearedKeys = True

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    WriteClearedKeys = False
    Resume Cleanup
End Function


' Purpose: Put a message on the StatusBar (cut to Excel's 255-character limit). It stays until the next message
'          or Clear Bonus Status Bar (BQS) - no timer, so nothing is ever scheduled behind Jaq's back.
Private Sub ShowStatus(ByVal strStatus As String)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ShowStatus"

    On Error GoTo ErrHandler

    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: ", 1, 3" -> "Bonus 1, 3"; "" -> "none".
Private Function KeyListText(ByVal strList As String) As String

    If Len(strList) = 0 Then
        KeyListText = "none"
    Else
        KeyListText = "Bonus " & Mid$(strList, 3)
    End If
End Function


' Purpose: Break text into lines of at most m_WRAP_WIDTH characters at spaces (a longer single word is cut).
'          Existing line breaks are kept. The task pane's Text view does not wrap on its own.
Private Function WrapText(ByVal strText As String) As String

    Dim varLines As Variant
    Dim varWords As Variant
    Dim lngLine  As Long
    Dim lngWord  As Long
    Dim strLine  As String
    Dim strWord  As String
    Dim strOut   As String

    varLines = Split(Replace(strText, vbCr, vbNullString), vbLf)
    For lngLine = LBound(varLines) To UBound(varLines)
        If lngLine > LBound(varLines) Then strOut = strOut & vbLf
        strLine = vbNullString
        varWords = Split(varLines(lngLine), " ")
        For lngWord = LBound(varWords) To UBound(varWords)
            strWord = varWords(lngWord)
            Do While Len(strWord) > m_WRAP_WIDTH
                If Len(strLine) > 0 Then strOut = strOut & strLine & vbLf: strLine = vbNullString
                strOut = strOut & Left$(strWord, m_WRAP_WIDTH) & vbLf
                strWord = Mid$(strWord, m_WRAP_WIDTH + 1)
            Loop
            If Len(strLine) = 0 Then
                strLine = strWord
            ElseIf Len(strLine) + 1 + Len(strWord) <= m_WRAP_WIDTH Then
                strLine = strLine & " " & strWord
            Else
                strOut = strOut & strLine & vbLf
                strLine = strWord
            End If
        Next lngWord
        strOut = strOut & strLine
    Next lngLine
    WrapText = strOut
End Function


' Purpose: " <middle dot> " separator (ChrW, so the source stays ASCII).
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


