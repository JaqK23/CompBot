Attribute VB_Name = "modCaseSetup"
Option Explicit
Option Base 1

' --- MODULE CONSTANTS ---
Private Const m_DEBUG_MODE          As Boolean = False
Private Const m_BACKUP_PREFIX       As String = "BU_"
Private Const m_MAX_SHEET_NAME      As Long = 31
Private Const m_MAX_STATUS          As Long = 250   ' Application.StatusBar rejects long strings
' Full Setup Case's working copy: Case.xlsx is saved as Case_Solve.xlsx before anything else runs.
Private Const m_COPY_SUFFIX         As String = "Solve"
' Level markers this close to a neighbouring marker are a Contents list, not level starts.
' Measured over all 152 library cases (crossref\level_gaps.txt, 2026-09-14): marker gaps are
' either 1 row (Contents lists, 9 cases, all 2026 UK) or 15+ rows (real levels). Nothing between.
Private Const m_CONTENTS_MAX_GAP    As Long = 4
' The words a case may use to label a worked-example row, matched at the LEFT of the cell.
' English, French, Brazilian Portuguese. ADD A LANGUAGE HERE AND NOWHERE ELSE - every example-row
' test in CompBot goes through ExamplePrefixLen, which reads this list.
' Measured in the 152-case library: the 2025/2026 Brazilian cases use "Exemplo1".."Exemplo7" and
' carry no English "Example" anywhere. Note the words happen to share a length today; code must
' still take the offset from the prefix that matched, never a hardcoded 7.
' "Sample" is not a translation - it is a different English word, used by older cases
' (2021 Lumberjack, 2021 Excelopolis, 2022 Square of Fortune). Jaq, 2026-09-22: those are
' unlikely to recur in competition but are worth having for training on past cases.
Private Const m_EXAMPLE_WORDS       As String = "Example|Exemple|Exemplo|Sample"
Private Const m_ERR_CASE_SHEET      As Long = vbObjectError + 513
Private Const m_ERR_NOT_FOUND       As Long = vbObjectError + 514

' --- MODULE VARIABLES ---
Private m_blnInSetup                As Boolean  ' True while Setup is running its steps
Private m_blnCaseAsked              As Boolean  ' the case-sheet InputBox has been shown this Setup run
Private m_strCaseSheet              As String   ' the answer given to it
Private m_strStepError              As String   ' last failure recorded by a step, read by Setup
Private m_strStepNote               As String   ' a step that skipped itself on purpose, read by Setup
Private m_blnLogWritten             As Boolean  ' LogError actually wrote a line this Setup run

'general case setup
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Setup Case
' Description:            Setup.
' Macro Expression:       modCaseSetup.Setup()
' Generated:              11/15/2024 08:41 PM
'----------------------------------------------------------------------------------------------------
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Setup Case
' Description:            Setup case by running Backup and Level sheet creation
' Macro Expression:       modCaseSetup.Setup()
' Generated:              01/03/2025 08:10 PM
'----------------------------------------------------------------------------------------------------
' Purpose: Run every setup step in turn. A step that fails is logged and reported on the
'          StatusBar, and the remaining steps still run. Setup itself never raises, so the
'          command chained after it (Import Case Lambdas) always runs.
'          The first step saves a _Solve working copy (Jaq, 2026-09-24), so the backups, level
'          sheets and the solve all happen in the copy and the original file is never touched.
'          Sheet renaming and unmerging are deliberately NOT part of setup (Jaq, 2026-09-14,
'          2026-09-24). Each step can be switched off per user in Setup Case Settings (SCS);
'          the step names below are the setting keys (modSetupSettings.SettingKeys).
Public Sub Setup()

    Dim wbCase As Workbook
    Dim strFailed As String
    Dim strNotes As String
    Dim strMsg As String
    Dim varCalc As Variant
    Dim blnScreen As Boolean
    Dim blnAlerts As Boolean
    Dim blnEvents As Boolean

    On Error GoTo ErrHandler

    ' Setup owns the application state: whatever a failing step leaves behind is put back here.
    Set wbCase = ActiveWorkbook
    varCalc = Application.Calculation
    blnScreen = Application.ScreenUpdating
    blnAlerts = Application.DisplayAlerts
    blnEvents = Application.EnableEvents

    m_blnInSetup = True
    m_blnCaseAsked = False
    m_strCaseSheet = vbNullString
    m_blnLogWritten = False

    RunSetupStep "SaveCopy", strFailed, strNotes
    RunSetupStep "Backup", strFailed, strNotes
    RunSetupStep "NameAllUsedRanges", strFailed, strNotes
    RunSetupStep "CreateLevelSheets", strFailed, strNotes
    RunSetupStep "CreateBonusSheet", strFailed, strNotes
    RunSetupStep "CreateCaseInputsSheet", strFailed, strNotes

Cleanup:
    On Error Resume Next
    Application.Calculation = varCalc
    Application.ScreenUpdating = blnScreen
    Application.DisplayAlerts = blnAlerts
    Application.EnableEvents = blnEvents
    m_blnInSetup = False
    m_blnCaseAsked = False
    m_strCaseSheet = vbNullString
    If Len(strFailed) = 0 Then
        strMsg = "Full Setup Case: all steps ran."
    ElseIf m_blnLogWritten Then
        ' Step names only - the detail is in the log, and the StatusBar has a length limit.
        strMsg = "Full Setup Case: FAILED" & StepNames(strFailed) & ". Other steps ran. See " & LogLocation()
    Else
        strMsg = "Full Setup Case: FAILED -" & strFailed & " Other steps ran."
    End If
    If Len(strNotes) > 0 Then strMsg = strMsg & " Skipped -" & strNotes
    Application.StatusBar = Left$(strMsg, m_MAX_STATUS)

    ' Land on a sheet ready to work: the bonus sheet if one was made, else Level 1.
    ' Activate is deliberate here - putting the user on a sheet is the purpose.
    If Not wbCase Is Nothing Then
        ' Worksheet.Activate fails if another workbook's window is in front - bring the case forward first.
        wbCase.Activate
        If SheetNameTaken(wbCase, "B") Then
            wbCase.Worksheets("B").Activate
        ElseIf SheetNameTaken(wbCase, "L01") Then
            wbCase.Worksheets("L01").Activate
        End If
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Capture first - LogError's own error handling resets Err.
    strMsg = Err.Number & ": " & Err.Description
    If LogError("Setup", Err.Number, Err.Description) Then m_blnLogWritten = True
    strFailed = strFailed & " Setup (" & strMsg & ");"
    Resume Cleanup
End Sub

' Purpose: Run one setup step, isolated. Errors the step does not handle itself land here;
'          errors it does handle are passed back through m_strStepError.
Private Sub RunSetupStep(ByVal strStep As String, ByRef strFailed As String, ByRef strNotes As String)

    On Error GoTo ErrHandler

    m_strStepError = vbNullString
    m_strStepNote = vbNullString

    ' Switched off in Setup Case Settings (SCS): skipped, and said so (Jaq, 2026-09-24).
    If Not modSetupSettings.SetupStepOn(strStep) Then
        m_strStepNote = "off in SCS"
        GoTo Cleanup
    End If

    Select Case strStep
        Case "SaveCopy":                SaveSolveCopy
        Case "Backup":                  Backup
        Case "NameAllUsedRanges":       NameAllUsedRanges
        Case "CreateLevelSheets":       CreateLevelSheets
        Case "CreateBonusSheet":        CreateBonusSheet
        Case "CreateCaseInputsSheet":   CreateCaseInputsSheet "detailed"
    End Select

Cleanup:
    On Error Resume Next
    If Len(m_strStepError) > 0 Then
        strFailed = strFailed & " " & strStep & " (" & m_strStepError & ");"
    End If
    If Len(m_strStepNote) > 0 Then
        strNotes = strNotes & " " & strStep & " (" & m_strStepNote & ");"
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    m_strStepError = Err.Number & ": " & Err.Description
    If LogError(strStep, Err.Number, Err.Description) Then m_blnLogWritten = True
    Resume Cleanup
End Sub

' Purpose: Setup's first step. Saves the case as a working copy (Save As, suffix m_COPY_SUFFIX) so
'          every later step, and the solve, happen in the copy and the original file is never
'          touched (Jaq, 2026-09-24). Skipped, with a note, when the workbook has never been saved
'          (no folder to write into) or is already a working copy (Setup run a second time, or a
'          Save Copy of File (SA) copy with the default suffix).
Private Sub SaveSolveCopy()

    Dim wb As Workbook
    Dim strOld As String

    On Error GoTo ErrHandler

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        m_strStepNote = "no active workbook"
        GoTo Cleanup
    End If
    If Len(wb.Path) = 0 Then
        m_strStepNote = "never saved, so no copy made"
        GoTo Cleanup
    End If
    If IsWorkingCopyName(wb.Name) Then
        m_strStepNote = "already a working copy"
        GoTo Cleanup
    End If

    strOld = wb.FullName
    SaveCopy m_COPY_SUFFIX
    ' SaveCopy reports its own refusals on the StatusBar and returns quietly; Setup's own message
    ' would overwrite that, so a copy that did not happen is recorded as a failure here.
    If StrComp(wb.FullName, strOld, vbTextCompare) = 0 Then m_strStepError = "no copy saved"

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    m_strStepError = Err.Number & ": " & Err.Description
    If LogError("SaveSolveCopy", Err.Number, Err.Description) Then m_blnLogWritten = True
    Resume Cleanup
End Sub

' Purpose: True if a file name is already a working copy: its base ends "_Solve" or "_Working",
'          optionally followed by SaveCopy's "_2".."_99" counter. Case-insensitive.
Private Function IsWorkingCopyName(ByVal strName As String) As Boolean

    Dim varSuffix As Variant
    Dim strBase As String
    Dim lngDot As Long

    On Error GoTo ErrHandler

    lngDot = InStrRev(strName, ".")
    If lngDot > 0 Then
        strBase = LCase$(Left$(strName, lngDot - 1))
    Else
        strBase = LCase$(strName)
    End If

    For Each varSuffix In Array(m_COPY_SUFFIX, "Working")
        If strBase Like "*_" & LCase$(varSuffix) Or _
           strBase Like "*_" & LCase$(varSuffix) & "_#" Or _
           strBase Like "*_" & LCase$(varSuffix) & "_##" Then
            IsWorkingCopyName = True
            GoTo Cleanup
        End If
    Next varSuffix

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    IsWorkingCopyName = False
    Resume Cleanup
End Function

'rename all sheets (to shorter names)
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Rename Sheets
' Description:            Shortens multi-word sheet names to their initials (Level 1 Data becomes L1D)
' Macro Expression:       modCaseSetup.RenameSht()
' Generated:              01/25/2025 07:41 PM
'----------------------------------------------------------------------------------------------------
' Purpose: shorten every multi-word sheet name to one letter per word, so cross-sheet references
'          need no quotes: 'Level 1 Data'!B4 becomes L1D!B4. Words split on spaces and underscores.
'          Rewritten 2026-09-24 (Jaq):
'            - a word's DIGITS are kept after its first letter, so Level1 Data -> L1D and
'              Level2 Data -> L2D, where both used to want LD;
'            - backup sheets (BU_...) are never renamed, so the original names stay readable;
'            - a name already taken gets _2, _3... on the end rather than the sheet being skipped
'              silently. Case, Case-Varsity and Answers are left alone, as before.
'          Excel updates every formula that refers to a renamed sheet. Reports on the status bar.
Public Sub RenameSht()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "RenameSht"
    Const CMD_TITLE As String = "Rename Sheets"

    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim strName     As String
    Dim strNewName  As String
    Dim strDone     As String
    Dim lngRenamed  As Long
    Dim strStatus   As String

    On Error GoTo ErrHandler
    VBAInit

    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        strStatus = CMD_TITLE & ": no workbook is active."
        GoTo Cleanup
    End If

    For Each ws In wb.Worksheets
        strName = ws.Name
        If KeepSheetName(strName) Then GoTo NextSheet

        strNewName = ShortSheetName(strName)
        If Len(strNewName) = 0 Or StrComp(strNewName, strName, vbBinaryCompare) = 0 Then GoTo NextSheet
        strNewName = FreeSheetName(wb, strNewName)

        ws.Name = strNewName
        lngRenamed = lngRenamed + 1
        strDone = strDone & ", " & strName & " > " & strNewName
NextSheet:
    Next ws

    If lngRenamed = 0 Then
        strStatus = CMD_TITLE & ": no multi-word sheet names to shorten."
    Else
        strStatus = CMD_TITLE & ": " & lngRenamed & " renamed: " & Mid$(strDone, 3)
    End If

Cleanup:
    VBAFin                                        ' clears the status bar, so report after it
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strStatus = CMD_TITLE & " failed at " & strName & ": " & Err.Description & " (see " & LogLocation() & ")."
    Resume Cleanup
End Sub

' Purpose: True for a sheet Rename Sheets must not touch: the case sheet, the answers, a
'          backup, or a name that is already one word. A clash counter of its own (L1D_2) does
'          not count as a second word, so a second run leaves its own names alone.
Private Function KeepSheetName(ByVal strName As String) As Boolean

    Dim lngU As Long

    lngU = InStrRev(strName, "_")
    If lngU > 1 Then
        If Mid$(strName, lngU + 1) Like "#" Or Mid$(strName, lngU + 1) Like "##" Then strName = Left$(strName, lngU - 1)
    End If

    If StrComp(strName, "Case", vbTextCompare) = 0 _
       Or StrComp(strName, "Case-Varsity", vbTextCompare) = 0 _
       Or StrComp(strName, "Answers", vbTextCompare) = 0 _
       Or StrComp(Left$(strName, Len(m_BACKUP_PREFIX)), m_BACKUP_PREFIX, vbTextCompare) = 0 Then
        KeepSheetName = True
    Else
        KeepSheetName = (InStr(1, strName, " ") = 0 And InStr(1, strName, "_") = 0)
    End If
End Function

' Purpose: one character per word (upper case), followed by any digits later in that word:
'          "Level 1 Data" -> "L1D", "Level1 Data" -> "L1D", "Level 10 Data" -> "L10D".
Private Function ShortSheetName(ByVal strName As String) As String

    Dim varWord As Variant
    Dim strWord As String
    Dim strOut  As String
    Dim lngI    As Long

    For Each varWord In Split(Replace(strName, "_", " "), " ")
        strWord = CStr(varWord)
        If Len(strWord) > 0 Then
            strOut = strOut & UCase$(Left$(strWord, 1))
            For lngI = 2 To Len(strWord)
                If Mid$(strWord, lngI, 1) Like "#" Then strOut = strOut & Mid$(strWord, lngI, 1)
            Next lngI
        End If
    Next varWord
    ShortSheetName = Left$(strOut, m_MAX_SHEET_NAME)
End Function

' Purpose: strName if no sheet has it yet, otherwise strName_2, strName_3... (sheet names ignore case).
'          The underscore keeps the counter off the name's own digits (L1 twice -> L1_2, not L12,
'          which a Level 12 sheet wants) and stops LD2 reading as a cell address.
Private Function FreeSheetName(ByVal wb As Workbook, ByVal strName As String) As String

    Dim lngN    As Long
    Dim strTry  As String

    strTry = strName
    lngN = 1
    Do While SheetNameTaken(wb, strTry)
        lngN = lngN + 1
        strTry = Left$(strName, m_MAX_SHEET_NAME - Len(CStr(lngN)) - 1) & "_" & lngN
    Loop
    FreeSheetName = strTry
End Function

'backup all sheets (to prevent overwrite errors)
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Backup sheets in workbook
' Description:            Copy all sheets in workbook
' Macro Expression:       modCaseSetup.Backup()
' Generated:              11/15/2024 08:44 PM
'----------------------------------------------------------------------------------------------------
' Purpose: Copy every worksheet to BU_<name>. The prefix goes at the front so that cropping a
'          long name at Excel's 31-character limit still leaves a distinct, recognisable name.
'          Sheets already starting BU_ are skipped (a second run does not back up backups);
'          very hidden sheets cannot be copied and are skipped.
Public Sub Backup()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "Backup"

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsAny As Worksheet
    Dim wsNew As Worksheet
    Dim colSheets As Collection
    Dim dicBefore As Object
    Dim varSheet As Variant
    Dim blnInit As Boolean
    Dim strErr As String

    On Error GoTo ErrHandler

    ' VBAInit turns DisplayAlerts off, so a copied sheet-scoped name clashing with a
    ' workbook name cannot raise Excel's "name already exists" prompt.
    VBAInit
    blnInit = True

    Set wb = ActiveWorkbook

    ' Snapshot the originals first - the copy loop adds sheets as it goes.
    Set colSheets = New Collection
    For Each ws In wb.Worksheets
        If StrComp(Left$(ws.Name, Len(m_BACKUP_PREFIX)), m_BACKUP_PREFIX, vbTextCompare) <> 0 _
        And ws.Visible <> xlSheetVeryHidden Then
            colSheets.Add ws
        End If
    Next ws

    For Each varSheet In colSheets
        Set ws = varSheet

        ' Identify the copy as the one sheet that was not there before, rather than
        ' trusting it to land last (a hidden last sheet can change where it goes).
        Set dicBefore = CreateObject("Scripting.Dictionary")
        dicBefore.CompareMode = vbTextCompare
        For Each wsAny In wb.Worksheets
            dicBefore(wsAny.Name) = True
        Next wsAny

        ws.Copy After:=wb.Worksheets(wb.Worksheets.Count)

        Set wsNew = Nothing
        For Each wsAny In wb.Worksheets
            If Not dicBefore.Exists(wsAny.Name) Then
                Set wsNew = wsAny
                Exit For
            End If
        Next wsAny
        If wsNew Is Nothing Then Err.Raise m_ERR_NOT_FOUND, PROC_NAME, "Copy of " & ws.Name & " not found"

        wsNew.Name = BackupSheetName(wb, ws.Name)
    Next varSheet

Cleanup:
    On Error Resume Next
    If blnInit Then VBAFin
    If Len(strErr) > 0 Then
        m_strStepError = strErr
        Application.StatusBar = Left$(PROC_NAME & " failed - " & strErr, m_MAX_STATUS)
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, Err.Description) Then m_blnLogWritten = True
    Resume Cleanup
End Sub

' Purpose: BU_ + the name, cropped to 31 characters. If that is already taken (two long names
'          sharing their first 28 characters), shorten further and add ~2, ~3 ...
Private Function BackupSheetName(ByVal wb As Workbook, ByVal strName As String) As String

    ' --- CONSTANTS (local to function) ---
    Const MAX_TRIES As Long = 999

    Dim strBase As String
    Dim strTry As String
    Dim lngN As Long

    On Error GoTo ErrHandler

    strBase = m_BACKUP_PREFIX & Left$(strName, m_MAX_SHEET_NAME - Len(m_BACKUP_PREFIX))
    strTry = strBase
    lngN = 1
    Do While SheetNameTaken(wb, strTry) And lngN < MAX_TRIES
        lngN = lngN + 1
        strTry = Left$(strBase, m_MAX_SHEET_NAME - Len(CStr(lngN)) - 1) & "~" & lngN
    Loop

    BackupSheetName = strTry
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If LogError("BackupSheetName", Err.Number, Err.Description) Then m_blnLogWritten = True
    BackupSheetName = vbNullString
End Function

' Purpose: True if any sheet in WB already has this name (sheet names ignore case).
Private Function SheetNameTaken(ByVal wb As Workbook, ByVal strName As String) As Boolean

    Dim objSheet As Object

    On Error GoTo ErrHandler

    For Each objSheet In wb.Sheets
        If StrComp(objSheet.Name, strName, vbTextCompare) = 0 Then
            SheetNameTaken = True
            Exit Function
        End If
    Next objSheet
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    SheetNameTaken = False
End Function

' Purpose: Give WSTO the same column widths as WSFROM, from column A to WSFROM's last used
'          column. Copying whole rows carries row heights and formats but NOT widths, so a
'          level or bonus sheet opened at the default width and cut off the wrapped question
'          text (Jaq, 2026-09-23). A hidden column (width 0) stays hidden.
'          A failure costs only the widths: it is logged and the sheet build carries on.
Private Sub CopyColumnWidths(ByVal wsFrom As Worksheet, ByVal wsTo As Worksheet)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "CopyColumnWidths"

    Dim lngLastCol As Long
    Dim lngCol As Long

    On Error GoTo ErrHandler

    With wsFrom.UsedRange
        lngLastCol = .Column + .Columns.Count - 1
    End With
    For lngCol = 1 To lngLastCol
        wsTo.Columns(lngCol).ColumnWidth = wsFrom.Columns(lngCol).ColumnWidth
    Next lngCol
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If LogError(PROC_NAME, Err.Number, Err.Description) Then m_blnLogWritten = True
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create level sheets
' Description:            Creates a sheet for each level in the Case sheet
' Macro Expression:       modCaseSetup.CreateLevelSheets()
' Generated:              11/15/2024 09:59 PM
'----------------------------------------------------------------------------------------------------
Public Sub CreateLevelSheets()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "CreateLevelSheets"

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim lngLevels As Long
    Dim lngCurr As Long
    Dim lngRow As Long
    Dim lngBegRow As Long
    Dim lngEndRow As Long
    Dim lngLastRow As Long
    Dim colLRows As Collection
    Dim rngColB As Range
    Dim rngAnswers As Range
    Dim rngNumbers As Range
    Dim rngScoreCell As Range
    Dim objCell As Range
    Dim rngScore As Range
    Dim lngACol As Long
    Dim booHdrs As Boolean
    Dim blnLink As Boolean
    Dim colColList As Collection
    Dim lngCol As Long
    Dim lngEndCol As Long
    Dim rngHdrRow As Range
    Dim rngCell As Range
    Dim excludedHeaders As Object
    Dim rngRow As Range
    Dim strHeader As String
    Dim varCol As Variant
    Dim strCurrVal As String
    Dim shp As Shape
    Dim blnInit As Boolean
    Dim strErr As String

    On Error GoTo ErrHandler

    'Initialize
    VBAInit
    blnInit = True
    ' Initialize excluded headers as a dictionary
    Set excludedHeaders = CreateObject("Scripting.Dictionary")
    excludedHeaders.Add "Answer", True
    excludedHeaders.Add "Level", True
    excludedHeaders.Add "Points", True
    'set up workbook and worksheet
    Set wb = ActiveWorkbook
    Set ws = GetCaseSheet(wb)
    If ws Is Nothing Then Err.Raise m_ERR_CASE_SHEET, PROC_NAME, "Case sheet not found"

    '___________________worksheet exists_____________________
    'get current score cell - probe, it is only there in training files
    On Error Resume Next
    Set rngScore = ws.UsedRange.Find(What:="Current Score", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Offset(1, 0)
    On Error GoTo ErrHandler

    'find each level's start row in column B (Contents-list entries already dropped)
    lngLastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Set colLRows = LevelMarkerRows(ws)
    If colLRows Is Nothing Then Err.Raise m_ERR_NOT_FOUND, PROC_NAME, "Could not read the level headings in column B of " & ws.Name

    'count levels
    lngLevels = colLRows.Count
    If lngLevels = 0 Then Err.Raise m_ERR_NOT_FOUND, PROC_NAME, "No 'Level #' or 'Section #' cells found in column B of " & ws.Name

    ' Refuse before creating anything if level sheets already exist (a re-run) -
    ' otherwise Sheets.Add succeeds, the rename fails, and a blank sheet is left behind.
    For lngCurr = 1 To lngLevels
        If SheetNameTaken(wb, "L" & Format$(lngCurr, "00")) Then
            Err.Raise m_ERR_NOT_FOUND, PROC_NAME, "Sheet L" & Format$(lngCurr, "00") & " already exists - level sheets not created"
        End If
    Next lngCurr
    lngCurr = 0

    'loop through the levels to create new sheets
    For lngCurr = 1 To lngLevels
        'create new worksheet for the level
        Set wsNew = wb.Sheets.Add(After:=wb.Sheets(ws.Index + lngCurr - 1))
        wsNew.Name = "L" & Format$(lngCurr, "00")

        'get row numbers
        lngBegRow = colLRows(lngCurr)
        If lngCurr < lngLevels Then
            lngEndRow = colLRows(lngCurr + 1) - 1
        Else
            lngEndRow = lngLastRow
        End If

        '____________________COPY DATA________________________________________________
        'copy rows to new worksheet
        ws.Rows(lngBegRow & ":" & lngEndRow).Copy Destination:=wsNew.Rows(1)
        CopyColumnWidths ws, wsNew
        wsNew.Calculate
        RelinkShiftedFormulas ws, wsNew, lngBegRow

        '____________________COLUMN SETUP WITHIN SHEETS_______________________________
        ' link answer cells to new cells in new sheet and mark input cells
        ' get column numbers for answers and inputs

        ' Find the header row with "Answer", "Level", or "Points"
        Set rngHdrRow = Nothing
        For Each rngRow In wsNew.Rows(1 & ":" & lngEndRow - lngBegRow + 1)
            Set rngHdrRow = rngRow.Find(What:="Answer", LookIn:=xlValues, LookAt:=xlWhole)
            If Not rngHdrRow Is Nothing Then Exit For
            Set rngHdrRow = rngRow.Find(What:="Level", LookIn:=xlValues, LookAt:=xlWhole)
            If Not rngHdrRow Is Nothing Then Exit For
            Set rngHdrRow = rngRow.Find(What:="Points", LookIn:=xlValues, LookAt:=xlWhole)
            If Not rngHdrRow Is Nothing Then Exit For
        Next rngRow

        ' No header row found (e.g. a one-row block from a contents list): no input
        ' columns can be identified, so input marking is skipped for this sheet.
        booHdrs = Not (rngHdrRow Is Nothing)

        ' Identify non-matching columns after column B
        Set colColList = New Collection
        lngACol = 5     'default value of 5 unless overwritten
        lngEndCol = 0
        If booHdrs Then
            For lngCol = 3 To wsNew.UsedRange.Columns.Count + 2 ' Start after column B
                strHeader = SafeText(wsNew.Cells(rngHdrRow.Row, lngCol).Value2)
                If strHeader = "Answer" Then lngACol = lngCol
                If Not excludedHeaders.Exists(strHeader) And strHeader <> "" Then
                    colColList.Add lngCol
                    If lngCol > lngEndCol Then
                        lngEndCol = lngCol
                    End If
                End If
            Next lngCol
        End If
        If lngEndCol = 0 Then lngEndCol = wsNew.UsedRange.Columns.Count + 2

        'loop through rows in level sheet and mark
        For lngRow = lngBegRow To lngEndRow
            'set answer lookups in Case sheet
            Set rngNumbers = ws.Cells(lngRow, 2)
            Set rngAnswers = ws.Cells(lngRow, lngACol)
            If IsError(wsNew.Cells(lngRow - lngBegRow + 1, lngACol).Value2) Then
                wsNew.Cells(lngRow - lngBegRow + 1, lngACol).Formula = ""
            End If

            blnLink = IsError(rngAnswers.Value2)
            If Not blnLink Then
                blnLink = (Not IsEmpty(rngNumbers.Value2) And IsNumeric(rngNumbers.Value2) And IsEmpty(rngAnswers.Value2))
            End If
            If Not blnLink Then
                blnLink = (InStr(rngAnswers.Formula, "#REF") > 0)
            End If
            If blnLink Then
                rngAnswers.Formula = "='" & wsNew.Name & "'!" & rngAnswers.Offset(1 - lngBegRow).Address
            End If

            'set score lookups - assumed this is the column after "Answer"
            'probe: a merged or protected cell here must not stop the level being built
            Set rngScoreCell = wsNew.Cells(lngRow - lngBegRow + 1, lngACol + 1)
            On Error Resume Next
            If IsError(rngScoreCell.Value2) Then
                rngScoreCell.Formula = "='" & ws.Name & "'!" & rngAnswers.Offset(0, 1).Address
            ElseIf Len(CStr(rngScoreCell.Value2)) > 0 Then
                rngScoreCell.Formula = "='" & ws.Name & "'!" & rngAnswers.Offset(0, 1).Address
            End If
            On Error GoTo ErrHandler

            'mark as input cells for non-standard columns with values
            If booHdrs Then
                If lngRow - lngBegRow + 1 > rngHdrRow.Row Then
                    For Each varCol In colColList
                        Set rngCell = wsNew.Cells(lngRow - lngBegRow + 1, varCol)
                        If Not IsEmpty(rngCell.Value2) And rngCell.HasFormula = False Then
                            Call MarkAsInputCells(rngCell, False)
                        End If
                    Next varCol
                End If
            End If

        Next lngRow

        'for workbooks with answers (for training), add the score at the top of each sheet
        If Not rngScore Is Nothing Then
            wsNew.Cells(1, 1).Formula = "='" & ws.Name & "'!" & rngScore.Address
        End If

        'clear shapes
        For Each shp In wsNew.Shapes
            shp.Delete
        Next shp

    Next lngCurr

Cleanup:
    On Error Resume Next
    If blnInit Then VBAFin
    If Len(strErr) > 0 Then
        m_strStepError = strErr
        Application.StatusBar = Left$(PROC_NAME & " failed - " & strErr, m_MAX_STATUS)
    Else
        Application.Calculate
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If Not wsNew Is Nothing Then strErr = strErr & " (at level " & lngCurr & " of " & lngLevels & ")"
    If LogError(PROC_NAME, Err.Number, strErr) Then m_blnLogWritten = True
    Resume Cleanup
End Sub

' Purpose: After a level block is copied to its level sheet, relink any formula whose value no
'          longer matches the case cell it came from. Copying shifts relative references with the
'          rows, so an input such as ='Wally (L5)'!X5 on Task row 197 became ='Wally (L5)'!#REF!
'          on L05 (2022 MEWC qualification, 2026-09-25), and a smaller shift would land on a valid
'          but WRONG cell with no error to show it. Such a cell becomes a link to the case cell
'          (with # when that cell spills). Formulas that still agree (game numbers =B30+1, Level
'          Total SUMs) are left alone. Never stops the level being built.
Private Sub RelinkShiftedFormulas(ByVal wsCase As Worksheet, ByVal wsLevel As Worksheet, _
                                  ByVal lngBegRow As Long)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "RelinkShiftedFormulas"

    Dim rngFormulas As Range
    Dim rngCell     As Range
    Dim rngOrig     As Range
    Dim varNew      As Variant
    Dim varOld      As Variant
    Dim blnSame     As Boolean
    Dim blnChild    As Boolean
    Dim strLink     As String

    On Error GoTo ErrHandler

    ' Probe: SpecialCells raises when the sheet holds no formulas.
    On Error Resume Next
    Set rngFormulas = wsLevel.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo ErrHandler
    If rngFormulas Is Nothing Then GoTo Cleanup

    For Each rngCell In rngFormulas.Cells
        ' Probe: SpillParent raises outside a spill. A spilled child is never rewritten.
        blnChild = False
        On Error Resume Next
        blnChild = (rngCell.SpillParent.Address <> rngCell.Address)
        On Error GoTo ErrHandler

        If Not blnChild Then
            Set rngOrig = wsCase.Cells(rngCell.Row + lngBegRow - 1, rngCell.Column)
            varNew = rngCell.Value2
            varOld = rngOrig.Value2
            If IsError(varNew) Or IsError(varOld) Then
                blnSame = IsError(varNew) And IsError(varOld)
                If blnSame Then blnSame = (CStr(varNew) = CStr(varOld))
            Else
                blnSame = (varNew = varOld)
            End If

            If Not blnSame Then
                strLink = "='" & Replace(wsCase.Name, "'", "''") & "'!" & rngOrig.Address(True, True)
                If rngOrig.HasSpill Then strLink = strLink & "#"
                rngCell.Formula2 = strLink
            End If
        End If
    Next rngCell

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If LogError(PROC_NAME, Err.Number, Err.Description) Then m_blnLogWritten = True
    Resume Cleanup
End Sub

'creates a bonus sheet with just the bonus info on
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Bonus Sheet
' Description:            Creates bonus sheet "B" with bonus questions
' Macro Expression:       modCaseSetup.CreateBonusSheet()
' Generated:              01/08/2025 04:37 PM
'----------------------------------------------------------------------------------------------------
Public Sub CreateBonusSheet()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "CreateBonusSheet"

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim lngRow As Long
    Dim lngBegRow As Long
    Dim lngEndRow As Long
    Dim lngLastRow As Long
    Dim rngColB As Range
    Dim rngAnswers As Range
    Dim rngNumbers As Range
    Dim rngScoreCell As Range
    Dim objCell As Range
    Dim rngScore As Range
    Dim lngACol As Long
    Dim booHdrs As Boolean
    Dim blnLink As Boolean
    Dim lngCol As Long
    Dim rngHdrRow As Range
    Dim rngRow As Range
    Dim strHeader As String
    Dim strCurrVal As String
    Dim strBonusKey As String
    Dim blnInit As Boolean
    Dim strErr As String

    On Error GoTo ErrHandler

    m_strStepNote = vbNullString

    'Initialize
    VBAInit
    blnInit = True
    'set up workbook and worksheet
    Set wb = ActiveWorkbook
    Set ws = GetCaseSheet(wb)
    If ws Is Nothing Then Err.Raise m_ERR_CASE_SHEET, PROC_NAME, "Case sheet not found"

    '___________________worksheet exists_____________________
    'get current score cell - probe, it is only there in training files
    On Error Resume Next
    Set rngScore = ws.UsedRange.Find(What:="Current Score", LookIn:=xlValues, LookAt:=xlPart, MatchCase:=False).Offset(1, 0)
    On Error GoTo ErrHandler

    'cycle through column B looking for "Bonus Questions" cell
    'cycle through column B of case sheet to find first level
    lngLastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Set rngColB = ws.Range(ws.Cells(1, 2), ws.Cells(lngLastRow, 2))
    ' Start = the "Bonus Questions" cell. End = the row before the next "Questions" /
    ' "Levels" / "Level #" cell after it, or the last row of column B.
    lngBegRow = 0
    lngEndRow = 0
    For Each objCell In rngColB.Cells
        strCurrVal = Trim$(SafeText(objCell.Value2))
        If Len(strCurrVal) > 0 Then
            If lngBegRow = 0 Then
                ' Was an exact match on "Bonus Questions" alone. The 2026 UK International
                ' Women's Day case heads its bonus block "Bonuses" and so lost all five of its
                ' bonuses - 250 points, the largest single block in that case - silently
                ' (2026-09-23). Now any column-B cell STARTING with "Bonus" also counts,
                ' provided its row carries nothing else. That proviso is what separates a
                ' section HEADING from the two things that also start with "Bonus": the
                ' Contents-table row (difficulty, points and title beside it) and the bonus
                ' rows themselves ("Bonus 1", with level, points and the question beside them).
                ' The exact phrase still matches unconditionally, so no layout that worked
                ' before can stop working now.
                If strCurrVal = "Bonus Questions" Then
                    lngBegRow = objCell.Row
                ElseIf StrComp(Left$(strCurrVal, 5), "Bonus", vbTextCompare) = 0 Then
                    If Application.WorksheetFunction.CountA(objCell.EntireRow) <= 1 Then
                        lngBegRow = objCell.Row
                    End If
                End If
            ElseIf strCurrVal = "Questions" Or strCurrVal = "Levels" Or _
            (strCurrVal Like "Level *" And strCurrVal <> "Level Code" And _
            IsNumeric(Trim(Replace(strCurrVal, "Level", "")))) Then
                lngEndRow = objCell.Row - 1
                Exit For
            End If
        End If
    Next objCell

    ' No bonus block is a normal layout, not a fault: skip, and say so.
    If lngBegRow = 0 Then
        m_strStepNote = "no 'Bonus Questions' cell in column B"
        GoTo Cleanup
    End If
    If lngEndRow = 0 Then lngEndRow = lngLastRow

    If SheetNameTaken(wb, "B") Then Err.Raise m_ERR_NOT_FOUND, PROC_NAME, "Sheet B already exists - bonus sheet not created"

    'create new worksheet for bonuses
    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(ws.Index))
    wsNew.Name = "B"

    '____________________COPY DATA________________________________________________
    'copy rows to new worksheet
    ws.Rows(lngBegRow & ":" & lngEndRow).Copy Destination:=wsNew.Rows(1)
    CopyColumnWidths ws, wsNew
    wsNew.Calculate

    '____________________COLUMN SETUP WITHIN SHEETS_______________________________
    ' link answer cells to new cells in new sheet
    ' get column numbers for answers

    ' Find the header row with "Answer", "Level", or "Points"
    For Each rngRow In wsNew.Rows(1 & ":" & lngEndRow - lngBegRow + 1)
        Set rngHdrRow = rngRow.Find(What:="Answer", LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngHdrRow Is Nothing Then Exit For
        Set rngHdrRow = rngRow.Find(What:="Level", LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngHdrRow Is Nothing Then Exit For
        Set rngHdrRow = rngRow.Find(What:="Points", LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngHdrRow Is Nothing Then Exit For
    Next rngRow

    booHdrs = Not (rngHdrRow Is Nothing)

    'identify answer column
    lngACol = 5     'default value of 5 unless overwritten
    If booHdrs Then
        For lngCol = 3 To wsNew.UsedRange.Columns.Count + 2 ' Start after column B
            strHeader = SafeText(wsNew.Cells(rngHdrRow.Row, lngCol).Value2)
            If strHeader = "Answer" Then lngACol = lngCol
        Next lngCol
    End If

    'loop through rows in bonus sheet and link
    For lngRow = lngBegRow To lngEndRow
        'set answer lookups in Case sheet
        Set rngNumbers = ws.Cells(lngRow, 2)
        Set rngAnswers = ws.Cells(lngRow, lngACol)
        If IsError(wsNew.Cells(lngRow - lngBegRow + 1, lngACol).Value2) Then
            wsNew.Cells(lngRow - lngBegRow + 1, lngACol).Formula = ""
        End If

        blnLink = IsError(rngAnswers.Value2)
        If Not blnLink Then
            ' "Bonus 1" or, since 2026-09-25, "Bonus A" (2025 MEWC qualification rounds): a single
            ' letter after "Bonus" counts as a bonus row too.
            strBonusKey = Trim(Replace(SafeText(rngNumbers.Value2), "Bonus", ""))
            blnLink = (Not IsEmpty(rngNumbers.Value2) And _
                       (IsNumeric(strBonusKey) Or strBonusKey Like "[A-Za-z]") And _
                       IsEmpty(rngAnswers.Value2))
        End If
        If Not blnLink Then
            blnLink = (InStr(rngAnswers.Formula, "#REF") > 0)
        End If
        If blnLink Then
            rngAnswers.Formula = "='" & wsNew.Name & "'!" & rngAnswers.Offset(1 - lngBegRow).Address
        End If

        'set score lookups - assumed this is the column after "Answer"
        'probe: a merged or protected cell here must not stop the sheet being built
        Set rngScoreCell = wsNew.Cells(lngRow - lngBegRow + 1, lngACol + 1)
        On Error Resume Next
        If IsError(rngScoreCell.Value2) Then
            rngScoreCell.Formula = "='" & ws.Name & "'!" & rngAnswers.Offset(0, 1).Address
        ElseIf Len(CStr(rngScoreCell.Value2)) > 0 Then
            rngScoreCell.Formula = "='" & ws.Name & "'!" & rngAnswers.Offset(0, 1).Address
        End If
        On Error GoTo ErrHandler

    Next lngRow

    'for workbooks with answers (for training), add the score at the top of each sheet
    If Not rngScore Is Nothing Then
        wsNew.Cells(1, 1).Formula = "='" & ws.Name & "'!" & rngScore.Address
    End If

Cleanup:
    On Error Resume Next
    If blnInit Then VBAFin
    If Len(strErr) > 0 Then
        m_strStepError = strErr
        Application.StatusBar = Left$(PROC_NAME & " failed - " & strErr, m_MAX_STATUS)
    ElseIf Len(m_strStepNote) > 0 Then
        Application.StatusBar = PROC_NAME & " skipped - " & m_strStepNote
    Else
        Application.Calculate
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, Err.Description) Then m_blnLogWritten = True
    Resume Cleanup
End Sub

' Purpose: Build a data table block at rngTarget for the level in view. Returns the block's input
'          cell, or Nothing on failure (the reason goes to the StatusBar, and to the log for errors).
'          Nothing is left half-built: if an error strikes after writing started, exactly the cells
'          written are cleared again, and answer cells already re-linked get their old formulas back.
'          Level in view:
'            - rngExample, when the caller passes the example-label cell (Create Data Table does);
'            - else on an L## sheet: the example row in column B (a row reading "Example3"/"Example 3"
'              with an Answer header two rows above);
'            - else the active cell, which must hold the "Example#" label.
'          Games: LevelGameRows - the numbers below the example row in the same column, stopping at the
'          next section (another example, a level heading, "Game #", "Level Code", a Bonus heading) or
'          at the first non-number after the games, so bonus questions or a reference list below are
'          never taken for games.
Public Function BuildDataTable(ByVal rngTarget As Range, Optional ByVal rngExample As Range = Nothing) As Range

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "BuildDataTable"
    Const MAX_SCAN_ROWS As Long = 1000

    Dim rngAC As Range
    Dim ws As Worksheet
    Dim strCV As String
    Dim lngQs() As Long
    Dim lngQRows() As Long      ' sheet row of each game, for linking its answer cell
    Dim strLabel As String      ' the example row's own label text, e.g. "Example3" or "Example 3"
    Dim blnLevelSheet As Boolean
    Dim lngGameCol As Long
    Dim lngStopRow As Long
    Dim lngRow As Long
    Dim lngQCnt As Long
    Dim lngCurrQ As Long
    Dim lngCol As Long
    Dim strCol As String
    Dim strGameCol As String
    Dim lngExRow As Long
    Dim lngHdRow As Long
    Dim lngACol As Long
    Dim lngInCnt As Long
    Dim colIn As Collection
    Dim varCol As Variant
    Dim blnInit As Boolean
    Dim blnWriting As Boolean
    Dim strMsg As String
    Dim strErr As String
    Dim rngExpected As Range
    Dim rngCheck As Range
    Dim strIn As String
    Dim strOut As String
    Dim strExp As String
    Dim objCond As Object
    Dim colGames As Collection
    Dim varGameRow As Variant
    Dim rngWritten As Range     ' every cell written, so a failure clears exactly those
    Dim varOldAns() As Variant  ' answer-cell formulas before linking, restored on failure
    Dim blnLinking As Boolean

    On Error GoTo ErrHandler

    VBAInit
    blnInit = True

    Set ws = rngTarget.Parent
    If rngExample Is Nothing Then
        If Not ActiveSheet Is ws Then
            strMsg = "Create Data Table: the target cell is not on the active sheet."
            GoTo Cleanup
        End If
        Set rngAC = ActiveCell
    Else
        Set rngAC = rngExample
    End If

    'make sure target row is row 2 or higher
    If rngTarget.Row = 1 Then Set rngTarget = ws.Cells(2, rngTarget.Column)

    ' 1. The example row, and the column the game numbers are in.
    blnLevelSheet = IsLevelSheetName(ws.Name) And (rngExample Is Nothing)
    If blnLevelSheet Then
        lngGameCol = 2
        lngStopRow = ws.Cells(ws.Rows.Count, lngGameCol).End(xlUp).Row
        lngExRow = FindExampleRow(ws, 1, lngStopRow)
        If lngExRow = 0 Then
            strMsg = "Create Data Table: no 'Example#' row found in column B of " & ws.Name & "."
            GoTo Cleanup
        End If
    Else
        If Not IsExampleLabel(rngAC.Value2) Then
            strMsg = "Create Data Table: select an 'Example#' cell or use a level sheet named 'L##'."
            GoTo Cleanup
        End If
        lngGameCol = rngAC.Column
        lngExRow = rngAC.Row
        lngStopRow = ws.Cells(ws.Rows.Count, lngGameCol).End(xlUp).Row
    End If
    If lngStopRow > lngExRow + MAX_SCAN_ROWS Then lngStopRow = lngExRow + MAX_SCAN_ROWS
    strLabel = SafeText(ws.Cells(lngExRow, lngGameCol).Value2)
    strGameCol = Replace(Replace(ws.Cells(1, lngGameCol).Address, "1", ""), "$", "")

    ' 2. Game numbers below the example row (the shared rule Create Data Table also uses).
    Set colGames = LevelGameRows(ws, lngGameCol, lngExRow, lngStopRow)
    lngQCnt = 0
    For Each varGameRow In colGames
        lngQCnt = lngQCnt + 1
        ReDim Preserve lngQs(1 To lngQCnt)
        ReDim Preserve lngQRows(1 To lngQCnt)
        lngQs(lngQCnt) = CLng(ws.Cells(varGameRow, lngGameCol).Value2)
        lngQRows(lngQCnt) = varGameRow
    Next varGameRow

    If lngQCnt = 0 Then
        strMsg = "Create Data Table: no game numbers found below " & ws.Cells(lngExRow, lngGameCol).Address(False, False) & "."
        GoTo Cleanup
    End If

    lngHdRow = HeaderRowFor(ws, lngExRow)
    If lngHdRow < 1 Then
        strMsg = "Create Data Table: no header row above the example row."
        GoTo Cleanup
    End If

    ' Refuse before writing anything if an old What-If table sits where the results would go -
    ' Excel will not overwrite part of one.
    If HasTableFormulas(ws.Range(ws.Cells(rngTarget.Row + 4, rngTarget.Column + 1), _
                                 ws.Cells(rngTarget.Row + 3 + lngQCnt, rngTarget.Column + 1))) Then
        strMsg = "Create Data Table: an old What-If data table is at " & _
                 ws.Cells(rngTarget.Row + 4, rngTarget.Column + 1).Address(False, False) & " - clear it first."
        GoTo Cleanup
    End If

    ' The inputs, and the Answer column: the one rule, shared with Create Data Table's placement.
    Set colIn = LevelInputColumns(ws, lngGameCol, lngExRow, lngACol)
    If colIn Is Nothing Then
        strMsg = "Create Data Table: could not read the inputs on the example row - see " & LogLocation()
        GoTo Cleanup
    End If

    ' 3. Write the block. From here on a failure clears what was written.
    blnWriting = True
    ' The input cell holds the example row's exact text, so the XLOOKUPs find it whichever way the
    ' case writes it.
    Set rngWritten = rngTarget
    rngTarget.value = strLabel
    For Each varCol In colIn
        lngCol = varCol
        strCol = Replace(Replace(ws.Cells(1, lngCol).Address, "1", ""), "$", "")
        lngInCnt = lngInCnt + 1
        Set rngWritten = AddToRange(rngWritten, ws.Cells(rngTarget.Row - 1, rngTarget.Column + lngInCnt))
        Set rngWritten = AddToRange(rngWritten, ws.Cells(rngTarget.Row, rngTarget.Column + lngInCnt))
        ' header, then the example's own cell (its formatting) - Range.Copy with a
        ' Destination does not touch the clipboard
        ws.Cells(lngHdRow, lngCol).Copy Destination:=ws.Cells(rngTarget.Row - 1, rngTarget.Column + lngInCnt)
        ws.Cells(lngExRow, lngCol).Copy Destination:=ws.Cells(rngTarget.Row, rngTarget.Column + lngInCnt)
        'lookup data
        ws.Cells(rngTarget.Row, rngTarget.Column + lngInCnt).Formula = "=XLOOKUP(" & _
            rngTarget.Address(1, 1) & "," & strGameCol & ":" & strGameCol & "," & strCol & ":" & strCol & ",,0)"
    Next varCol

    'set up formula cell (and the Expected / Check label and value cells below the input)
    Set rngWritten = AddToRange(rngWritten, ws.Range(ws.Cells(rngTarget.Row + 1, rngTarget.Column), _
                                                     ws.Cells(rngTarget.Row + 2, rngTarget.Column + 1)))
    Set rngWritten = AddToRange(rngWritten, ws.Cells(rngTarget.Row + 3, rngTarget.Column + 1))
    ws.Cells(rngTarget.Row + 3, rngTarget.Column + 1).value = "ANSWER"
    ws.Cells(rngTarget.Row + 3, rngTarget.Column + 1).Interior.ColorIndex = 4

    'example check, in the two free rows just above the table (Jaq, 2026-09-14):
    '   Expected | the example's own answer from the Answer column
    '   Check    | MATCH / DIFFERS between ANSWER and Expected (or which game is loaded)
    Set rngExpected = ws.Cells(rngTarget.Row + 1, rngTarget.Column + 1)
    Set rngCheck = ws.Cells(rngTarget.Row + 2, rngTarget.Column + 1)
    ws.Cells(rngTarget.Row + 1, rngTarget.Column).value = "Expected"
    ws.Cells(rngTarget.Row + 2, rngTarget.Column).value = "Check"
    If lngACol > 2 Then
        ' A blank example answer must stay blank - a bare link would turn it into 0 and give a
        ' false MATCH / DIFFERS.
        strExp = ws.Cells(lngExRow, lngACol).Address(True, True)
        rngExpected.Formula = "=IF(ISBLANK(" & strExp & "),""""," & strExp & ")"
    End If
    strIn = rngTarget.Address(True, True)
    strOut = ws.Cells(rngTarget.Row + 3, rngTarget.Column + 1).Address(False, False)
    strExp = rngExpected.Address(False, False)
    ' Errors are caught first (a half-built solve often returns #N/A); numbers compare within a
    ' small tolerance; text compares as text.
    rngCheck.Formula = "=IF(ISNUMBER(" & strIn & "),""game ""&" & strIn & "&"" loaded""," & _
                       "IF(ISERROR(" & strOut & "),""output is an error""," & _
                       "IF(ISERROR(" & strExp & "),""expected is an error""," & _
                       "IF(" & strExp & "="""",""no expected answer""," & _
                       "IF(" & strOut & "=""ANSWER"",""point ANSWER at output""," & _
                       "IF(IFERROR(ABS(" & strOut & "-" & strExp & ")<0.000001," & strOut & "&""""=" & strExp & "&""""),""MATCH"",""DIFFERS""))))))"
    rngCheck.FormatConditions.Delete
    Set objCond = rngCheck.FormatConditions.Add(Type:=xlTextString, String:="is an error", TextOperator:=xlContains)
    objCond.Interior.Color = RGB(255, 199, 206)
    objCond.Font.Color = RGB(156, 0, 6)
    Set objCond = rngCheck.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""MATCH""")
    objCond.Interior.Color = RGB(198, 239, 206)
    objCond.Font.Color = RGB(0, 97, 0)
    Set objCond = rngCheck.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""DIFFERS""")
    objCond.Interior.Color = RGB(255, 199, 206)
    objCond.Font.Color = RGB(156, 0, 6)

    'game numbers and results
    Set rngWritten = AddToRange(rngWritten, ws.Range(ws.Cells(rngTarget.Row + 4, rngTarget.Column), _
                                                     ws.Cells(rngTarget.Row + 3 + lngQCnt, rngTarget.Column + 1)))
    For lngCurrQ = 1 To lngQCnt
        ws.Cells(rngTarget.Row + 3 + lngCurrQ, rngTarget.Column).value = lngQs(lngCurrQ)
    Next lngCurrQ

    'results column: #N/A until Run Data Table (RDT) fills it with static values.
    'No What-If data table - Excel's TABLE() recalculates on every change and breaks down as the
    'solve formula grows (Jaq, 2026-09-14).
    ws.Range(ws.Cells(rngTarget.Row + 4, rngTarget.Column + 1), _
             ws.Cells(rngTarget.Row + 3 + lngQCnt, rngTarget.Column + 1)).Value2 = CVErr(xlErrNA)

    'points answers at table - on a level sheet or straight on the Case sheet, using the row each
    'game was found on
    If lngACol > 2 Then
        ' Keep the old answer formulas first, so a failure part-way can put them back.
        ReDim varOldAns(1 To lngQCnt)
        For lngCurrQ = 1 To lngQCnt
            varOldAns(lngCurrQ) = ws.Cells(lngQRows(lngCurrQ), lngACol).Formula
        Next lngCurrQ
        blnLinking = True
        For lngCurrQ = 1 To lngQCnt
            ws.Cells(lngQRows(lngCurrQ), lngACol).Formula = "=" & _
                ws.Cells(rngTarget.Row + 3 + lngCurrQ, rngTarget.Column + 1).Address(0, 0)
        Next lngCurrQ
        blnLinking = False
    End If

    blnWriting = False
    Set BuildDataTable = rngTarget
    strMsg = "Create Data Table: block ready at " & rngTarget.Address(False, False) & _
             " - point ANSWER at your output, then run Run Data Table (RDT)."

Cleanup:
    On Error Resume Next
    If Len(strErr) > 0 And blnWriting Then
        ' Undo the half-built block: exactly the cells written, and any answer links already changed.
        If blnLinking Then
            For lngCurrQ = 1 To lngQCnt
                ws.Cells(lngQRows(lngCurrQ), lngACol).Formula = varOldAns(lngCurrQ)
            Next lngCurrQ
        End If
        If Not rngWritten Is Nothing Then rngWritten.Clear
        Set BuildDataTable = Nothing
    End If
    If blnInit Then VBAFin
    If Len(strErr) > 0 Then
        Application.StatusBar = Left$(PROC_NAME & " failed - " & strErr, m_MAX_STATUS)
    ElseIf Len(strMsg) > 0 Then
        Application.StatusBar = Left$(strMsg, m_MAX_STATUS)
    End If
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, Err.Description) Then
        m_blnLogWritten = True
        strErr = strErr & " - see " & LogLocation()
    End If
    Resume Cleanup
End Function

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Case Inputs Sheet
' Description:            Create case inputs sheet listing columns B onwards for question rows in "Case" sheet
' Macro Expression:       modCaseSetup.CreateCaseInputsSheet()
' Generated:              12/03/2024 10:04 AM
'----------------------------------------------------------------------------------------------------
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Case Inputs Sheet
' Description:            Creates a case inputs sheet with the inputs for the current case
' Macro Expression:       modCaseSetup.CreateCaseInputsSheet()
' Generated:              01/03/2025 08:10 PM
'----------------------------------------------------------------------------------------------------
' inputs cell value of active cell when command is run
'if cell is blank, regular method
Public Sub CreateCaseInputsSheet(Optional ByVal strDetailed As String = "")

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "CreateCaseInputsSheet"

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim wsNew As Worksheet
    Dim booDetailed As Boolean
    Dim blnInit As Boolean
    Dim strErr As String

    On Error GoTo ErrHandler

    VBAInit
    blnInit = True

    Set wb = ActiveWorkbook
    Set ws = GetCaseSheet(wb)
    If ws Is Nothing Then Err.Raise m_ERR_CASE_SHEET, PROC_NAME, "Case sheet not found"
    booDetailed = (Len(strDetailed) > 0)
    If SheetNameTaken(wb, "CaseInputs") Then Err.Raise m_ERR_NOT_FOUND, PROC_NAME, "Sheet CaseInputs already exists - inputs sheet not created"

    'worksheet exists
    Set wsNew = wb.Sheets.Add(After:=wb.Sheets(ws.Index))
    wsNew.Name = "CaseInputs"
    'freeze panes below header
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 2
        .SplitRow = 2
        .FreezePanes = True
    End With
    If Not booDetailed Then
        wsNew.Cells(2, 2).Formula2 = "=CaseInputs()"
    Else
        Call DetailedInputs(ws, wsNew)
    End If

Cleanup:
    On Error Resume Next
    If blnInit Then VBAFin
    If Len(strErr) > 0 Then
        m_strStepError = strErr
        Application.StatusBar = Left$(PROC_NAME & " failed - " & strErr, m_MAX_STATUS)
    End If
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strErr = Err.Number & ": " & Err.Description
    If LogError(PROC_NAME, Err.Number, Err.Description) Then m_blnLogWritten = True
    Resume Cleanup
End Sub

'more detailed breakdown of inputs for non-standard layouts
'No handler of its own by design: it is only called from CreateCaseInputsSheet, and an error
'here must propagate to that caller's handler so the step is reported as failed.
Private Sub DetailedInputs(ByVal ws As Worksheet, ByVal wsNew As Worksheet)

    ' --- CONSTANTS (local to function) ---
    Const HDR_OUT_ROW   As Long = 2     ' headers written from column C on this row
    Const FIRST_OUT_ROW As Long = 3     ' game rows from here, game # in column B

    Dim rngLast As Range
    Dim varData As Variant
    Dim varFormulas As Variant
    Dim varOut() As Variant
    Dim varHdrOut() As Variant
    Dim lngIncCols() As Long
    Dim strHeaders() As String
    Dim lngInc As Long
    Dim lngEndRow As Long
    Dim lngEndCol As Long
    Dim lngHdrRow As Long
    Dim lngHdr As Long
    Dim lngRowOut As Long
    Dim lngRow As Long
    Dim lngCol As Long
    Dim lngExCont As Long
    Dim lngFirstCol As Long
    Dim lngContLeft As Long
    Dim strCV As String
    Dim strHdr As String
    Dim strHdrCV As String
    Dim strGame As String
    Dim booHdr As Boolean
    Dim booSkip As Boolean
    Dim lngHdrAnsCol As Long
    Dim lngScan As Long

    ' Last row: column B's last row, or the last cell with content if that is lower (the final
    ' game's extra rows can sit below the last game number). Find, not UsedRange, which
    ' over-reports when formatting runs past the data.
    lngEndRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Set rngLast = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not rngLast Is Nothing Then
        If rngLast.Row > lngEndRow Then lngEndRow = rngLast.Row
    End If
    ' Last column the same way (+2 as before) - UsedRange stretches to any stray formatted cell.
    Set rngLast = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If rngLast Is Nothing Then
        lngEndCol = 3
    Else
        lngEndCol = rngLast.Column + 2
    End If
    If lngEndCol < 3 Then lngEndCol = 3
    If lngEndCol > ws.Columns.Count Then lngEndCol = ws.Columns.Count

    ' One read of the whole block. .Value, not .Value2, on purpose: this sheet is a human-readable
    ' view of the inputs, so a date should read as a date rather than a serial number.
    varData = ws.Range(ws.Cells(1, 1), ws.Cells(lngEndRow, lngEndCol)).value
    ReDim varOut(1 To lngEndRow, 1 To lngEndCol)
    ReDim varHdrOut(1 To 1, 1 To lngEndCol)
    lngInc = 0
    lngRowOut = 0

    ' Cycle through the rows in memory.
    For lngRow = 1 To lngEndRow
        strCV = SafeText(varData(lngRow, 2))
        ' Was Left$(strCV, 7) = "Example" - case-SENSITIVE, and English only. Now goes through the
        ' shared prefix list. The Len cap stays: it is what stops a long prose sentence beginning
        ' "Example 1: The requested character is..." being taken for a label.
        ' Trimmed (ordinary AND non-breaking spaces) before testing: this site reads the left of
        ' the raw string, so a LEADING space used to make " Example3" invisible here. Never seen in
        ' the 152-case library, but it costs nothing to be safe. The Len cap uses the trimmed length
        ' too, or a padded cell would slip past the guard that stops prose matching.
        If ExamplePrefixLen(Trim$(Replace(strCV, Chr$(160), " "))) > 0 _
           And Len(Trim$(Replace(strCV, Chr$(160), " "))) <= 10 Then
            ' Example row: rebuild this level's output-column -> source-column map by header name,
            ' so levels whose input columns are reordered still line up.
            lngContLeft = 0
            ' Header row = nearest non-blank cell in column B above the example, STEPPING OVER any
            ' further example labels on the way up. A level with two worked rows (Example3a and
            ' Example3b) has them adjacent, with the spacer below the pair rather than between
            ' them, so the plain walk stopped on the FIRST example row and took its DATA VALUES
            ' as the header names. Measured 2026-09-23 on the 2026 UK International Women's Day
            ' case, levels 3 and 7: the level's real input column was left empty and its data
            ' landed in junk columns headed "GB144587/31758" and "Possenhofen Castle", while the
            ' Level/Points/check columns - no longer excluded, because their headers now read
            ' "3", "0", "6" - were pulled in as inputs. Silent: no error, nothing in the log.
            ' Same test as the example-row test above, so the two cannot drift apart.
            lngHdrRow = lngRow - 1
            Do While lngHdrRow > 1
                strHdrCV = Trim$(Replace(SafeText(varData(lngHdrRow, 2)), Chr$(160), " "))
                If Len(strHdrCV) > 0 Then
                    If Not (ExamplePrefixLen(strHdrCV) > 0 And Len(strHdrCV) <= 10) Then Exit Do
                End If
                lngHdrRow = lngHdrRow - 1
            Loop
            If lngHdrRow < 1 Then lngHdrRow = 1
            ' Where the preamble ends: the Answer column in THIS level's header row. Everything
            ' from the game number across to it is the fixed Game/Level/Points/Answer block;
            ' anything right of it is an input, whatever it happens to be called.
            lngHdrAnsCol = 0
            For lngScan = 3 To lngEndCol
                If SafeText(varData(lngHdrRow, lngScan)) = "Answer" Then
                    lngHdrAnsCol = lngScan
                    Exit For
                End If
            Next lngScan
            ' reset current column numbers
            If lngInc > 0 Then ReDim lngIncCols(1 To lngInc)
            ' formulas for this one row only - example inputs that are formulas are not inputs
            varFormulas = ws.Range(ws.Cells(lngRow, 1), ws.Cells(lngRow, lngEndCol)).Formula
            For lngCol = 3 To lngEndCol
                strHdr = SafeText(varData(lngHdrRow, lngCol))
                strCV = SafeText(varData(lngRow, lngCol))
                ' Was a plain name test - ANY column headed Level/Points/Answer was excluded,
                ' wherever it sat. Airplane Battleship's level 2 takes the MAP LEVEL as a real
                ' per-game input and heads that column "Level", so the input was dropped and the
                ' level arrived in CaseInputs with half its data (2026-09-23). Its level 5 asks
                ' for the same quantity under the header "Map Level" and survived - the same
                ' input kept or lost on the author's choice of word. Now only the REAL preamble
                ' columns are excluded: those at or left of this level's Answer column.
                booSkip = False
                If strHdr = "Level" Or strHdr = "Points" Or strHdr = "Answer" Then
                    If lngHdrAnsCol = 0 Or lngCol <= lngHdrAnsCol Then booSkip = True
                End If
                If strCV <> "" And Left$(SafeText(varFormulas(1, lngCol)), 1) <> "=" _
                   And Not booSkip Then
                    ' find the current header match and update the column to output from
                    booHdr = False
                    For lngHdr = 1 To lngInc
                        If strHeaders(lngHdr) = strHdr And lngIncCols(lngHdr) = 0 Then
                            lngIncCols(lngHdr) = lngCol
                            booHdr = True
                            Exit For
                        End If
                    Next lngHdr
                    If Not booHdr Then
                        lngInc = lngInc + 1
                        ' headers accumulate across levels, so the output can be wider than the
                        ' source - grow the output arrays only when a new slot is needed
                        If lngInc + 1 > UBound(varOut, 2) Then
                            ReDim Preserve varOut(1 To lngEndRow, 1 To lngInc + 1)
                            ReDim Preserve varHdrOut(1 To 1, 1 To lngInc + 1)
                        End If
                        If lngInc = 1 Then
                            ReDim strHeaders(1 To 1)
                            ReDim lngIncCols(1 To 1)
                        Else
                            ReDim Preserve strHeaders(1 To lngInc)
                            ReDim Preserve lngIncCols(1 To lngInc)
                        End If
                        strHeaders(lngInc) = strHdr
                        lngIncCols(lngInc) = lngCol
                        If strHdr <> "" Then varHdrOut(1, lngInc) = strHdr
                    End If
                End If
            Next lngCol
            ' Multi-row games: count the example's extra rows - blank game #, with a value in the
            ' level's FIRST input column (leftmost mapped source column). Only that column is
            ' tested, so text further right that happens to share the rows (a legend or board
            ' table) cannot make a single-row level look multi-row. Games of this level may then
            ' carry up to that many extra rows. Single-row levels stop on the first check.
            lngFirstCol = 0
            For lngHdr = 1 To lngInc
                If lngIncCols(lngHdr) <> 0 Then
                    If lngFirstCol = 0 Or lngIncCols(lngHdr) < lngFirstCol Then lngFirstCol = lngIncCols(lngHdr)
                End If
            Next lngHdr
            lngExCont = 0
            If lngFirstCol > 0 Then
                Do While lngRow + lngExCont < lngEndRow
                    If Len(SafeText(varData(lngRow + lngExCont + 1, 2))) > 0 Then Exit Do
                    If Len(SafeText(varData(lngRow + lngExCont + 1, lngFirstCol))) = 0 Then Exit Do
                    lngExCont = lngExCont + 1
                Loop
            End If

        ElseIf Len(strCV) > 0 And IsNumeric(strCV) Then
            ' game row
            strGame = strCV
            lngContLeft = lngExCont
            lngRowOut = lngRowOut + 1
            LoadGameRow varData, varOut, lngRow, lngRowOut, strGame, lngIncCols, lngInc

        ElseIf Len(strCV) = 0 And lngContLeft > 0 Then
            ' extra row of a multi-row game (e.g. Player 2), capped at the example's count
            lngContLeft = lngContLeft - 1
            lngRowOut = lngRowOut + 1
            LoadGameRow varData, varOut, lngRow, lngRowOut, strGame, lngIncCols, lngInc

        Else
            lngContLeft = 0
        End If
    Next lngRow

    ' One write each for the headers and the game rows.
    If lngInc > 0 Then wsNew.Cells(HDR_OUT_ROW, 3).Resize(1, lngInc).value = varHdrOut
    If lngRowOut > 0 Then wsNew.Cells(FIRST_OUT_ROW, 2).Resize(lngRowOut, lngInc + 1).value = varOut
    wsNew.UsedRange.Columns.AutoFit

End Sub

' Purpose: Copy one game row from the source array into the output array: game # first, then each
'          mapped input as text ("'" keeps "1/2", "007" and the like as typed). Part of DetailedInputs;
'          no handler of its own, errors propagate to CreateCaseInputsSheet.
Private Sub LoadGameRow(ByRef varData As Variant, ByRef varOut() As Variant, ByVal lngSrcRow As Long, _
                        ByVal lngOutRow As Long, ByVal strGame As String, ByRef lngIncCols() As Long, _
                        ByVal lngInc As Long)
    Dim lngHdr As Long

    varOut(lngOutRow, 1) = strGame
    For lngHdr = 1 To lngInc
        If lngIncCols(lngHdr) <> 0 Then
            varOut(lngOutRow, lngHdr + 1) = "'" & SafeText(varData(lngSrcRow, lngIncCols(lngHdr)))
        End If
    Next lngHdr
End Sub



'--------------------------------------------< OA Robot >--------------------------------------------
' Function:             SaveAnswersToLeft
' Description:          Saves references to the selected cells in the green answer cells to the left on the same row.
' Created By:           Erik Oehm
' Source:               https://github.com/ExcelRobot/MEWC-Robot/blob/main/MEWC%20Robot.xlsm
'----------------------------------------------------------------------------------------------------
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Save Answers To Left
' Description:            Saves references to the selected cells in the green answer cells to the left on the same row.
' Macro Expression:       modCaseSetup.SaveAnswersToLeft()
' Generated:              01/08/2025 01:34 PM
'----------------------------------------------------------------------------------------------------
Sub SaveAnswersToLeft(Optional ByVal lngAnswerCol As Long = 0)
    Dim cell As Range
    Dim greenCol As Long
    Dim dest As Range

    ' SELF-CALIBRATING ANSWER COLOUR, added 2026-09-22.
    ' This used to match one hardcoded constant - MEWC's green, 3631104. Local chapter cases
    ' are known to use a different green, and against one of those this command found nothing
    ' and did nothing, silently. Now the case tells us its own colour.
    Dim lngAnswerColour As Long
    Dim lngSelCol As Long

    lngSelCol = Selection.Cells(1, 1).Column

    ' THE COLUMN, IF THE CALLER ALREADY KNOWS IT. Save From Example works it out from the
    ' "Answer" HEADER, which is the reliable route, and then used to throw it away so this
    ' could rediscover it by colour. Deriving the same column twice by two different rules
    ' is what let them disagree. When it is handed over, no colour matching happens at all.
    greenCol = lngAnswerCol

    ' THE HEADER ROUTE, added 2026-09-23 - the same route Save From Example uses, now taken by
    ' this command when it is run ON ITS OWN and nothing was handed to it. Before this, a
    ' standalone run went straight to colour, and colour takes the first NON-WHITE fill to the
    ' left: on the 2026 Airplane Battleship case that is the pale-yellow #FFFFCC INPUT column,
    ' so the command wrote "=H16".."=H20" over the games' own guess cells and left the green
    ' answer column untouched. Silent - nothing in the log, nothing on the status bar. The
    ' 2026-09-22 fix taught the scan to skip white; it had no way to know about any other
    ' non-answer fill. The header is the reliable route, so try it first and keep colour as
    ' the fallback for a sheet that genuinely has no "Answer" header.
    If greenCol = 0 Then
        greenCol = AnswerColumnAbove(ActiveSheet, ActiveCell.Row, lngSelCol)
    End If

    If greenCol = 0 Then
        lngAnswerColour = AnswerFillColour(Selection)

        ' Find the answer cell on this row. LEFT of the selection only - the command is
        ' "save answers to LEFT", and without that bound the first matching cell anywhere
        ' in the row wins, which is how the answers ended up in column A.
        For Each cell In Intersect(ActiveCell.EntireRow, ActiveSheet.UsedRange)
            If cell.Column < lngSelCol Then
                If cell.Interior.Color = lngAnswerColour Then
                    greenCol = cell.Column
                    Exit For
                End If
            End If
        Next
    End If

    If greenCol <> 0 Then
        Dim calcMode As Integer
        On Error Resume Next
        calcMode = Application.Calculation
        Application.Calculation = xlCalculationManual
        For Each cell In Selection
            ' When the caller handed the column over, lngAnswerColour was never worked out.
            ' Testing the target cell against 0 would match nothing and the command would
            ' write no answers at all - so only colour-test when colour is how we got here.
            If lngAnswerColour = 0 _
               Or ActiveSheet.Cells(cell.Row, greenCol).Interior.Color = lngAnswerColour Then
                ActiveSheet.Cells(cell.Row, greenCol).Formula = "=" & cell.Address(False, False)
                If dest Is Nothing Then
                    Set dest = ActiveSheet.Cells(cell.Row, greenCol)
                Else
                    Set dest = Union(dest, ActiveSheet.Cells(cell.Row, greenCol))
                End If
            End If
        Next
        Application.Calculation = calcMode
        On Error GoTo 0

        ' if some green cells were saved to, select them and copy either those answers or the formula below.
        If Not dest Is Nothing Then
            dest.Select
            If Left(dest(1).Offset(dest.Rows.Count + 1).Formula, 1) = "=" Then
                dest(1).Offset(dest.Rows.Count + 1).Select
            Else
                dest.Select
            End If
            ShowRange Selection
            Selection.Copy
        End If
    End If
End Sub

' Purpose: True if row lngRow carries anything at all between column B and the solve column - which is
'          what separates a real question row from the BLANK SPACER that sits between a level's worked
'          example and its questions. Every case has that spacer (Jaq, 2026-09-22: "there's never no
'          spacer between an example and the questions"), and it is the one row whose answer cell is
'          empty for a reason that has nothing to do with the question being unanswered.
Private Function RowHasContent(ByVal ws As Worksheet, ByVal lngRow As Long, _
                               ByVal lngSolveCol As Long) As Boolean

    Dim varRow  As Variant
    Dim lngLast As Long
    Dim lngCol  As Long

    On Error GoTo ErrHandler

    lngLast = lngSolveCol
    If lngLast < 2 Then lngLast = 2
    ' One read for the row rather than a cell at a time - this runs once per row scanned.
    varRow = ws.Range(ws.Cells(lngRow, 2), ws.Cells(lngRow, lngLast)).Value2
    If Not IsArray(varRow) Then
        RowHasContent = (Len(SafeText(varRow)) > 0)
        GoTo Cleanup
    End If
    For lngCol = 1 To UBound(varRow, 2)
        If Len(SafeText(varRow(1, lngCol))) > 0 Then
            RowHasContent = True
            GoTo Cleanup
        End If
    Next lngCol

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Fail CLOSED: an unreadable row is not treated as a question row, so the worst case is
    ' that the command does nothing rather than writing into the wrong place.
    RowHasContent = False
    Resume Cleanup
End Function

' Purpose: The answer column for the level a row sits in, found by walking UP to that level's header
'          row and reading its "Answer" header - the route Save From Example already uses, made
'          available to a standalone Save Answers To Left, which has no example row to work from.
'          Returns 0 when there is no header above, or when the header's Answer column is not LEFT
'          of the selection (this serves "Save Answers to LEFT"); the caller then falls back to
'          colour. Never raises.
Private Function AnswerColumnAbove(ByVal ws As Worksheet, ByVal lngRow As Long, _
                                   ByVal lngSelCol As Long) As Long

    ' --- CONSTANTS (local to function) ---
    Const MAX_LOOK_UP As Long = 200   ' a level's own header is never further above than this

    Dim lngScan As Long
    Dim lngCol  As Long
    Dim lngTop  As Long

    On Error GoTo ErrHandler

    lngTop = lngRow - MAX_LOOK_UP
    If lngTop < 1 Then lngTop = 1

    ' Nearest header wins - stop at the first row that has an "Answer" header, so a level takes
    ' its OWN header and not the one belonging to the level above it.
    For lngScan = lngRow To lngTop Step -1
        lngCol = AnswerColumn(ws, lngScan)
        If lngCol > 0 Then
            If lngCol < lngSelCol Then AnswerColumnAbove = lngCol
            Exit For
        End If
    Next lngScan

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    AnswerColumnAbove = 0
    Resume Cleanup
End Function

' Purpose: The fill colour this sheet uses for answer cells, worked out from the sheet itself rather
'          than assumed. Scans LEFT from the selection along the active row and takes the first cell
'          that has any fill - the answer column is left of the working area by definition of
'          Save Answers To Left. Falls back to MEWC's green only when the row carries no fill at all.
'          WHY: the hardcoded green breaks on local chapter cases that use a different colour, and it
'          breaks silently, which is the worst way for a command to fail mid-case.
Public Function AnswerFillColour(ByVal rngSel As Range) As Long

    ' --- CONSTANTS (local to function) ---
    Const MEWC_GREEN As Long = 3631104
    Const PLAIN_WHITE As Long = 16777215

    Dim lngCol  As Long
    Dim lngLeft As Long
    Dim rngCell As Range

    On Error GoTo ErrHandler

    lngLeft = rngSel.Cells(1, 1).Column
    For lngCol = lngLeft - 1 To 1 Step -1
        ' rngSel's OWN row, not ActiveCell's. They are the same when the command is run by
        ' hand and they are not when it is called from Save From Example, which selects the
        ' filled block first.
        Set rngCell = rngSel.Worksheet.Cells(rngSel.Cells(1, 1).Row, lngCol)
        If rngCell.Interior.ColorIndex <> xlColorIndexNone Then
            ' WHITE IS NOT AN ANSWER COLOUR. Cases routinely carry a white fill across the
            ' sheet, and taking that as the answer colour made EVERY cell in the row match
            ' - so the answers were written into column A. Reported by Jaq, 2026-09-22:
            ' "it didn't identify the colour of the example answer cell properly because it
            ' was putting it in column A, not looking for the colour green." Skip white and
            ' keep looking left.
            If rngCell.Interior.Color <> PLAIN_WHITE Then
                AnswerFillColour = rngCell.Interior.Color
                GoTo Cleanup
            End If
        End If
    Next lngCol

    AnswerFillColour = MEWC_GREEN

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    AnswerFillColour = MEWC_GREEN
    Resume Cleanup
End Function

' Purpose: Solve once on the worked-example row, then apply it to every game in the level and save the
'          answers - in one command.
'
'          Select your solve cells ON THE EXAMPLE ROW (one row, however many columns), then run this.
'          It copies them to the first real question row, fills down to the end of the level, then
'          saves references into the answer cells and leaves them on the clipboard to paste.
'
'          HOW IT DIFFERS FROM MEWC ROBOT'S "Save From Example", which it is modelled on:
'            1. MEWC's copies to Offset(2) - exactly two rows down. That assumes one example row and
'               one blank row. Cases with two or three worked examples land in the wrong place. This
'               one finds the first row whose ANSWER CELL IS EMPTY, because examples always have their
'               answer filled in and real questions never do. That skips any number of example rows,
'               and it never reads the word "Example" in any language.
'            2. The answer column is found by its "Answer" HEADER TEXT where possible, falling back to
'               the sheet's own answer fill colour - not a hardcoded green.
'          Auto Fill Down does the rest: it works out how far the level runs, so 10 questions or 20
'          makes no difference.
Public Sub SaveFromExample()

    ' --- CONSTANTS (local to function) ---
    Const MAX_SCAN_ROWS As Long = 200

    Dim ws          As Worksheet
    Dim rngSel      As Range
    Dim rngBlock    As Range
    Dim lngExRow    As Long
    Dim lngAnsCol   As Long
    Dim lngSolveCol As Long
    Dim lngFirstRow As Long
    Dim lngLastRow  As Long
    Dim lngRow      As Long
    Dim lngColour   As Long
    Dim strCell     As String

    On Error GoTo ErrHandler

    Set rngSel = Selection
    If rngSel Is Nothing Then GoTo Cleanup
    If rngSel.Areas.Count <> 1 Then GoTo Cleanup
    If rngSel.Rows.Count <> 1 Then GoTo Cleanup

    Set ws = rngSel.Worksheet
    lngExRow = rngSel.Row

    ' The answer column: by header text on the level's header row (HeaderRowFor), else by the
    ' sheet's own answer fill colour.
    lngAnsCol = AnswerColumn(ws, HeaderRowFor(ws, lngExRow))
    If lngAnsCol = 0 Then
        lngColour = AnswerFillColour(rngSel)
        For lngRow = rngSel.Cells(1, 1).Column - 1 To 1 Step -1
            If ws.Cells(lngExRow, lngRow).Interior.ColorIndex <> xlColorIndexNone Then
                If ws.Cells(lngExRow, lngRow).Interior.Color = lngColour Then
                    lngAnsCol = lngRow
                    Exit For
                End If
            End If
        Next lngRow
    End If
    If lngAnsCol = 0 Then GoTo Cleanup

    lngSolveCol = rngSel.Cells(1, 1).Column

    ' First real question row = the first row below the example that HAS CONTENT and whose
    ' ANSWER CELL IS EMPTY. Worked examples always carry their answer; questions do not, so
    ' this skips any number of worked rows and never reads the word "Example" in any
    ' language.
    '
    ' A WHOLLY BLANK ROW IS THE SPACER, and it is stepped over rather than taken as the
    ' first question. Jaq, 2026-09-22: "there's never no spacer between an example and the
    ' questions" - so the blank row is part of the layout, not a question with no answer
    ' yet. The old loop accepted it and then stopped scanning, which put the solve on the
    ' blank row and left the level untouched. Measured on a three-level case the same day.
    For lngRow = lngExRow + 1 To lngExRow + MAX_SCAN_ROWS
        If lngRow >= ws.Rows.Count Then Exit For
        If RowHasContent(ws, lngRow, lngSolveCol) Then
            ' The next level's heading or the bonus block: the level is over and there was
            ' no question row in it. A FURTHER WORKED ROW (Example3b) is NOT the end - it is
            ' still the header, so step over it and keep looking.
            strCell = SafeText(ws.Cells(lngRow, 2).Value2)
            If IsSectionBreak(strCell) And Not IsExampleLabel(strCell) Then Exit For
            If Len(SafeText(ws.Cells(lngRow, lngAnsCol).Value2)) = 0 Then
                lngFirstRow = lngRow
                Exit For
            End If
        End If
    Next lngRow
    If lngFirstRow = 0 Then
        Application.StatusBar = "Save From Example: no question row with an empty answer cell below the example."
        GoTo Cleanup
    End If

    ' The level's last question row: keep going while the rows carry content and none of
    ' them starts a new section. The blank spacer before the next level ends it, and so
    ' does a level heading with no spacer before it.
    lngLastRow = lngFirstRow
    For lngRow = lngFirstRow + 1 To lngFirstRow + MAX_SCAN_ROWS
        If lngRow >= ws.Rows.Count Then Exit For
        If Not RowHasContent(ws, lngRow, lngSolveCol) Then Exit For
        If IsSectionBreak(SafeText(ws.Cells(lngRow, 2).Value2)) Then Exit For
        lngLastRow = lngRow
    Next lngRow

    ' Copy the solve row over the whole level in one hit.
    '
    ' This used to copy one row and then call Formula Robot's Auto Fill Down through
    ' CreateObject("OARobot.ExcelAddin").RunCommandByName to find the level's end. Two
    ' things were wrong with that. It was an undeclared dependency on another collection,
    ' in the command whose selling point is that it is self-contained; and the call fails
    ' outright in an automated session ("OA Robot formula processing has not been
    ' initialized"), which took the fill-down AND the answer-saving with it, silently,
    ' because both live after it. The level's extent is the case's own structure and this
    ' module already knows how to read it.
    Set rngBlock = ws.Range(ws.Cells(lngFirstRow, lngSolveCol), _
                            ws.Cells(lngLastRow, lngSolveCol + rngSel.Columns.Count - 1))
    rngSel.Copy rngBlock
    Application.CutCopyMode = False

    ' The answer lives in the LAST column of the filled block - the working may be several
    ' columns wide, and only its result is the answer.
    rngBlock.Columns(rngBlock.Columns.Count).Select

    ' Hand the answer column straight over rather than letting it be rediscovered by fill
    ' colour. It was worked out from the "Answer" header above, which is the reliable
    ' route; the colour scan is only the fallback for a sheet that has no header.
    SaveAnswersToLeft lngAnsCol

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "SaveFromExample", Err.Number, Err.Description
    Resume Cleanup
End Sub

'sub to allow editing and save a working copy of the active workbook
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Enable Editing And Save Copy
' Description:            Enable editing and save copy of file with suffix based on active cell (otherwise Working)
' Macro Expression:       modCaseSetup.EnableEditingAndSaveCopy([[ActiveCell]])
' Generated:              01/08/2025 05:11 PM
'----------------------------------------------------------------------------------------------------
Sub SaveCopy(Optional strSuff As String = "Working")
    Dim wb As Workbook
    Dim strPath As String
    Dim strBase As String
    Dim strExt As String
    Dim lngDot As Long
    Dim lngTry As Long

    ' Set WB to the active workbook
    Set wb = ActiveWorkbook
    If IsError(strSuff) Then
        strSuff = "Working"
    ElseIf Len(strSuff) = 0 Then
        strSuff = "Working"
    End If

    ' No MsgBox anywhere in this routine. CompBot is used under time pressure in
    ' competition, where a modal dialog hidden behind a window is worse than a
    ' silent no-op. Anything worth saying goes to the status bar.
    If wb Is Nothing Then Exit Sub

    ' Allow editing if the workbook is protected
    If wb.ProtectStructure Then
        wb.Unprotect ' Unprotect the workbook (password may be required if protected with one)
    End If

    ' An unsaved workbook has no folder to write the copy into.
    If Len(wb.Path) = 0 Then
        Application.StatusBar = "Save Copy: save the workbook first."
        Exit Sub
    End If

    ' Split on the LAST dot so any extension survives - .xlsm, .xlsb, .xltm.
    ' The old code matched the literal ".xlsx", so on any other extension the
    ' replace did nothing and SaveAs wrote straight over the original.
    lngDot = InStrRev(wb.Name, ".")
    If lngDot > 0 Then
        strBase = Left$(wb.Name, lngDot - 1)
        strExt = Mid$(wb.Name, lngDot)      ' includes the dot
    Else
        strBase = wb.Name
        strExt = vbNullString
    End If

    strPath = wb.Path & "\" & strBase & "_" & strSuff & strExt

    ' Do not overwrite an existing copy, but never block either - bump a counter
    ' until the name is free. A slightly different file name beats a popup.
    lngTry = 1
    Do While Len(Dir$(strPath)) > 0 And lngTry < 100
        lngTry = lngTry + 1
        strPath = wb.Path & "\" & strBase & "_" & strSuff & "_" & lngTry & strExt
    Loop

    If Len(Dir$(strPath)) > 0 Then
        Application.StatusBar = "Save Copy: no free file name found."
        Exit Sub
    End If

    ' Keep the workbook's own format, so macros survive an .xlsm.
    wb.SaveAs Filename:=strPath, FileFormat:=wb.FileFormat

End Sub

' ====================================================================================================
'  Private helpers
' ====================================================================================================

' Purpose: Return the case sheet - "Case", else "Case-Varsity", else "Task" (the MEWC qualification
'          rounds of 2022/2024 call it that; added 2026-09-25, Jaq), else ask for its name.
'          During a Setup run the question is asked at most once and the answer reused by
'          every later step. Returns Nothing if no such sheet exists; callers raise.
Private Function GetCaseSheet(ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet
    Dim strWS As String

    On Error GoTo ErrHandler

    ' Probes: the sheet may simply not exist.
    On Error Resume Next
    Set ws = wb.Worksheets("Case")
    If ws Is Nothing Then Set ws = wb.Worksheets("Case-Varsity")
    If ws Is Nothing Then Set ws = wb.Worksheets("Task")
    On Error GoTo ErrHandler

    If ws Is Nothing Then
        If m_blnInSetup And m_blnCaseAsked Then
            strWS = m_strCaseSheet
        Else
            strWS = InputBox("What is the case sheet called?")
            If m_blnInSetup Then
                m_blnCaseAsked = True
                m_strCaseSheet = strWS
            End If
        End If

        If Len(strWS) > 0 Then
            On Error Resume Next
            Set ws = wb.Worksheets(strWS)
            On Error GoTo ErrHandler
        End If
    End If

    Set GetCaseSheet = ws
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Set GetCaseSheet = Nothing
End Function

' Purpose: The start row of every level on a case sheet, top to bottom: column B cells reading
'          "Level #" or "Section #", with Contents-list entries removed. Shared by Create Level Sheets
'          and Create Data Table. Returns an empty collection when there are no levels, and Nothing when the
'          read itself failed (logged) - so callers can tell "no levels" from "could not read".
Public Function LevelMarkerRows(ByVal ws As Worksheet) As Collection

    Dim colRows As Collection
    Dim varColB As Variant
    Dim lngLastRow As Long
    Dim lngRow As Long

    On Error GoTo ErrHandler

    Set colRows = New Collection
    lngLastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lngLastRow >= 2 Then
        varColB = ws.Range(ws.Cells(1, 2), ws.Cells(lngLastRow, 2)).Value2
        For lngRow = 1 To lngLastRow
            If IsLevelMarker(varColB(lngRow, 1)) Then colRows.Add lngRow
        Next lngRow
    End If

    ' Drop Contents-list entries: a marker within m_CONTENTS_MAX_GAP rows of the marker before
    ' or after it. Real levels are never that close.
    Set LevelMarkerRows = DropContentsMarkers(colRows)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If LogError("LevelMarkerRows", Err.Number, Err.Description) Then m_blnLogWritten = True
    Set LevelMarkerRows = Nothing
    Resume Cleanup
End Function

' Purpose: True for a level heading: "Level 3", "LEVEL 3", "Section 3", "Level 3:" (not "Level Code").
Public Function IsLevelMarker(ByVal varValue As Variant) As Boolean

    Dim strText As String

    On Error GoTo ErrHandler

    If VarType(varValue) <> vbString Then GoTo Cleanup
    strText = UCase$(Trim$(Replace(varValue, Chr$(160), " ")))
    If Right$(strText, 1) = ":" Then strText = Left$(strText, Len(strText) - 1)
    If strText = "LEVEL CODE" Then GoTo Cleanup
    If Not (strText Like "LEVEL *" Or strText Like "SECTION *") Then GoTo Cleanup
    IsLevelMarker = IsNumeric(Trim$(Replace(Replace(strText, "LEVEL", ""), "SECTION", "")))

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    IsLevelMarker = False
    Resume Cleanup
End Function

' Purpose: True for a level sheet name made by Create Level Sheets: "L" then digits only.
'          The one definition used by Create Data Table and Solve Level.
Public Function IsLevelSheetName(ByVal strName As String) As Boolean

    On Error GoTo ErrHandler

    If Len(strName) < 2 Then GoTo Cleanup
    If UCase$(Left$(strName, 1)) <> "L" Then GoTo Cleanup
    IsLevelSheetName = Not (Mid$(strName, 2) Like "*[!0-9]*")

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    IsLevelSheetName = False
    Resume Cleanup
End Function

' Purpose: Length of the example word this text STARTS with, or 0 when it starts with none.
'          English, French and Brazilian Portuguese. Case-insensitive.
'          Callers strip spaces first if they want "Example 3" to match.
'
'          ADD A LANGUAGE BY EDITING m_EXAMPLE_WORDS AND NOTHING ELSE. Every place in CompBot that
'          recognises a worked-example row goes through this one function, so the list is the only
'          thing that needs to change.
'
'          The caller must NEVER assume the prefix is 7 characters, even though all three words
'          happen to be. Use the returned length - the next language added will break a hardcoded 7.
' Purpose: The accepted example words as an array, for code that must SEARCH for them (Range.Find)
'          rather than test a string it already has. Same single list as ExamplePrefixLen.
Public Function ExampleWords() As Variant
    ExampleWords = Split(m_EXAMPLE_WORDS, "|")
End Function

Public Function ExamplePrefixLen(ByVal strText As String) As Long

    Dim varWord As Variant
    Dim lngLen  As Long

    On Error GoTo ErrHandler

    ' Trim here, once, rather than at each call site: a LEADING space would otherwise make
    ' " Example3" invisible to any caller that tests the raw string. Callers that already
    ' strip spaces entirely are unaffected - this is a no-op for them.
    strText = Trim$(Replace(strText, Chr$(160), " "))

    For Each varWord In Split(m_EXAMPLE_WORDS, "|")
        lngLen = Len(varWord)
        If Len(strText) >= lngLen Then
            If StrComp(Left$(strText, lngLen), CStr(varWord), vbTextCompare) = 0 Then
                ExamplePrefixLen = lngLen
                GoTo Cleanup
            End If
        End If
    Next varWord

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ExamplePrefixLen = 0
    Resume Cleanup
End Function

' Purpose: True for an example-row label: "Example3", "Example 3", or "Example3a" - an example word,
'          then a number, then an OPTIONAL short letter suffix. All spaces ignored (leading and
'          non-breaking ones too). "Example3:" (the worked example in the instructions) is not a
'          label, nor is "Examples". Also accepts "Exemple3" (fr) and "Exemplo3" (pt-BR).
Public Function IsExampleLabel(ByVal varValue As Variant) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const MAX_SUFFIX As Long = 2

    Dim strText   As String
    Dim strTail   As String
    Dim strSuffix As String
    Dim lngPfx    As Long
    Dim lngDigits As Long

    On Error GoTo ErrHandler

    If VarType(varValue) <> vbString Then GoTo Cleanup
    strText = Replace(Replace(varValue, Chr$(160), ""), " ", "")
    lngPfx = ExamplePrefixLen(strText)
    If lngPfx = 0 Then GoTo Cleanup
    ' Offset taken from the prefix that MATCHED, never a fixed 7.
    If Len(strText) <= lngPfx Then GoTo Cleanup
    strTail = Mid$(strText, lngPfx + 1)

    ' DIGITS, THEN AN OPTIONAL SHORT LETTER SUFFIX. "Example3a" and "Example3b" are the two
    ' worked rows of ONE level, which 2026 cases use routinely - and which Create Data Table
    ' and Solve Level refused outright until 2026-09-22, while Exclude Examples and Save
    ' From Example both read them. Jaq: "fix finding D so that it doesn't refuse example 3A."
    '
    ' THE SUFFIX IS CAPPED AT TWO LETTERS ON PURPOSE. The all-digits test this replaces was
    ' there to stop an instruction line that merely MENTIONS the word being read as a label,
    ' and spaces are stripped before this runs - so "Example 3 of the following" arrives as
    ' "Example3ofthefollowing", and letters with no limit would happily accept it.
    lngDigits = 0
    Do While lngDigits < Len(strTail)
        If Not (Mid$(strTail, lngDigits + 1, 1) Like "[0-9]") Then Exit Do
        lngDigits = lngDigits + 1
    Loop
    If lngDigits = 0 Then GoTo Cleanup

    strSuffix = Mid$(strTail, lngDigits + 1)
    If Len(strSuffix) > MAX_SUFFIX Then GoTo Cleanup
    If Len(strSuffix) > 0 Then
        If strSuffix Like "*[!A-Za-z]*" Then GoTo Cleanup
    End If
    IsExampleLabel = True

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    IsExampleLabel = False
    Resume Cleanup
End Function

' Purpose: True for a column-B cell that ends a level's list of games: another example label, a level
'          heading, a "Game #" or "Level Code" header, or a Bonus heading ("Bonus Questions", "Bonuses").
'          Case-insensitive.
Public Function IsSectionBreak(ByVal strText As String) As Boolean

    Dim strUpper As String

    On Error GoTo ErrHandler

    strUpper = UCase$(Trim$(Replace(strText, Chr$(160), " ")))
    If Len(strUpper) = 0 Then GoTo Cleanup
    IsSectionBreak = IsExampleLabel(strText) Or IsLevelMarker(strText) _
                     Or strUpper = "GAME #" Or strUpper = "LEVEL CODE" _
                     Or IsBonusHeading(strUpper)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    IsSectionBreak = False
    Resume Cleanup
End Function

' Purpose: True for a bonus-section heading ("Bonus Questions", "BONUSES") - text starting "Bonus" that is
'          not a bonus game label such as "Bonus 1".
Private Function IsBonusHeading(ByVal strText As String) As Boolean

    Dim strUpper As String

    On Error GoTo ErrHandler

    strUpper = UCase$(Trim$(strText))
    If Not (strUpper Like "BONUS*") Then GoTo Cleanup
    IsBonusHeading = Not IsNumeric(Mid$(strUpper, 6))

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    IsBonusHeading = False
    Resume Cleanup
End Function

' Purpose: The rows of a level's games: numbers in lngGameCol below lngExRow, up to lngStopRow. The list
'          ends at a section break, or at the first non-blank non-number once games have started (so a
'          reference list further down is never taken for games). The one rule used by Create Data
'          Table and Solve Level. Returns an empty collection on failure (logged).
Public Function LevelGameRows(ByVal ws As Worksheet, ByVal lngGameCol As Long, ByVal lngExRow As Long, _
                              ByVal lngStopRow As Long) As Collection

    Dim colRows As Collection
    Dim varCol As Variant
    Dim lngRow As Long
    Dim strText As String

    On Error GoTo ErrHandler

    Set colRows = New Collection
    Set LevelGameRows = colRows
    If lngStopRow <= lngExRow Then GoTo Cleanup

    varCol = ws.Range(ws.Cells(lngExRow + 1, lngGameCol), ws.Cells(lngStopRow, lngGameCol)).Value2
    If Not IsArray(varCol) Then
        ReDim varCol(1 To 1, 1 To 1)
        varCol(1, 1) = ws.Cells(lngStopRow, lngGameCol).Value2
    End If

    For lngRow = 1 To UBound(varCol, 1)
        If IsError(varCol(lngRow, 1)) Then
            strText = "#error"
        Else
            strText = Trim$(CStr(varCol(lngRow, 1)))
        End If
        If Len(strText) > 0 Then
            If IsSectionBreak(strText) Then
                ' A SECOND OR THIRD WORKED ROW IS NOT THE END OF THE LEVEL. Example3a and
                ' Example3b are both part of its header and the games start below them.
                ' Once games HAVE started an example label means the next level, and does
                ' end it. Without this, loosening IsExampleLabel to accept "Example3a"
                ' would make a two-worked-row level report no games at all.
                If colRows.Count > 0 Or Not IsExampleLabel(strText) Then Exit For
            End If
            If IsNumeric(strText) Then
                colRows.Add lngExRow + lngRow
            ElseIf colRows.Count > 0 Then
                Exit For
            End If
        End If
    Next lngRow

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "LevelGameRows", Err.Number, Err.Description
    Set LevelGameRows = New Collection
    Resume Cleanup
End Function

' Purpose: The columns of a level's inputs on its example row, left to right: every column right of
'          the game column with a value on the example row, except Level / Points / Answer / Game #
'          headers (on the level's header row, HeaderRowFor) and the Score column straight after Answer (a check mark in answer
'          files, never an input - Jaq, 2026-09-14). Five blank columns in a row end the run: anything
'          further right is calculation, not input. lngACol returns the Answer column, or 2 when there
'          is none within the run. The ONE input rule: BuildDataTable copies these columns into the
'          block, and Create Data Table places the block two columns right of the last of them.
'          Returns Nothing when the read fails (logged).
Public Function LevelInputColumns(ByVal ws As Worksheet, ByVal lngGameCol As Long, ByVal lngExRow As Long, _
                                  ByRef lngACol As Long) As Collection

    Dim colCols As Collection
    Dim lngHdRow As Long
    Dim lngEndCol As Long
    Dim lngCol As Long
    Dim lngBlnk As Long
    Dim strHdr As String
    Dim strVal As String

    On Error GoTo ErrHandler

    Set colCols = New Collection
    lngACol = 2
    lngHdRow = HeaderRowFor(ws, lngExRow)
    If lngHdRow >= 1 Then
        lngEndCol = ws.Cells(lngExRow, ws.Columns.Count).End(xlToLeft).Column
        For lngCol = lngGameCol + 1 To lngEndCol
            strHdr = SafeText(ws.Cells(lngHdRow, lngCol).Value2)
            strVal = SafeText(ws.Cells(lngExRow, lngCol).Value2)
            If strHdr = "Answer" Then lngACol = lngCol
            If lngACol > 2 And lngCol = lngACol + 1 Then
                ' the Score column
            ElseIf strHdr <> "Level" And strHdr <> "Points" And strHdr <> "Answer" And strHdr <> "Game #" Then
                If strVal <> "" Then
                    lngBlnk = 0
                    colCols.Add lngCol
                Else
                    lngBlnk = lngBlnk + 1
                    If lngBlnk >= 5 Then Exit For
                End If
            End If
        Next lngCol
    End If
    Set LevelInputColumns = colCols

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If LogError("LevelInputColumns", Err.Number, Err.Description) Then m_blnLogWritten = True
    Set LevelInputColumns = Nothing
    Resume Cleanup
End Function

' Purpose: True if row lngRow holds a cell reading exactly "Answer" (case as the cases write it).
'          Reads the row as an array - no Range.Find, so the user's Find dialog settings are untouched.
Public Function HasAnswerHeader(ByVal ws As Worksheet, ByVal lngRow As Long) As Boolean
    On Error GoTo ErrHandler
    HasAnswerHeader = (AnswerColumn(ws, lngRow) > 0)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    HasAnswerHeader = False
    Resume Cleanup
End Function

' Purpose: The column of the "Answer" header in row lngRow, or 0. Array read, exact text.
Public Function AnswerColumn(ByVal ws As Worksheet, ByVal lngRow As Long) As Long

    Dim varRow As Variant
    Dim lngLastCol As Long
    Dim lngCol As Long

    On Error GoTo ErrHandler

    If lngRow < 1 Then GoTo Cleanup
    lngLastCol = ws.Cells(lngRow, ws.Columns.Count).End(xlToLeft).Column
    If lngLastCol < 2 Then GoTo Cleanup
    varRow = ws.Range(ws.Cells(lngRow, 1), ws.Cells(lngRow, lngLastCol)).Value2
    For lngCol = 1 To lngLastCol
        If VarType(varRow(1, lngCol)) = vbString Then
            If StrComp(Trim$(varRow(1, lngCol)), "Answer", vbBinaryCompare) = 0 Then
                AnswerColumn = lngCol
                GoTo Cleanup
            End If
        End If
    Next lngCol

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    AnswerColumn = 0
    Resume Cleanup
End Function

' Purpose: The header row of the level whose example label is on lngExRow: the nearest row above it,
'          at most four rows up and never past a "Level #" heading, holding an exact "Answer" header.
'          Real cases put a blank row between the header and the example (header two rows up); the
'          tutorial's demo puts it straight above; a second worked row (Example3b) sits one row lower
'          still. The Game # label varies and inputs may have no header at all, so "Answer" is the
'          one header every case has (Jaq, 2026-09-24). Falls back to two rows above, the old rule,
'          when none is found (so callers still see a row number, possibly < 1).
Public Function HeaderRowFor(ByVal ws As Worksheet, ByVal lngExRow As Long) As Long

    ' --- CONSTANTS (local to function) ---
    Const MAX_ROWS_UP As Long = 4

    Dim lngRow As Long
    Dim lngStop As Long

    On Error GoTo ErrHandler

    HeaderRowFor = lngExRow - 2
    lngStop = lngExRow - MAX_ROWS_UP
    If lngStop < 1 Then lngStop = 1
    For lngRow = lngExRow - 1 To lngStop Step -1
        If IsLevelMarker(ws.Cells(lngRow, 2).Value2) Then Exit For
        If AnswerColumn(ws, lngRow) > 0 Then
            HeaderRowFor = lngRow
            Exit For
        End If
    Next lngRow

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    HeaderRowFor = lngExRow - 2
    Resume Cleanup
End Function

' Purpose: The row of a level's example label in column B between lngTop and lngBottom, or 0.
'          A label with an "Answer" header just above it (HeaderRowFor) wins (so an instruction line that happens to
'          read "Example 3" is passed over); otherwise the first label found.
Public Function FindExampleRow(ByVal ws As Worksheet, ByVal lngTop As Long, ByVal lngBottom As Long) As Long

    Dim varColB As Variant
    Dim lngRow As Long
    Dim lngFirst As Long

    On Error GoTo ErrHandler

    If lngTop < 1 Then lngTop = 1
    If lngBottom < 2 Or lngBottom < lngTop Then GoTo Cleanup
    varColB = ws.Range(ws.Cells(1, 2), ws.Cells(lngBottom, 2)).Value2

    For lngRow = lngTop To lngBottom
        If IsExampleLabel(varColB(lngRow, 1)) Then
            If lngFirst = 0 Then lngFirst = lngRow
            If lngRow > 1 Then
                If HasAnswerHeader(ws, HeaderRowFor(ws, lngRow)) Then
                    FindExampleRow = lngRow
                    GoTo Cleanup
                End If
            End If
        End If
    Next lngRow

    FindExampleRow = lngFirst

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    FindExampleRow = 0
    Resume Cleanup
End Function

' Purpose: The first Bonus heading ("Bonus Questions", "Bonuses", any case) in column B below lngAfterRow, or 0.
Public Function BonusStartRow(ByVal ws As Worksheet, ByVal lngAfterRow As Long) As Long

    Dim varColB As Variant
    Dim lngLastRow As Long
    Dim lngRow As Long

    On Error GoTo ErrHandler

    lngLastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If lngLastRow <= lngAfterRow Or lngLastRow < 2 Then GoTo Cleanup
    varColB = ws.Range(ws.Cells(1, 2), ws.Cells(lngLastRow, 2)).Value2
    For lngRow = lngAfterRow + 1 To lngLastRow
        If VarType(varColB(lngRow, 1)) = vbString Then
            If IsBonusHeading(Replace(varColB(lngRow, 1), Chr$(160), " ")) Then
                BonusStartRow = lngRow
                GoTo Cleanup
            End If
        End If
    Next lngRow

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    BonusStartRow = 0
    Resume Cleanup
End Function

' Purpose: Scroll the active window so the top-left cell of rngShow is on screen, for commands that
'          finish by landing on a cell (Save Answers To Left, Save From Example, Save Answer To Bonus).
'          .Select and Application.Goto do NOT scroll while ScreenUpdating is off (tested 2026-09-28),
'          so the user could not see the answer they had just saved and copied. Only scrolls when the
'          cell is out of view. The sheet must already be active. Cosmetic: a failure is logged only.
Public Sub ShowRange(ByVal rngShow As Range)

    Dim rngSeen As Range

    On Error GoTo ErrHandler

    If rngShow Is Nothing Then GoTo Cleanup
    If Not ActiveSheet Is rngShow.Worksheet Then GoTo Cleanup

    With ActiveWindow
        Set rngSeen = Intersect(.VisibleRange, rngShow.Cells(1, 1))
        If rngSeen Is Nothing Then
            ' A few rows of context above and columns to the left, as ShowBlock does.
            .ScrollRow = Application.WorksheetFunction.Max(1, rngShow.Row - 3)
            .ScrollColumn = Application.WorksheetFunction.Max(1, rngShow.Column - 2)
        End If
    End With

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "ShowRange", Err.Number, Err.Description
    Resume Cleanup
End Sub

' Purpose: rngAdd joined onto rngBase (rngBase may be Nothing). Both on the same sheet.
Private Function AddToRange(ByVal rngBase As Range, ByVal rngAdd As Range) As Range

    On Error GoTo ErrHandler

    If rngBase Is Nothing Then
        Set AddToRange = rngAdd
    Else
        Set AddToRange = Application.Union(rngBase, rngAdd)
    End If

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Set AddToRange = rngBase
    Resume Cleanup
End Function

' Purpose: The case sheet of the active workbook (for Solve Level). Public wrapper over the
'          private lookup; may show the "What is the case sheet called?" InputBox (kept by choice).
Public Function CaseSheetOf(ByVal wb As Workbook) As Worksheet

    On Error GoTo ErrHandler

    Set CaseSheetOf = GetCaseSheet(wb)

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Set CaseSheetOf = Nothing
    Resume Cleanup
End Function

' Purpose: Return the marker rows with Contents-list entries removed. A row is a Contents
'          entry when the previous or next marker row is within m_CONTENTS_MAX_GAP rows.
'          Input rows are in ascending order (column B scanned top to bottom).
Private Function DropContentsMarkers(ByVal colRows As Collection) As Collection

    Dim colKeep As Collection
    Dim lngCount As Long
    Dim lngI As Long
    Dim blnNearPrev As Boolean
    Dim blnNearNext As Boolean

    On Error GoTo ErrHandler

    Set colKeep = New Collection
    lngCount = colRows.Count
    For lngI = 1 To lngCount
        blnNearPrev = False
        blnNearNext = False
        If lngI > 1 Then blnNearPrev = (colRows(lngI) - colRows(lngI - 1) <= m_CONTENTS_MAX_GAP)
        If lngI < lngCount Then blnNearNext = (colRows(lngI + 1) - colRows(lngI) <= m_CONTENTS_MAX_GAP)
        If Not (blnNearPrev Or blnNearNext) Then colKeep.Add colRows(lngI)
    Next lngI

    Set DropContentsMarkers = colKeep
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ' Fall back to the unfiltered list rather than losing every level.
    Set DropContentsMarkers = colRows
End Function

' Purpose: Turn " StepA (91: ...); StepB (1004: ...);" into " StepA, StepB".
Private Function StepNames(ByVal strFailed As String) As String

    Dim varParts As Variant
    Dim varPart As Variant
    Dim strName As String
    Dim strOut As String

    On Error GoTo ErrHandler

    varParts = Split(strFailed, ";")
    For Each varPart In varParts
        strName = Trim$(CStr(varPart))
        If InStr(1, strName, " (") > 0 Then strName = Left$(strName, InStr(1, strName, " (") - 1)
        If Len(strName) > 0 Then
            If Len(strOut) > 0 Then strOut = strOut & ","
            strOut = strOut & " " & strName
        End If
    Next varPart

    StepNames = strOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    StepNames = strFailed
End Function

' Purpose: A cell value as text, with error values (#N/A, #REF! ...) read as "".
Private Function SafeText(ByVal varValue As Variant) As String
    On Error GoTo ErrHandler
    If IsError(varValue) Then
        SafeText = vbNullString
    Else
        SafeText = CStr(varValue)
    End If
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    SafeText = vbNullString
End Function












































