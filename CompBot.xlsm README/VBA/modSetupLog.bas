Attribute VB_Name = "modSetupLog"
Option Explicit

'==============================================================================
' FULL SETUP CASE STEP LOG (GitHub #29, Jaq, 2026-10-29). Each setup step writes its row in the
' "Last setup run" block of the Setup Settings sheet AS SOON AS IT ENDS: result, seconds, time and
' a short message. A run that is cancelled (Esc) or crashes still shows how far it got. Written for
' #26 (Dr. Kazza: setup ~100 s on his machine, 5-6 s on Jaq's), so one screenshot shows which step
' is slow.
'
' The cells are found by name: SCS_Log_<key> is the row's Result..Message cells (C:G; D is a spacer).
' The same lines are appended to logs\CompBot\CompBot-setup.log beside CompBot, because CompBot is normally
' open read-only (the cells then last the session only). CompBot is never saved by the log: a CompBot
' that was saved before the run is marked saved again after each write, so it never asks.
'
' The Environment row names the KIND of place (local, OneDrive, web address, network), never a path,
' so a screenshot is safe to share.
'==============================================================================

' --- MODULE CONSTANTS ---
Private Const m_DEBUG_MODE          As Boolean = False
Private Const m_NAME_PREFIX         As String = "SCS_Log_"
Private Const m_LOG_FILE            As String = "CompBot-setup.log"
Private Const m_LOG_COLUMNS         As Long = 5         ' Result | (spacer) | Seconds | Time | Message
Private Const m_MAX_MESSAGE         As Long = 250
Private Const m_SECS_PER_DAY        As Double = 86400#
Private Const m_ENV_KEY             As String = "Env"
Private Const m_TOTAL_KEY           As String = "Total"
Private Const m_NOT_RUN             As String = "NOT RUN"

' --- MODULE VARIABLES ---
Private m_blnActive                 As Boolean      ' a Full Setup Case run is being logged
Private m_blnWasSaved               As Boolean      ' CompBot was saved when the run began
Private m_dblRunStart               As Double       ' Timer at the start of the run
Private m_dblSlowest                As Double
Private m_strSlowest                As String       ' label of the slowest step so far
Private m_dicWritten                As Object       ' keys logged this run
Private m_blnAnyFail                As Boolean      ' a row was logged FAIL (the lambda imports are not in Setup's list)
Private m_strEnv                    As String       ' the Environment row's text, notes first (AddNote)


' Purpose: the step keys, in the order the rows sit in the block. Setup's own step names
'          (modCaseSetup.RunSetupStep) plus the lambda imports and the extra timed rows.
Private Function LogKeys() As Variant
    LogKeys = Array("SaveCopy", "Backup", "NameAllUsedRanges", "CreateLevelSheets", "CreateBonusSheet", _
                    "CreateCaseInputsSheet", "ImportCompBot", "ImportLibrary", "LandCase", "LevelSummary", _
                    "OpenBonusDock", "CalcRestore")
End Function


' Purpose: begin a run: clear the block, write the Environment row and a header line in the text log.
'          Never raises.
Public Sub StartRun(ByVal wbCase As Workbook, ByVal varCalcAtStart As Variant, ByVal blnManualCalc As Boolean)

    Dim varKey      As Variant
    Dim strEnv      As String
    Dim strCase     As String

    On Error GoTo ErrHandler

    m_blnWasSaved = ThisWorkbook.Saved
    m_dblRunStart = Timer
    m_dblSlowest = -1
    m_strSlowest = vbNullString
    m_blnAnyFail = False
    Set m_dicWritten = CreateObject("Scripting.Dictionary")
    m_dicWritten.CompareMode = vbTextCompare
    m_blnActive = True

    For Each varKey In LogKeys()
        WriteRow CStr(varKey), Empty, Empty, Empty, Empty
    Next varKey
    WriteRow m_TOTAL_KEY, Empty, Empty, Empty, Empty

    If Not wbCase Is Nothing Then strCase = wbCase.Name
    strEnv = EnvironmentText(wbCase, varCalcAtStart, blnManualCalc)
    m_strEnv = strEnv
    WriteRow m_ENV_KEY, "INFO", Empty, Now, strEnv
    AppendText "=== Full Setup Case: " & strCase
    AppendText "Environment | " & strEnv

Cleanup:
    KeepSaved
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "modSetupLog.StartRun", Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: a note about the run as a whole, e.g. which sheet was taken as the case sheet (GitHub #30) or which
'          sheets were unprotected (#31). Goes at the FRONT of the Environment row, so the row's length limit
'          never cuts it, and on a line of the text log. Does nothing outside a run. Never raises.
Public Sub AddNote(ByVal strNote As String)

    On Error GoTo ErrHandler
    If Not m_blnActive Then Exit Sub

    strNote = Trim$(HidePaths(strNote))
    If Len(strNote) = 0 Then Exit Sub
    m_strEnv = strNote & " | " & m_strEnv
    WriteRow m_ENV_KEY, "INFO", Empty, Now, Left$(m_strEnv, m_MAX_MESSAGE)
    AppendText "Note | " & strNote

Cleanup:
    KeepSaved
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "modSetupLog.AddNote", Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: log one step: its row in the block and a line in the text log. Does nothing outside a run,
'          so ILC / ILL run on their own never touch the log. Never raises.
Public Sub LogStep(ByVal strKey As String, ByVal strResult As String, ByVal dblSeconds As Double, _
                   ByVal strMessage As String)

    Dim strLabel As String

    On Error GoTo ErrHandler
    If Not m_blnActive Then Exit Sub

    dblSeconds = Round(dblSeconds, 1)
    strMessage = Left$(Trim$(HidePaths(strMessage)), m_MAX_MESSAGE)
    ' A pass with nothing more to say still says so (Jaq, 2026-10-29).
    If strResult = "PASS" And Len(strMessage) = 0 Then strMessage = "Successful"
    WriteRow strKey, strResult, dblSeconds, Now, strMessage
    m_dicWritten(strKey) = True
    If strResult = "FAIL" Then m_blnAnyFail = True

    strLabel = StepLabel(strKey)
    If dblSeconds > m_dblSlowest Then
        m_dblSlowest = dblSeconds
        m_strSlowest = strLabel
    End If
    AppendText strLabel & " | " & strResult & " | " & Format$(dblSeconds, "0.0") & " s | " & strMessage

Cleanup:
    KeepSaved
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "modSetupLog.LogStep", Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: end the run: every step not logged is NOT RUN, the Total row gets the outcome (a PASS becomes FAIL
'          when any row was logged FAIL), the whole time and the slowest step, and the same summary is returned for the status bar, e.g.
'          "Setup 98.2 s: slowest Import CompBot's lambdas 61.0 s". Never raises.
Public Function FinishRun(ByVal strOutcome As String) As String

    Dim varKey      As Variant
    Dim dblTotal    As Double
    Dim strSummary  As String

    On Error GoTo ErrHandler
    If Not m_blnActive Then Exit Function

    For Each varKey In LogKeys()
        If Not m_dicWritten.Exists(CStr(varKey)) Then
            WriteRow CStr(varKey), m_NOT_RUN, Empty, Empty, vbNullString
            AppendText StepLabel(CStr(varKey)) & " | " & m_NOT_RUN
        End If
    Next varKey

    If strOutcome = "PASS" And m_blnAnyFail Then strOutcome = "FAIL"
    dblTotal = Round(SecondsSince(m_dblRunStart), 1)
    strSummary = "Setup " & Format$(dblTotal, "0.0") & " s"
    If Len(m_strSlowest) > 0 And m_dblSlowest > 0 Then
        strSummary = strSummary & ": slowest " & m_strSlowest & " " & Format$(m_dblSlowest, "0.0") & " s"
    End If
    WriteRow m_TOTAL_KEY, strOutcome, dblTotal, Now, strSummary
    AppendText "Whole setup | " & strOutcome & " | " & strSummary
    FinishRun = strSummary

Cleanup:
    m_blnActive = False
    KeepSaved
    Set m_dicWritten = Nothing
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "modSetupLog.FinishRun", Err.Number, Err.Description
    Resume Cleanup
End Function


' Purpose: seconds since a Timer reading, across midnight. Never raises.
Public Function SecondsSince(ByVal dblStart As Double) As Double
    SecondsSince = Timer - dblStart
    If SecondsSince < 0 Then SecondsSince = SecondsSince + m_SECS_PER_DAY
End Function


' Purpose: a message with every path or web address in it replaced by "(a path)", so a screenshot of the log
'          is safe to share: a step's error text can carry one ("your lambda library is not at D:\..."). A path
'          runs to the end of a sentence (". "), a bracket, a ";" or "|", or the end. Never raises.
Private Function HidePaths(ByVal strText As String) As String

    ' --- CONSTANTS (local to function) ---
    Const PATH_PATTERN As String = "([A-Za-z]:\\|\\\\|https?://).*?(?=\.\s|\.$|\s\(|;|\||$)"

    Dim objRegEx As Object

    On Error GoTo ErrHandler

    HidePaths = strText
    If Len(strText) = 0 Then GoTo Cleanup
    Set objRegEx = CreateObject("VBScript.RegExp")
    objRegEx.Global = True
    objRegEx.IgnoreCase = True
    objRegEx.Pattern = PATH_PATTERN
    HidePaths = objRegEx.Replace(strText, "(a path)")

Cleanup:
    Set objRegEx = Nothing
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    HidePaths = strText
    Resume Cleanup
End Function


' Purpose: the row label for a key: the text in column B beside its log cells, else the key itself.
Private Function StepLabel(ByVal strKey As String) As String

    Dim rngRow As Range

    On Error GoTo ErrHandler

    StepLabel = strKey
    Set rngRow = LogRow(strKey)
    If rngRow Is Nothing Then GoTo Cleanup
    If Len(CStr(rngRow.Cells(1, 1).Offset(0, -1).Value2)) > 0 Then
        StepLabel = CStr(rngRow.Cells(1, 1).Offset(0, -1).Value2)
    End If

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    StepLabel = strKey
    Resume Cleanup
End Function


' Purpose: write one row of the block in a single assignment, events off (the sheet's Change handler
'          must not see it). A missing name is skipped.
Private Sub WriteRow(ByVal strKey As String, ByVal varResult As Variant, ByVal varSeconds As Variant, _
                     ByVal varTime As Variant, ByVal varMessage As Variant)

    Dim rngRow      As Range
    Dim varOut      As Variant
    Dim blnEvents   As Boolean

    On Error GoTo ErrHandler
    blnEvents = Application.EnableEvents

    Set rngRow = LogRow(strKey)
    If rngRow Is Nothing Then GoTo Cleanup

    ReDim varOut(1 To 1, 1 To m_LOG_COLUMNS)
    varOut(1, 1) = varResult
    varOut(1, 3) = varSeconds
    varOut(1, 4) = varTime
    varOut(1, 5) = varMessage
    Application.EnableEvents = False
    rngRow.Value2 = varOut

Cleanup:
    Application.EnableEvents = blnEvents
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "modSetupLog.WriteRow", Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: the log cells named m_NAME_PREFIX & strKey in this CompBot, or Nothing. Never raises.
Private Function LogRow(ByVal strKey As String) As Range

    Dim nmItem As Name

    ' Probe: an older CompBot, or a deleted name, has no such cells.
    On Error Resume Next
    Set nmItem = ThisWorkbook.Names(m_NAME_PREFIX & strKey)
    If Not nmItem Is Nothing Then Set LogRow = nmItem.RefersToRange
    On Error GoTo 0
End Function


' Purpose: a CompBot that was saved when the run began stays "saved": the log is a view, not a change,
'          and a read-only collection must never ask to be saved.
Private Sub KeepSaved()
    On Error Resume Next                          ' narrow: never raise from the log
    If m_blnWasSaved Then ThisWorkbook.Saved = True
    On Error GoTo 0
End Sub


' Purpose: append one line to logs\CompBot\CompBot-setup.log beside CompBot (modLog.LogFolder). Never raises; when CompBot's folder
'          cannot be resolved to a local path, nothing is written.
Private Sub AppendText(ByVal strLine As String)

    Dim strFolder   As String
    Dim lngFile     As Long

    On Error GoTo ErrHandler

    strFolder = modLog.LogFolder()
    If Len(strFolder) = 0 Then GoTo Cleanup

    lngFile = FreeFile
    Open JoinPath(strFolder, m_LOG_FILE) For Append As #lngFile
    Print #lngFile, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & strLine
    Close #lngFile
    lngFile = 0

Cleanup:
    Exit Sub

ErrHandler:
    ' The text log must never cascade an error into setup.
    On Error Resume Next
    If lngFile <> 0 Then Close #lngFile
End Sub


' Purpose: the Environment row: Excel version, calculation mode, and the KIND of place CompBot, the
'          case, the Solve folder and the lambda library are in. No paths.
Private Function EnvironmentText(ByVal wbCase As Workbook, ByVal varCalcAtStart As Variant, _
                                 ByVal blnManualCalc As Boolean) As String

    Dim strOut      As String
    Dim strBits     As String
    Dim strAutoSave As String
    Dim blnAutoSave As Boolean

    On Error GoTo ErrHandler

    #If Win64 Then
        strBits = "64-bit"
    #Else
        strBits = "32-bit"
    #End If

    strOut = "Excel " & Application.Version & " build " & Application.Build & " " & strBits & _
             " | calc " & CalcModeName(varCalcAtStart)
    If blnManualCalc Then
        strOut = strOut & " (manual during setup: on)"
    Else
        strOut = strOut & " (manual during setup: off)"
    End If

    ' Calculation is Excel-wide: every recalc during setup reaches every open workbook (Jaq, 2026-10-29).
    strOut = strOut & " | other workbooks open: " & OtherWorkbooks(wbCase) & CalcSettingsText()

    strOut = strOut & " | CompBot: " & PlaceKind(ThisWorkbook.Path)
    If ThisWorkbook.ReadOnly Then strOut = strOut & ", read-only" Else strOut = strOut & ", open for editing"

    If wbCase Is Nothing Then
        strOut = strOut & " | case: none"
    Else
        strOut = strOut & " | case: " & PlaceKind(wbCase.Path)
        ' Probe: AutoSaveOn does not exist before Excel 2016/365.
        On Error Resume Next
        blnAutoSave = wbCase.AutoSaveOn
        If Err.Number = 0 Then
            If blnAutoSave Then strAutoSave = ", AutoSave on" Else strAutoSave = ", AutoSave off"
        End If
        On Error GoTo ErrHandler
        strOut = strOut & strAutoSave
    End If

    strOut = strOut & " | Solve folder: " & PlaceKind(modSetupSettings.SolveFolder())
    strOut = strOut & " | library: " & PlaceKind(modLambdas.CurrentLambdaLibrary())
    If Not modSetupSettings.SetupStepOn("ImportLibrary") Then strOut = strOut & " (import off)"

Cleanup:
    EnvironmentText = strOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "modSetupLog.EnvironmentText", Err.Number, Err.Description
    strOut = strOut & " | (environment incomplete: " & Err.Description & ")"
    Resume Cleanup
End Function


' Purpose: how many workbooks are open besides the case and CompBot, hidden ones (PERSONAL.XLSB, OA Robot's
'          on-demand collections) counted apart. A count only, never a name. Never raises.
Private Function OtherWorkbooks(ByVal wbCase As Workbook) As String

    Dim wbItem      As Workbook
    Dim lngShown    As Long
    Dim lngHidden   As Long
    Dim blnVisible  As Boolean
    Dim lngW        As Long

    On Error GoTo ErrHandler

    For Each wbItem In Application.Workbooks
        If Not wbItem Is ThisWorkbook And Not wbItem Is wbCase Then
            blnVisible = False
            For lngW = 1 To wbItem.Windows.Count
                If wbItem.Windows(lngW).Visible Then blnVisible = True
            Next lngW
            If blnVisible Then lngShown = lngShown + 1 Else lngHidden = lngHidden + 1
        End If
    Next wbItem
    OtherWorkbooks = lngShown & " (+" & lngHidden & " hidden)"

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    OtherWorkbooks = "unknown"
    Resume Cleanup
End Function


' Purpose: ", recalc before save on, 8 calc threads": the two calculation settings that change how long a
'          recalc during setup takes. Each is probed: either can be unavailable. Never raises.
Private Function CalcSettingsText() As String

    Dim strOut      As String
    Dim blnBefore   As Boolean
    Dim lngThreads  As Long

    ' Probe: CalculateBeforeSave is not readable with no workbook open.
    On Error Resume Next
    blnBefore = Application.CalculateBeforeSave
    If Err.Number = 0 Then
        If blnBefore Then strOut = ", recalc before save on" Else strOut = ", recalc before save off"
    End If
    Err.Clear
    ' Probe: the multi-threading settings are missing on some builds.
    If Application.MultiThreadedCalculation.Enabled Then
        lngThreads = Application.MultiThreadedCalculation.ThreadCount
        If Err.Number = 0 Then strOut = strOut & ", " & lngThreads & " calc threads"
    ElseIf Err.Number = 0 Then
        strOut = strOut & ", 1 calc thread"
    End If
    On Error GoTo 0
    CalcSettingsText = strOut
End Function


' Purpose: the calculation mode as a word.
Private Function CalcModeName(ByVal varCalc As Variant) As String
    Select Case varCalc
        Case xlCalculationAutomatic:        CalcModeName = "Automatic"
        Case xlCalculationManual:           CalcModeName = "Manual"
        Case xlCalculationSemiautomatic:    CalcModeName = "Automatic except tables"
        Case Else:                          CalcModeName = CStr(varCalc)
    End Select
End Function


' Purpose: the KIND of place a path is in, never the path: "not set", "web address (OneDrive personal)",
'          "web address (OneDrive business)", "web address", "OneDrive personal (synced)", "OneDrive
'          business (synced)", "OneDrive (synced)", "network", "network drive" or "local". Never raises.
Private Function PlaceKind(ByVal strPath As String) As String

    ' --- CONSTANTS (local to function) ---
    Const DRIVE_REMOTE  As Long = 3                 ' Scripting.DriveTypeConst Remote

    Dim objFso      As Object
    Dim strKind     As String

    On Error GoTo ErrHandler

    If Len(strPath) = 0 Then
        strKind = "not set"
    ElseIf LCase$(Left$(strPath, 8)) = "https://" Or LCase$(Left$(strPath, 7)) = "http://" Then
        If InStr(1, strPath, "d.docs.live.net", vbTextCompare) > 0 Then
            strKind = "web address (OneDrive personal)"
        ElseIf InStr(1, strPath, "sharepoint", vbTextCompare) > 0 Then
            strKind = "web address (OneDrive business)"
        Else
            strKind = "web address"
        End If
    ElseIf Left$(strPath, 2) = "\\" Then
        strKind = "network"
    ElseIf UnderFolder(strPath, Environ$("OneDriveCommercial")) Then
        strKind = "OneDrive business (synced)"
    ElseIf UnderFolder(strPath, Environ$("OneDriveConsumer")) Then
        strKind = "OneDrive personal (synced)"
    ElseIf UnderFolder(strPath, Environ$("OneDrive")) Then
        strKind = "OneDrive (synced)"
    Else
        strKind = "local"
        If Mid$(strPath, 2, 1) = ":" Then
            Set objFso = CreateObject("Scripting.FileSystemObject")
            If objFso.GetDrive(Left$(strPath, 2)).DriveType = DRIVE_REMOTE Then strKind = "network drive"
        End If
    End If

Cleanup:
    PlaceKind = strKind
    Set objFso = Nothing
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    If Len(strKind) = 0 Then strKind = "unknown"
    Resume Cleanup
End Function


' Purpose: True when strPath is strFolder or inside it. An empty folder is never a match.
Private Function UnderFolder(ByVal strPath As String, ByVal strFolder As String) As Boolean
    If Len(strFolder) = 0 Then Exit Function
    If StrComp(strPath, strFolder, vbTextCompare) = 0 Then
        UnderFolder = True
    Else
        UnderFolder = (StrComp(Left$(strPath, Len(strFolder) + 1), strFolder & "\", vbTextCompare) = 0)
    End If
End Function


