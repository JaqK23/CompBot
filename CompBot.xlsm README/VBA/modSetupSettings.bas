Attribute VB_Name = "modSetupSettings"
Option Explicit

'==============================================================================
' SETUP CASE SETTINGS (SCS) - what Full Setup Case (SC) does, per user. Jaq, 2026-09-24:
' "some people don't like splitting out the levels but may just want to save as and
' import their own lambda library", plus a toggle for whose lambda wins a name clash.
'
' THE SHEET IS THE STORE (Jaq, 2026-09-28). The choices are the yellow cells on this copy of
' CompBot's Setup Settings sheet, so they travel with the user's own CompBot to any machine.
' They used to live in the registry (HKCU, per Windows user), which on a shared arena login
' meant the previous competitor's choices silently applied to the next one.
' Changing a cell saves CompBot at once when it is open for editing; a read-only CompBot
' (OA Robot's normal on-demand open) keeps the change for this session only, and the status
' bar says so. A GitHub update replaces CompBot, so choices go back to the defaults: re-tick.
'
' An unreadable or missing choice reads as the default: every step Yes, CompBot wins.
' Setup (modCaseSetup) and the lambda imports (modLambdas) read the sheet through
' SetupStepOn / CompBotLambdasWin. The lambda library's location stays in the registry
' (modLambdas): it is a path on this machine, not a preference.
'==============================================================================

' --- MODULE CONSTANTS ---
Private Const m_DEBUG_MODE          As Boolean = False
Private Const m_SHEET_CODENAME      As String = "shtSetupSettings"   ' tab: Setup Settings
Private Const m_FIRST_ROW           As Long = 11     ' first step row on the sheet
Private Const m_CHOICE_COL          As Long = 3      ' column C: the yellow choices
Private Const m_LIBRARY_ROW         As Long = 22     ' library file name; its full path is the row below
Private Const m_YES                 As String = "Yes"
Private Const m_NO                  As String = "No"
Private Const m_WINNER_KEY          As String = "LambdaWinner"
Private Const m_WINNER_COMPBOT      As String = "CompBot"
Private Const m_WINNER_LIBRARY      As String = "Your library"
Private Const m_MAX_STATUS          As Long = 250    ' Application.StatusBar rejects long strings

' --- MODULE VARIABLES ---
Private m_blnLoading                As Boolean       ' the sheet is being filled from the store: ignore its Change events


' Purpose: the settings, in sheet order from m_FIRST_ROW. The step keys are Setup's own step names
'          (modCaseSetup.RunSetupStep), so a step and its setting cannot drift apart; the last is the
'          lambda winner.
Private Function SettingKeys() As Variant
    SettingKeys = Array("SaveCopy", "Backup", "NameAllUsedRanges", "CreateLevelSheets", "CreateBonusSheet", _
                        "CreateCaseInputsSheet", "ImportLibrary", "ImportCompBot", m_WINNER_KEY)
End Function


' Purpose: True unless the user has switched this step off. Never raises: a missing or unreadable
'          setting is the default, Yes.
Public Function SetupStepOn(ByVal strKey As String) As Boolean
    SetupStepOn = (StrComp(ReadSetting(strKey), m_NO, vbTextCompare) <> 0)
End Function


' Purpose: True when CompBot's lambdas win a name clash (the default); False when the user's own
'          library wins. The winner's import replaces a same-named lambda; the other never does.
Public Function CompBotLambdasWin() As Boolean
    CompBotLambdasWin = (StrComp(ReadSetting(m_WINNER_KEY), m_WINNER_LIBRARY, vbTextCompare) <> 0)
End Function


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Setup Case Settings
' Description:            Shows CompBot's Setup Settings sheet: choose what Full Setup Case does
' Macro Expression:       modSetupSettings.ShowSetupSettings()
'----------------------------------------------------------------------------------------------------
' Purpose: reload the sheet from the store and put the user on it. Activate is deliberate here -
'          showing the sheet is the whole point. A hidden CompBot window is made visible, and the
'          status bar says how to hide it again.
Public Sub ShowSetupSettings()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "ShowSetupSettings"

    Dim wsSettings  As Worksheet
    Dim blnWasHidden As Boolean
    Dim blnWasSaved As Boolean
    Dim strStatus   As String

    On Error GoTo ErrHandler

    Set wsSettings = SettingsSheet()
    If wsSettings Is Nothing Then
        strStatus = "Setup Case Settings: this CompBot has no Setup Settings sheet. Update CompBot."
        GoTo Cleanup
    End If

    RefreshSetupSettings
    ' Showing a hidden window marks the workbook changed; CompBot must not ask to be saved over it.
    blnWasSaved = ThisWorkbook.Saved
    blnWasHidden = Not ThisWorkbook.Windows(1).Visible
    If blnWasHidden Then ThisWorkbook.Windows(1).Visible = True
    If blnWasSaved Then ThisWorkbook.Saved = True
    ThisWorkbook.Activate
    wsSettings.Activate
    wsSettings.Cells(m_FIRST_ROW, m_CHOICE_COL).Select

    If ThisWorkbook.ReadOnly Then
        strStatus = "Setup Case Settings: CompBot is read-only, so a change lasts this session only; open " & _
                    "CompBot for editing to keep it."
    Else
        strStatus = "Setup Case Settings: tick a box or pick a winner and it is saved in this CompBot at once."
    End If
    If blnWasHidden Then strStatus = strStatus & " CompBot was hidden: View > Hide puts it back."

Cleanup:
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strStatus = "Setup Case Settings failed: " & Err.Description & " (see " & LogLocation() & ")."
    Resume Cleanup
End Sub


' Purpose: fill the sheet from the store: every choice, and the lambda library set with SLL.
'          Called by SCS, by the sheet's Worksheet_Activate, and by SLL / CLL. Never raises.
Public Sub RefreshSetupSettings()

    Dim wsSettings  As Worksheet
    Dim varKeys     As Variant
    Dim varOut      As Variant
    Dim strLibrary  As String
    Dim lngI        As Long
    Dim blnEvents   As Boolean
    Dim blnWasSaved As Boolean

    On Error GoTo ErrHandler
    blnEvents = Application.EnableEvents
    ' The sheet is a view of the registry: filling it is not a change to CompBot, so a CompBot that
    ' was saved stays "saved" and never asks (a read-only collection must never be saved).
    blnWasSaved = ThisWorkbook.Saved

    Set wsSettings = SettingsSheet()
    If wsSettings Is Nothing Then GoTo Cleanup

    varKeys = SettingKeys()
    ReDim varOut(1 To UBound(varKeys) - LBound(varKeys) + 1, 1 To 1)
    For lngI = LBound(varKeys) To UBound(varKeys)
        varOut(lngI - LBound(varKeys) + 1, 1) = DisplayValue(CStr(varKeys(lngI)))
    Next lngI

    strLibrary = modLambdas.CurrentLambdaLibrary()

    m_blnLoading = True
    Application.EnableEvents = False
    wsSettings.Cells(m_FIRST_ROW, m_CHOICE_COL).Resize(UBound(varOut, 1), 1).Value2 = varOut
    If Len(strLibrary) = 0 Then
        wsSettings.Cells(m_LIBRARY_ROW, m_CHOICE_COL).Value2 = "none set"
        wsSettings.Cells(m_LIBRARY_ROW + 1, m_CHOICE_COL).Value2 = "Run Set Lambda Library (SLL) to choose one"
    Else
        wsSettings.Cells(m_LIBRARY_ROW, m_CHOICE_COL).Value2 = Mid$(strLibrary, InStrRev(strLibrary, "\") + 1)
        wsSettings.Cells(m_LIBRARY_ROW + 1, m_CHOICE_COL).Value2 = strLibrary
    End If

Cleanup:
    On Error Resume Next
    Application.EnableEvents = blnEvents
    If blnWasSaved Then ThisWorkbook.Saved = True
    m_blnLoading = False
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "RefreshSetupSettings", Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: the sheet's Worksheet_Change handler. Saves each changed choice to the store; a value
'          that is not one of the choices is put back to what is saved, and the status bar says so.
Public Sub SetupSettingsChanged(ByVal rngTarget As Range)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "SetupSettingsChanged"

    Dim rngChoices  As Range
    Dim rngHit      As Range
    Dim rngCell     As Range
    Dim varKeys     As Variant
    Dim strKey      As String
    Dim strValue    As String
    Dim strStatus   As String
    Dim blnSave     As Boolean

    On Error GoTo ErrHandler
    If m_blnLoading Then Exit Sub

    varKeys = SettingKeys()
    Set rngChoices = rngTarget.Worksheet.Cells(m_FIRST_ROW, m_CHOICE_COL).Resize(UBound(varKeys) - LBound(varKeys) + 1, 1)
    Set rngHit = Application.Intersect(rngTarget, rngChoices)
    If rngHit Is Nothing Then Exit Sub

    For Each rngCell In rngHit.Cells
        strKey = CStr(varKeys(LBound(varKeys) + rngCell.Row - m_FIRST_ROW))
        strValue = CellChoice(strKey, rngCell.Value2)
        If Len(strValue) = 0 Then
            strStatus = strStatus & " not a choice for " & rngCell.Offset(0, -1).Value2 & ", set to the default."
        Else
            strStatus = strStatus & " " & rngCell.Offset(0, -1).Value2 & ": " & strValue & "."
        End If
    Next rngCell

    ' The sheet IS the store, so keeping a choice means saving CompBot.
    If ThisWorkbook.ReadOnly Then
        strStatus = "Setup Case Settings, THIS SESSION ONLY (CompBot is read-only; open it for editing and " & _
                    "change it there to keep it):" & strStatus
    Else
        blnSave = True
        strStatus = "Setup Case Settings, saved in this CompBot:" & strStatus
    End If

Cleanup:
    ' Show every choice in its canonical form; a value that is not a choice becomes the default.
    ' Runs on the error path too, so a bad value never stays on the sheet.
    RefreshSetupSettings
    If blnSave Then
        On Error Resume Next                      ' narrow: a failed save is reported, never raised
        ThisWorkbook.Save
        If Err.Number <> 0 Then strStatus = "Setup Case Settings: NOT saved (" & Err.Description & ")."
        On Error GoTo 0
    ElseIf ThisWorkbook.ReadOnly Then
        ThisWorkbook.Saved = True                 ' a read-only collection must never ask to be saved
    End If
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError PROC_NAME, Err.Number, Err.Description
    strStatus = "Setup Case Settings: not saved, " & Err.Description & " (see " & LogLocation() & ")."
    Resume Cleanup
End Sub


' Purpose: the stored value, as the sheet shows it: TRUE/FALSE for a step (the cells are in-cell
'          checkboxes; an older Excel shows the words), the winner's name for the winner.
Private Function DisplayValue(ByVal strKey As String) As Variant
    If strKey = m_WINNER_KEY Then
        If CompBotLambdasWin() Then DisplayValue = m_WINNER_COMPBOT Else DisplayValue = m_WINNER_LIBRARY
    Else
        DisplayValue = SetupStepOn(strKey)
    End If
End Function


' Purpose: a typed value in its stored spelling, or "" if it is not a choice for this setting.
'          Accepts y/n and TRUE/FALSE for the steps, and "library" / "yours" for the winner.
Private Function CanonicalValue(ByVal strKey As String, ByVal strTyped As String) As String

    Dim strLow As String

    strLow = LCase$(strTyped)
    If strKey = m_WINNER_KEY Then
        If strLow = "compbot" Then
            CanonicalValue = m_WINNER_COMPBOT
        ElseIf strLow = "your library" Or strLow = "library" Or strLow = "yours" Then
            CanonicalValue = m_WINNER_LIBRARY
        End If
    Else
        Select Case strLow
            Case "yes", "y", "true":    CanonicalValue = m_YES
            Case "no", "n", "false":    CanonicalValue = m_NO
        End Select
    End If
End Function


' Purpose: one setting in its stored spelling (Yes/No, CompBot/Your library), read from this copy
'          of CompBot's Setup Settings sheet; "" when the sheet or a valid value is missing, so the
'          default applies. Never raises.
Private Function ReadSetting(ByVal strKey As String) As String

    Dim wsSettings  As Worksheet
    Dim varKeys     As Variant
    Dim lngI        As Long

    On Error GoTo ErrHandler

    Set wsSettings = SettingsSheet()
    If wsSettings Is Nothing Then GoTo Cleanup

    varKeys = SettingKeys()
    For lngI = LBound(varKeys) To UBound(varKeys)
        If StrComp(CStr(varKeys(lngI)), strKey, vbTextCompare) = 0 Then
            ReadSetting = CellChoice(strKey, _
                wsSettings.Cells(m_FIRST_ROW + lngI - LBound(varKeys), m_CHOICE_COL).Value2)
            Exit For
        End If
    Next lngI

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    ReadSetting = vbNullString
    Resume Cleanup
End Function


' Purpose: a choice cell's value in its stored spelling, or "" if it is not a choice. A checkbox
'          holds a real Boolean, handled directly: CStr(True) is locale-dependent.
Private Function CellChoice(ByVal strKey As String, ByVal varCell As Variant) As String
    If IsError(varCell) Then Exit Function
    If VarType(varCell) = vbBoolean And strKey <> m_WINNER_KEY Then
        If varCell Then CellChoice = m_YES Else CellChoice = m_NO
    Else
        CellChoice = CanonicalValue(strKey, Trim$(CStr(varCell)))
    End If
End Function


' Purpose: CompBot's Setup Settings sheet, found by CodeName (the tab name can be edited), or Nothing.
Private Function SettingsSheet() As Worksheet

    Dim wsItem As Worksheet

    For Each wsItem In ThisWorkbook.Worksheets
        If StrComp(wsItem.CodeName, m_SHEET_CODENAME, vbTextCompare) = 0 Then
            Set SettingsSheet = wsItem
            Exit Function
        End If
    Next wsItem
End Function









