Attribute VB_Name = "modLambdas"
Option Explicit

' --- MODULE CONSTANTS ---
Private Const m_DEBUG_MODE      As Boolean = False
Private Const m_REPORT_SHEET    As String = "Lambdas Used"
Private Const m_REPORT_TABLE    As String = "tblLambdasUsed"
Private Const m_ANCHOR_ROW      As Long = 3     ' blank rows above, per layout convention
Private Const m_ANCHOR_COL      As Long = 2     ' column A left blank
Private Const m_INCLUDE_NESTED  As Boolean = True
Private Const m_MAX_STATUS      As Long = 250   ' Application.StatusBar rejects long strings
' Where the user's own lambda library is: a per-user Windows setting, never a cell in
' CompBot (a GitHub update replaces CompBot wholesale). See the LAMBDA IMPORTS block.
Private Const m_SETTINGS_APP     As String = "CompBot"
Private Const m_SETTINGS_SECTION As String = "Lambdas"
Private Const m_SETTINGS_KEY     As String = "LibraryPath"
' Stored values (non-LAMBDA names) ILL copies from the user's library carry this in their Name
' Manager comment, so a later import can tell its own copy from a name the case defined itself
' (GitHub #5, Jaq, 2026-10-08: follow whose lambdas win, never overwrite the case's own).
Private Const m_VALUE_TAG        As String = "[CompBot ILL]"

Private m_strScanErrors         As String
' How many lambdas ExpandKeepChain added that nothing DECLARED - reported by
' ClearLambdas, because a silent rescue is indistinguishable from one that never
' happened, and this is the number that says whether the chain is earning its keep.
Private m_lngChainAdded         As Long
Private Const m_REMOVED_SHEET   As String = "Lambdas Removed"
Private Const m_PREVIEW_ONLY    As Boolean = False    ' False = actually delete

'function to list lambdas in this workbook
Function ListLambdas() As Variant
    Dim nm As Name
    Dim strLambdaNames() As String
    Dim lngLambdaCount As Long
    Dim result As Variant

    ' Initialize array for lambda names
    lngLambdaCount = 0

    ' Loop through all named ranges in the workbook
    For Each nm In ActiveWorkbook.Names
        ' Check if the name refers to a LAMBDA function
        If InStr(1, nm.RefersTo, "=LAMBDA(", vbTextCompare) > 0 Then
            lngLambdaCount = lngLambdaCount + 1
            ReDim Preserve strLambdaNames(1 To lngLambdaCount)
            strLambdaNames(lngLambdaCount) = nm.Name
        End If
    Next nm

    ' Handle case where no lambdas are found
    If lngLambdaCount = 0 Then
        ListLambdas = "No LAMBDA functions found."
    Else
        ' Return the list of lambda names
        ListLambdas = strLambdaNames
    End If
End Function

'==============================================================================
' LAMBDA IMPORTS - Set Lambda Library (SLL), Import Lambda Library (ILL),
' Import Lambdas From CompBot (ILC), and Import Case Lambdas (Full Setup Case's
' chained step). Rebuilt 2026-09-24 with Jaq.
'
' THE DESIGN. Everyone keeps their OWN lambda library in a workbook of their
' own. CompBot remembers where it is (SLL) and loads it into whatever workbook
' is active (ILL). Nothing personal is ever written INTO CompBot: a GitHub
' update replaces CompBot.xlsm wholesale, and most users run it hidden and
' read-only anyway, so anything kept inside it is either unsaveable or wiped at
' the next release. The library's location is a per-user Windows setting
' (SaveSetting), which survives every CompBot update.
'
' WHO WINS A NAME CLASH is the user's choice, in Setup Case Settings (SCS;
' Jaq, 2026-09-24). The WINNER's import replaces an existing LAMBDA of the same
' name; the other's never does, it skips and lists. By default CompBot wins: ILC
' replaces, so an older case file picks up CompBot's newer versions, and ILL
' skips, so a personal copy cannot replace CompBot's improved one. Set "Your
' library" and it is the other way round. Order then does not matter. Neither
' ever touches a name that is not a lambda - that is the case's own range or
' constant.
'
' ILC REPLACES Lambda Robot's ImportAllLambdas, which the old IL called. Same
' job, no dependency on an Oehm collection being installed.
'
' No MsgBox anywhere: ILC runs mid-setup in competition, and a dialog nobody
' notices freezes Excel. Every outcome goes to the status bar.
'==============================================================================

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Set Lambda Library
' Macro Expression:       modLambdas.SetLambdaLibrary()
'----------------------------------------------------------------------------------------------------
' Purpose: ask once for the user's own lambda library and remember where it is.
'          A setup-time command: the file picker is its only dialog, and the user asked for it.
Public Sub SetLambdaLibrary()

    Dim varPick     As Variant
    Dim strPath     As String
    Dim strStatus   As String

    On Error GoTo ErrHandler

    varPick = Application.GetOpenFilename( _
                "Excel workbooks (*.xlsx;*.xlsm;*.xlsb;*.xlam),*.xlsx;*.xlsm;*.xlsb;*.xlam", , _
                "Choose YOUR lambda library: the workbook your own lambdas live in")

    If VarType(varPick) = vbBoolean Then
        strStatus = "Set Lambda Library: cancelled, nothing changed. " & LibraryStatus()
        GoTo Cleanup
    End If
    strPath = CStr(varPick)

    If StrComp(FileNameOf(strPath), ThisWorkbook.Name, vbTextCompare) = 0 Then
        strStatus = "Set Lambda Library: that is CompBot itself. Choose the workbook YOUR lambdas live in."
        GoTo Cleanup
    End If

    SaveSetting m_SETTINGS_APP, m_SETTINGS_SECTION, m_SETTINGS_KEY, strPath
    modSetupSettings.RefreshSetupSettings
    strStatus = "Set Lambda Library: " & strPath & ". ILL, and Full Setup Case, now load it."

Cleanup:
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "SetLambdaLibrary", Err.Number, Err.Description
    strStatus = "Set Lambda Library failed: " & Err.Description
    Resume Cleanup
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Clear Lambda Library
' Macro Expression:       modLambdas.ClearLambdaLibrary()
'----------------------------------------------------------------------------------------------------
' Purpose: forget the saved library location, so ILL and Full Setup Case stop loading it.
'          Touches only the setting: no workbook and no lambda is changed. To swap one
'          library for another, Set Lambda Library (SLL) on its own is enough.
Public Sub ClearLambdaLibrary()

    Dim strWas      As String
    Dim strStatus   As String

    On Error GoTo ErrHandler

    strWas = LibraryPath()
    If Len(strWas) = 0 Then
        strStatus = "Clear Lambda Library: no library was set, nothing to clear."
        GoTo Cleanup
    End If

    DeleteSetting m_SETTINGS_APP, m_SETTINGS_SECTION, m_SETTINGS_KEY
    modSetupSettings.RefreshSetupSettings
    strStatus = "Clear Lambda Library: forgot " & strWas & ". Full Setup Case now loads CompBot's lambdas only; run SLL to set a new library."

Cleanup:
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ClearLambdaLibrary", Err.Number, Err.Description
    strStatus = "Clear Lambda Library failed: " & Err.Description
    Resume Cleanup
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Import Lambda Library
' Macro Expression:       modLambdas.ImportLambdaLibrary()
'----------------------------------------------------------------------------------------------------
' Purpose: load the user's own library (set once with SLL) into the active workbook.
'          Existing lambdas are skipped, unless Setup Case Settings says your library wins.
Public Sub ImportLambdaLibrary()

    Dim wbTarget    As Workbook
    Dim strStatus   As String

    On Error GoTo ErrHandler
    VBAInit

    Set wbTarget = ActiveWorkbook
    If wbTarget Is Nothing Then
        strStatus = "ILL FAILED: no workbook is active."
    ElseIf IsProtectedWorkbook(wbTarget, False) Then
        strStatus = "ILL FAILED: " & wbTarget.Name & " is CompBot or your lambda library; run it in a case file."
    Else
        strStatus = "ILL: " & ImportFromLibrary(wbTarget, True, Not modSetupSettings.CompBotLambdasWin())
    End If

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ImportLambdaLibrary", Err.Number, Err.Description
    strStatus = "ILL FAILED: " & Err.Description
    Resume Cleanup
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Import Lambdas From CompBot
' Macro Expression:       modLambdas.ImportLambdasFromCompBot()
'----------------------------------------------------------------------------------------------------
' Purpose: load CompBot's own lambdas into the active workbook. Replaces any older copy of
'          a CompBot lambda already there, unless Setup Case Settings says your library wins.
Public Sub ImportLambdasFromCompBot()

    Dim wbTarget    As Workbook
    Dim strStatus   As String

    On Error GoTo ErrHandler
    VBAInit

    Set wbTarget = ActiveWorkbook
    If wbTarget Is Nothing Then
        strStatus = "ILC: no workbook is active."
    ElseIf wbTarget Is ThisWorkbook Then
        strStatus = "ILC: refused, the active workbook is CompBot itself."
    Else
        strStatus = "ILC: " & CopyLambdas(ThisWorkbook, wbTarget, modSetupSettings.CompBotLambdasWin())
    End If

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ImportLambdasFromCompBot", Err.Number, Err.Description
    strStatus = "ILC failed: " & Err.Description
    Resume Cleanup
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Import Case Lambdas   (hidden; chained after Full Setup Case)
' Macro Expression:       modLambdas.ImportCaseLambdas()
'----------------------------------------------------------------------------------------------------
' Purpose: load the user's library (if one is set) and CompBot's lambdas, each only if it is
'          switched on in Setup Case Settings (SCS). The WINNER goes first and replaces older
'          copies; the other goes second and never overwrites, so where a name clashes the
'          winner's version is the one left. CompBot wins unless SCS says otherwise
'          (Jaq, 2026-09-24).
'          Its report is APPENDED to Full Setup Case's own status-bar line, which would
'          otherwise be lost the moment this ran.
Public Sub ImportCaseLambdas()

    ' --- CONSTANTS (local to function) ---
    Const SETUP_PREFIX As String = "Full Setup Case"

    Dim wbTarget    As Workbook
    Dim varPrior    As Variant
    Dim strPrior    As String
    Dim strStatus   As String
    Dim strMine     As String
    Dim strCompBot  As String
    Dim blnMine     As Boolean
    Dim blnCompBot  As Boolean
    Dim blnCBWins   As Boolean

    On Error GoTo ErrHandler

    varPrior = Application.StatusBar
    If VarType(varPrior) = vbString Then
        If InStr(1, CStr(varPrior), SETUP_PREFIX, vbTextCompare) = 1 Then strPrior = CStr(varPrior) & "  |  "
    End If
    VBAInit

    blnMine = (Len(LibraryPath()) > 0) And modSetupSettings.SetupStepOn("ImportLibrary")
    blnCompBot = modSetupSettings.SetupStepOn("ImportCompBot")
    blnCBWins = modSetupSettings.CompBotLambdasWin()

    Set wbTarget = ActiveWorkbook
    If wbTarget Is Nothing Then
        strStatus = "Lambdas: no workbook is active."
    ElseIf IsProtectedWorkbook(wbTarget, False) Then
        strStatus = "Lambdas: not imported, " & wbTarget.Name & " is CompBot or your lambda library."
    ElseIf Not blnMine And Not blnCompBot Then
        If Len(LibraryPath()) = 0 And modSetupSettings.SetupStepOn("ImportLibrary") Then
            strStatus = "Lambdas: none imported (CompBot's are off in SCS, and no library is set with SLL)."
        Else
            strStatus = "Lambdas: none imported (switched off in Setup Case Settings, SCS)."
        End If
    Else
        ' The winner first, replacing; the other second, skipping what is already there.
        If blnCBWins Then
            If blnCompBot Then strCompBot = "CompBot: " & CopyLambdas(ThisWorkbook, wbTarget, True) & "  "
            If blnMine Then strMine = "Yours: " & ImportFromLibrary(wbTarget, False, False) & "  "
            strStatus = strCompBot & strMine
        Else
            If blnMine Then strMine = "Yours (win): " & ImportFromLibrary(wbTarget, False, True) & "  "
            If blnCompBot Then strCompBot = "CompBot: " & CopyLambdas(ThisWorkbook, wbTarget, False, False) & "  "
            strStatus = strMine & strCompBot
        End If
        If Len(LibraryPath()) = 0 Then strStatus = strStatus & "(no library set)"
        strStatus = "Lambdas: " & Trim$(strStatus)
    End If

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strPrior & strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ImportCaseLambdas", Err.Number, Err.Description
    strStatus = strStatus & " FAILED: " & Err.Description
    Resume Cleanup
End Sub


' Purpose: open the user's library (read-only, its own macros kept quiet), copy its lambdas
'          into wbTarget, close it again, and return a one-line report. A library the user
'          already has open is used as it is and left open.
'          blnListSkipped: name the skipped lambdas in the report (ILL alone), or just count
'          them (Full Setup Case, where CompBot's import shares the line).
'          blnOverwrite: replace a same-named lambda (your library wins) or skip it.
Private Function ImportFromLibrary(ByVal wbTarget As Workbook, ByVal blnListSkipped As Boolean, _
                                   ByVal blnOverwrite As Boolean) As String

    Dim wbLibrary   As Workbook
    Dim strPath     As String
    Dim blnOpened   As Boolean
    Dim blnEvents   As Boolean
    Dim strReport   As String

    On Error GoTo ErrHandler
    blnEvents = Application.EnableEvents

    strPath = LibraryPath()
    If Len(strPath) = 0 Then
        strReport = "no lambda library set. Run Set Lambda Library (SLL) once first."
        GoTo Cleanup
    End If
    If Len(Dir$(strPath)) = 0 Then
        strReport = "your lambda library is not at " & strPath & ". Run SLL to point at it again."
        GoTo Cleanup
    End If

    Set wbLibrary = OpenWorkbookByName(FileNameOf(strPath))
    If wbLibrary Is Nothing Then
        Application.EnableEvents = False         ' the library's own Workbook_Open must not run
        Set wbLibrary = Workbooks.Open(Filename:=strPath, UpdateLinks:=0, ReadOnly:=True, AddToMru:=False)
        blnOpened = True
    End If

    ' The library's stored values come too (GitHub #5); CompBot's own non-LAMBDA names are its
    ' locale settings and never leave it, so only this route passes True.
    strReport = CopyLambdas(wbLibrary, wbTarget, blnOverwrite, blnListSkipped, True)

Cleanup:
    On Error Resume Next                          ' closing must never raise out of here
    If blnOpened Then wbLibrary.Close SaveChanges:=False
    Application.EnableEvents = blnEvents
    On Error GoTo 0
    ImportFromLibrary = strReport
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ImportFromLibrary", Err.Number, Err.Description
    strReport = "could not read " & FileNameOf(strPath) & ": " & Err.Description
    Resume Cleanup
End Function


' Purpose: copy every workbook-level LAMBDA from wbSource into wbTarget, with its comment,
'          and return a one-line report. See the header above for the clash rules.
'          One bad definition is counted as failed and the rest still copy.
'          blnValues (the user's library only - GitHub #5): also copy its STORED VALUES, the
'          non-LAMBDA names with no cell reference (=0.05, ="Red", ={1,2,3}, =SEQUENCE(9)).
'          A name pointing at a range is skipped and counted: copied, it would become an
'          external link back to the library. A copied value is tagged m_VALUE_TAG in its
'          comment. On a clash the winner rule applies, but ONLY to a value carrying the tag:
'          an untagged name is the case's own and is never overwritten (Jaq, 2026-10-08).
Private Function CopyLambdas(ByVal wbSource As Workbook, ByVal wbTarget As Workbook, _
                             ByVal blnOverwrite As Boolean, _
                             Optional ByVal blnListSkipped As Boolean = True, _
                             Optional ByVal blnValues As Boolean = False) As String

    ' --- CONSTANTS (local to function) ---
    Const KIND_OTHER    As Long = 0     ' the case's own name, a range, anything else
    Const KIND_LAMBDA   As Long = 1
    Const KIND_TAGGED   As Long = 2     ' a stored value an earlier ILL copied in

    Dim dicExisting As Object       ' target's workbook-level names -> KIND_*
    Dim nmItem      As Name
    Dim nmNew       As Name
    Dim strName     As String
    Dim strRefers   As String
    Dim strSkipped  As String
    Dim strVSkipped As String
    Dim strReport   As String
    Dim strComment  As String
    Dim blnReplace  As Boolean
    Dim blnIsValue  As Boolean
    Dim lngAdded    As Long
    Dim lngReplaced As Long
    Dim lngSkipped  As Long
    Dim lngFailed   As Long
    Dim lngVAdded   As Long
    Dim lngVReplaced As Long
    Dim lngVSkipped As Long
    Dim lngVFailed  As Long
    Dim lngRanges   As Long

    On Error GoTo ErrHandler

    Set dicExisting = CreateObject("Scripting.Dictionary")
    dicExisting.CompareMode = vbTextCompare
    For Each nmItem In wbTarget.Names
        If IsWorkbookLevelName(nmItem) Then
            If IsLambdaDefinition(SafeRefersTo(nmItem)) Then
                dicExisting(nmItem.Name) = KIND_LAMBDA
            ElseIf InStr(1, SafeComment(nmItem), m_VALUE_TAG, vbBinaryCompare) > 0 Then
                dicExisting(nmItem.Name) = KIND_TAGGED
            Else
                dicExisting(nmItem.Name) = KIND_OTHER
            End If
        End If
    Next nmItem

    For Each nmItem In wbSource.Names
        strName = nmItem.Name
        strRefers = SafeRefersTo(nmItem)
        If Not IsWorkbookLevelName(nmItem) Or StrComp(Left$(strName, 3), "_xl", vbTextCompare) = 0 _
           Or Len(strRefers) = 0 Then GoTo NextName

        If IsLambdaDefinition(strRefers) Then
            blnIsValue = False
        ElseIf Not blnValues Then
            GoTo NextName
        ElseIf IsStoredValue(strRefers) Then
            blnIsValue = True
        Else
            lngRanges = lngRanges + 1                  ' a range or external reference: never copied
            GoTo NextName
        End If

        blnReplace = False
        If dicExisting.Exists(strName) Then
            If blnIsValue Then
                If blnOverwrite And dicExisting(strName) = KIND_TAGGED Then
                    blnReplace = True
                Else
                    lngVSkipped = lngVSkipped + 1
                    strVSkipped = strVSkipped & ", " & strName
                    GoTo NextName
                End If
            ElseIf blnOverwrite And dicExisting(strName) = KIND_LAMBDA Then
                blnReplace = True
            Else
                lngSkipped = lngSkipped + 1
                strSkipped = strSkipped & ", " & strName
                GoTo NextName
            End If
        End If

        Set nmNew = Nothing
        On Error Resume Next                  ' narrow: one bad definition must not stop the rest
        Set nmNew = wbTarget.Names.Add(Name:=strName, RefersTo:=strRefers)
        On Error GoTo ErrHandler

        If nmNew Is Nothing Then
            If blnIsValue Then lngVFailed = lngVFailed + 1 Else lngFailed = lngFailed + 1
        Else
            strComment = SafeComment(nmItem)
            If blnIsValue Then
                ' The tag marks it as ILL's copy; the 255-character comment limit keeps the tag.
                strComment = Trim$(Left$(Replace(strComment, m_VALUE_TAG, vbNullString), _
                                         254 - Len(m_VALUE_TAG)) & " " & m_VALUE_TAG)
            End If
            On Error Resume Next              ' narrow: a comment is nice to have, never fatal
            nmNew.Comment = strComment
            nmNew.Visible = nmItem.Visible
            On Error GoTo ErrHandler
            If blnIsValue Then
                If blnReplace Then lngVReplaced = lngVReplaced + 1 Else lngVAdded = lngVAdded + 1
            Else
                If blnReplace Then lngReplaced = lngReplaced + 1 Else lngAdded = lngAdded + 1
            End If
        End If
NextName:
    Next nmItem

    strReport = lngAdded & " added"
    If blnOverwrite Then strReport = strReport & ", " & lngReplaced & " updated"
    If lngSkipped > 0 Then
        strReport = strReport & ", " & lngSkipped & " skipped (already here"
        If blnListSkipped Then strReport = strReport & ": " & Mid$(strSkipped, 3)
        strReport = strReport & ")"
    End If
    If lngFailed > 0 Then strReport = strReport & ", " & lngFailed & " FAILED to copy"

    If blnValues Then
        strReport = strReport & "; values: " & lngVAdded & " added"
        If blnOverwrite Then strReport = strReport & ", " & lngVReplaced & " updated"
        If lngVSkipped > 0 Then
            strReport = strReport & ", " & lngVSkipped & " skipped (the case's own or CompBot wins"
            If blnListSkipped Then strReport = strReport & ": " & Mid$(strVSkipped, 3)
            strReport = strReport & ")"
        End If
        If lngVFailed > 0 Then strReport = strReport & ", " & lngVFailed & " FAILED to copy"
        If lngRanges > 0 Then strReport = strReport & ", " & lngRanges & " range name(s) not copied"
    End If
    CopyLambdas = strReport & "."
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "CopyLambdas", Err.Number, Err.Description
    CopyLambdas = "stopped after " & lngAdded + lngReplaced & " (" & Err.Description & ")."
End Function


' Purpose: True for CompBot itself and for the user's own lambda library. With
'          blnAnyCompBotName, also for any workbook with "CompBot" in its name (the Clear
'          commands: Jaq, 2026-09-24 - never strip a library, never touch the tutorial).
Private Function IsProtectedWorkbook(ByVal wbCheck As Workbook, ByVal blnAnyCompBotName As Boolean) As Boolean

    Dim strLibrary As String

    If wbCheck Is ThisWorkbook Then
        IsProtectedWorkbook = True
        Exit Function
    End If
    If blnAnyCompBotName And InStr(1, wbCheck.Name, "CompBot", vbTextCompare) > 0 Then
        IsProtectedWorkbook = True
        Exit Function
    End If
    strLibrary = LibraryPath()
    If Len(strLibrary) > 0 Then
        IsProtectedWorkbook = (StrComp(wbCheck.Name, FileNameOf(strLibrary), vbTextCompare) = 0)
    End If
End Function


' Purpose: the saved library location, or "" if none has been set - for the Setup Settings sheet.
Public Function CurrentLambdaLibrary() As String
    CurrentLambdaLibrary = LibraryPath()
End Function


' Purpose: the saved library location, or "" if none has been set.
Private Function LibraryPath() As String
    On Error Resume Next                          ' narrow: a missing registry key is simply "not set"
    LibraryPath = GetSetting(m_SETTINGS_APP, m_SETTINGS_SECTION, m_SETTINGS_KEY, vbNullString)
    On Error GoTo 0
End Function


' Purpose: one phrase saying what the library is currently set to.
Private Function LibraryStatus() As String
    If Len(LibraryPath()) = 0 Then
        LibraryStatus = "No library is set."
    Else
        LibraryStatus = "Library is still " & LibraryPath() & "."
    End If
End Function


' Purpose: the file name at the end of a path.
Private Function FileNameOf(ByVal strPath As String) As String
    FileNameOf = Mid$(strPath, InStrRev(strPath, "\") + 1)
End Function


' Purpose: an open workbook by file name, or Nothing.
Private Function OpenWorkbookByName(ByVal strName As String) As Workbook

    Dim wbOpen As Workbook

    For Each wbOpen In Application.Workbooks
        If StrComp(wbOpen.Name, strName, vbTextCompare) = 0 Then
            Set OpenWorkbookByName = wbOpen
            Exit Function
        End If
    Next wbOpen
End Function


' Purpose: True when a name belongs to the workbook rather than to one sheet.
Private Function IsWorkbookLevelName(ByVal nmCheck As Name) As Boolean
    IsWorkbookLevelName = (TypeName(nmCheck.Parent) = "Workbook")
End Function


' Purpose: True when a RefersTo string is a LAMBDA definition.
Private Function IsLambdaDefinition(ByVal strRefers As String) As Boolean
    IsLambdaDefinition = (InStr(1, Replace(strRefers, " ", vbNullString), "=LAMBDA(", vbTextCompare) = 1)
End Function


' Purpose: True when a non-LAMBDA RefersTo holds no cell, sheet or external reference - a stored
'          value that means the same in any workbook (=0.05, ="Red", ={1,2,3}, =SEQUENCE(9)).
'          A "!" outside quoted text marks a sheet reference, "[" an external workbook, and
'          #REF! a broken one. Text in quotes is ignored, so ="Hi!" is still a value.
Private Function IsStoredValue(ByVal strRefers As String) As Boolean

    Dim strBare As String

    If Left$(strRefers, 1) <> "=" Then Exit Function
    strBare = StripStringLiterals(strRefers)
    IsStoredValue = (InStr(1, strBare, "!") = 0 And InStr(1, strBare, "[") = 0 _
                     And InStr(1, strBare, "#REF", vbTextCompare) = 0)
End Function


' Purpose: a name's comment, or "" where Excel will not give one up.
Private Function SafeComment(ByVal nmItem As Name) As String
    On Error Resume Next                          ' narrow: a name's comment can be unreadable
    SafeComment = nmItem.Comment
    On Error GoTo 0
End Function


' Purpose: a name's RefersTo, or "" where Excel will not give one up.
Private Function SafeRefersTo(ByVal nmItem As Name) As String
    On Error Resume Next                          ' narrow: some names have no readable RefersTo
    SafeRefersTo = nmItem.RefersTo
    On Error GoTo 0
End Function

'function to clear lambdas to only those in lambda table
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Clear Lambdas
' Description:            Clear lambdas that aren't in Lambdas table - for Comp Bot maintenance
' Macro Expression:       modLambdas.ClearLambdas()
' Generated:              01/10/2025 11:00 PM
'----------------------------------------------------------------------------------------------------
' KEEPS the Lambdas table AND every lambda the COMMANDS declare they need.
'
' Jaq, 2026-09-21: "that's a good shout, and more robust than just adding those lambdas
' in to the table, in case we forget something!"
'
' The Lambdas table is the TOOLBOX - what a competitor should remember. It was never the
' whole library. Nine lambdas are command plumbing that nobody types: IFBLANK, LkpRC,
' LkpRCByRow, FilterArray_byErikOehm, MM_byHaDang, RightAlignedArray_byJaqKennedy,
' DiffByRow_byJaqKennedy, ARROWSHIFT_byEmilieWilliams and Extract_byJaqKennedy. The old
' version deleted every one of them, silently breaking six commands. Reading the
' collection's OWN declarations instead of a hand-kept list means it cannot drift the
' next time a command gains a dependency.
'
' THREE OTHER FAULTS FIXED while in here:
'   - Range("Lambdas") was unqualified, so it bound to whatever sheet was active.
'   - It DELETED WHILE ITERATING WB.Names, which skips entries in VBA - so a run could
'     leave behind lambdas it meant to remove, and you would not know.
'   - No error handler and no VBAInit/VBAFin, so a failure left calculation manual, and
'     it reported nothing. A destructive command should say what it did.
Sub ClearLambdas()

    ' --- CONSTANTS (local to function) ---
    Const TABLE_NAME As String = "Lambdas"
    Const NAME_COL   As String = "Name"

    Dim wb As Workbook
    Dim lob As ListObject
    Dim dicKeep As Object
    Dim nmLambda As Name
    Dim cel As Range
    Dim varNames As Variant
    Dim lngIndex As Long
    Dim lngFound As Long
    Dim lngRemoved As Long
    Dim lngKept As Long
    Dim strName As String

    On Error GoTo ErrHandler
    VBAInit

    ' THISWORKBOOK, NOT ACTIVEWORKBOOK. This one is CompBot maintenance: it clears the
    ' lambdas in CompBot that are neither in the Lambdas table nor needed by a command.
    ' ClearUnusedLambdas is the one that works on whatever workbook is in front of you.
    '
    ' On ActiveWorkbook this was a loaded gun: run from the palette with a case file or
    ' Lambda Review in front, it read CompBot's table and then deleted 51 defined names
    ' out of the OTHER workbook. Measured 2026-09-21, by previewing it with Lambda
    ' Review active.
    Set wb = ThisWorkbook
    Set dicKeep = CreateObject("Scripting.Dictionary")
    dicKeep.CompareMode = 1                 ' vbTextCompare - Excel names ignore case

    ' 1. Everything the Lambdas table lists. Qualified to ThisWorkbook, because the
    '    table lives in CompBot however the command was launched.
    '
    '    FOUND BY THE TABLE LOAD, 2026-09-21: a ListObject's name does NOT live in
    '    Workbook.Names - tables have their own namespace - so
    '    Names("Lambdas").RefersToRange.ListObject raises run-time error 1004 and the
    '    whole command falls straight into its error handler. The previous line was
    '    written to fix an UNQUALIFIED Range("Lambdas"), which was a real bug, but the
    '    replacement never resolved either. Look the table up by name instead.
    Set lob = FindLambdasTable(ThisWorkbook, TABLE_NAME)
    If lob Is Nothing Then
        NoteError "ClearLambdas", 0, "no ListObject called '" & TABLE_NAME & "'"
        GoTo Cleanup
    End If
    For Each cel In lob.ListColumns(NAME_COL).DataBodyRange
        strName = Trim$(CStr(cel.Value2))
        If Len(strName) > 0 Then dicKeep(strName) = True
    Next cel

    ' 2. Everything the COMMANDS declare. Every lambda in the collection definition is
    '    written "<name>.lambda", both in each command's FormulaDependencies and in the
    '    Texts list, so harvesting that one token catches the lot without parsing JSON.
    AddDeclaredLambdas dicKeep

    ' 2b. ...and everything THOSE lambdas call, however deep. A declaration names the
    '     lambda a command invokes, not the helpers that lambda stands on, so without
    '     this a chain breaks at its second link. Jaq, 2026-09-21, on being shown the
    '     hazard: "that definitely needs a chain, but one which doesn't go into an
    '     infinite loop when a lambda is self-referential."
    ExpandKeepChain dicKeep, wb

    ' 3. Snapshot the names BEFORE deleting any - see the note above about iterating a
    '    collection you are removing from.
    ReDim varNames(1 To wb.Names.Count + 1)
    For Each nmLambda In wb.Names
        strName = vbNullString
        On Error Resume Next
        If InStr(1, nmLambda.RefersTo, "LAMBDA(", vbTextCompare) > 0 Then
            strName = nmLambda.Name
        End If
        On Error GoTo ErrHandler
        If Len(strName) > 0 Then
            lngFound = lngFound + 1
            varNames(lngFound) = strName
        End If
    Next nmLambda

    For lngIndex = 1 To lngFound
        strName = CStr(varNames(lngIndex))
        If dicKeep.Exists(strName) Then
            lngKept = lngKept + 1
        Else
            On Error Resume Next
            wb.Names(strName).Delete
            If Err.Number = 0 Then lngRemoved = lngRemoved + 1
            Err.Clear
            On Error GoTo ErrHandler
        End If
    Next lngIndex

Cleanup:
    VBAFin
    Application.StatusBar = "Clear Lambdas: kept " & lngKept & ", removed " & _
                            lngRemoved & " (Lambdas table + every lambda the commands " & _
                            "declare + " & m_lngChainAdded & " reached through the " & _
                            "call chain)."
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ClearLambdas", Err.Number, Err.Description
    Resume Cleanup
End Sub

' Purpose: report exactly what ClearLambdas WOULD keep and remove, without removing it.
'
'          This deliberately calls the SAME two helpers ClearLambdas calls, in the same
'          order, and reads the same Lambdas table. A preview computed by a separate
'          route is a preview that can disagree with the thing it is previewing, which
'          is worse than none - you would sign off a list that was never the real one.
'
'          Writes two tab-separated sections to strOutPath: KEEP and REMOVE, each with
'          the reason the name is in that section.
Public Function PreviewClearLambdas(ByVal strOutPath As String) As String

    ' --- CONSTANTS (local to function) ---
    Const TABLE_NAME   As String = "Lambdas"
    Const NAME_COL     As String = "Name"
    Const ADTYPESTRING As Long = 2
    Const ADSAVECREATEOVERWRITE As Long = 2

    Dim wb As Workbook
    Dim lob As ListObject
    Dim dicKeep As Object
    Dim dicTable As Object
    Dim dicDeclared As Object
    Dim nmLambda As Name
    Dim cel As Range
    Dim objOut As Object
    Dim varKey As Variant
    Dim strName As String
    Dim strWhy As String
    Dim strKeep As String
    Dim strGo As String
    Dim lngKept As Long
    Dim lngGo As Long

    On Error GoTo PreviewErr

    ' Same target as ClearLambdas, or the preview previews the wrong workbook.
    Set wb = ThisWorkbook
    Set dicKeep = CreateObject("Scripting.Dictionary")
    dicKeep.CompareMode = 1
    Set dicTable = CreateObject("Scripting.Dictionary")
    dicTable.CompareMode = 1
    Set dicDeclared = CreateObject("Scripting.Dictionary")
    dicDeclared.CompareMode = 1

    Set lob = FindLambdasTable(ThisWorkbook, TABLE_NAME)
    If lob Is Nothing Then
        PreviewClearLambdas = "ERROR no ListObject called '" & TABLE_NAME & "'"
        Exit Function
    End If
    For Each cel In lob.ListColumns(NAME_COL).DataBodyRange
        strName = Trim$(CStr(cel.Value2))
        If Len(strName) > 0 Then
            dicKeep(strName) = True
            dicTable(strName) = True
        End If
    Next cel

    AddDeclaredLambdas dicKeep
    For Each varKey In dicKeep.Keys
        If Not dicTable.Exists(CStr(varKey)) Then dicDeclared(CStr(varKey)) = True
    Next varKey

    ExpandKeepChain dicKeep, wb

    For Each nmLambda In wb.Names
        strName = vbNullString
        On Error Resume Next
        If InStr(1, nmLambda.RefersTo, "LAMBDA(", vbTextCompare) > 0 Then
            strName = nmLambda.Name
        End If
        On Error GoTo PreviewErr

        If Len(strName) > 0 Then
            If dicKeep.Exists(strName) Then
                If dicTable.Exists(strName) Then
                    strWhy = "in the Lambdas table (a pick)"
                ElseIf dicDeclared.Exists(strName) Then
                    strWhy = "declared by a command"
                Else
                    strWhy = "reached through the call chain"
                End If
                lngKept = lngKept + 1
                strKeep = strKeep & "KEEP" & vbTab & strName & vbTab & strWhy & vbLf
            Else
                lngGo = lngGo + 1
                strGo = strGo & "REMOVE" & vbTab & strName & vbTab & _
                        "not a pick, not declared, not reached" & vbLf
            End If
        End If
    Next nmLambda

    Set objOut = CreateObject("ADODB.Stream")
    objOut.Type = ADTYPESTRING
    objOut.Charset = "utf-8"
    objOut.Open
    objOut.WriteText strKeep & strGo
    objOut.SaveToFile strOutPath, ADSAVECREATEOVERWRITE
    objOut.Close

    PreviewClearLambdas = "would KEEP " & lngKept & ", would REMOVE " & lngGo & _
                          " (table " & dicTable.Count & ", declared " & _
                          dicDeclared.Count & ", chain " & m_lngChainAdded & ")"
    Exit Function

PreviewErr:
    PreviewClearLambdas = "ERROR " & Err.Number & " " & Err.Description
End Function


' Purpose: find a ListObject by name anywhere in a workbook.
'
'          A table's name is not in Workbook.Names, so it cannot be reached that way.
'          Scanning the sheets is reliable and costs nothing at this size.
Private Function FindLambdasTable(ByVal wbTarget As Workbook, _
                                  ByVal strName As String) As ListObject

    Dim ws As Worksheet
    Dim lob As ListObject

    For Each ws In wbTarget.Worksheets
        For Each lob In ws.ListObjects
            If StrComp(lob.Name, strName, vbTextCompare) = 0 Then
                Set FindLambdasTable = lob
                Exit Function
            End If
        Next lob
    Next ws
End Function


' Purpose: add every lambda the COLLECTION declares to the keep-set.
'
'          The collection definition is stored as a custom XML part on CompBot. Every
'          lambda in it is referenced as "<name>.lambda" - in each command's
'          FormulaDependencies and again in the Texts list - so scanning for that suffix
'          collects the whole declared set without needing a JSON parser in VBA.
'
'          Silent by design: if the part cannot be read, the table alone still governs,
'          which is the old behaviour rather than a crash.
Private Sub AddDeclaredLambdas(ByRef dicKeep As Object)

    ' --- CONSTANTS (local to function) ---
    Const SUFFIX As String = ".lambda"

    Dim objPart As Object
    Dim strXml As String
    Dim lngAt As Long
    Dim lngStart As Long
    Dim strName As String

    On Error Resume Next
    For Each objPart In ThisWorkbook.CustomXMLParts
        strXml = vbNullString
        strXml = objPart.XML
        If InStr(1, strXml, SUFFIX, vbTextCompare) > 0 Then
            lngAt = InStr(1, strXml, SUFFIX, vbTextCompare)
            Do While lngAt > 0
                ' Walk back to the opening quote of the name.
                lngStart = InStrRev(strXml, """", lngAt)
                If lngStart > 0 And lngAt - lngStart - 1 > 0 Then
                    strName = Mid$(strXml, lngStart + 1, lngAt - lngStart - 1)
                    ' Guard against picking up an escaped or malformed fragment.
                    If Len(strName) > 0 And InStr(1, strName, "<") = 0 Then
                        dicKeep(strName) = True
                    End If
                End If
                lngAt = InStr(lngAt + Len(SUFFIX), strXml, SUFFIX, vbTextCompare)
            Loop
        End If
    Next objPart
    On Error GoTo 0
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           List Lambdas Used
' Description:            List lambdas used.
' Macro Expression:       modLambdas.ListLambdasUsed()
' Generated:              2026-08-22 08:49 AM
'----------------------------------------------------------------------------------------------------
Public Sub ListLambdasUsed()
    On Error GoTo ErrHandler
    VBAInit

    Dim wbTarget As Workbook
    Set wbTarget = ActiveWorkbook
    If wbTarget Is Nothing Then GoTo Cleanup

    BuildUsageReport wbTarget

Cleanup:
    VBAFin
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ListLambdasUsed", Err.Number, Err.Description
    Resume Cleanup
End Sub

Public Sub BuildUsageReport(ByVal wbTarget As Workbook)
    On Error GoTo ErrHandler

    m_strScanErrors = vbNullString

    Dim dicDefined  As Object
    Dim dicCallable As Object
    Dim dicUsage    As Object

    Set dicDefined = GetDefinedLambdas(wbTarget)
    Set dicCallable = GetCallableLambdas()
    Set dicUsage = GetLambdaUsage(wbTarget, dicCallable)

    WriteUsageReport wbTarget, dicDefined, dicCallable, dicUsage

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "BuildUsageReport", Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: Collect workbook-level defined names whose definition is a LAMBDA.
'          Key = normalised name, Item = the RefersTo string.
Private Function GetDefinedLambdas(ByVal wbTarget As Workbook) As Object

    Dim dicOut      As Object
    Dim nmItem      As Name
    Dim strRefers   As String

    On Error GoTo ErrHandler

    Set dicOut = CreateObject("Scripting.Dictionary")
    dicOut.CompareMode = vbTextCompare

    For Each nmItem In wbTarget.Names
        strRefers = vbNullString
        On Error Resume Next                    ' narrow probe: some names have no readable RefersTo
        strRefers = nmItem.RefersTo
        On Error GoTo ErrHandler
        If InStr(1, Replace(strRefers, " ", vbNullString), "=LAMBDA(", vbTextCompare) = 1 Then
            dicOut(NormaliseName(nmItem.Name)) = strRefers
        End If
    Next nmItem

    Set GetDefinedLambdas = dicOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "GetDefinedLambdas", Err.Number, Err.Description
    Set GetDefinedLambdas = dicOut
End Function


' Purpose: Scan every formula cell for non-built-in function calls, then optionally walk
'          the definitions of the LAMBDAs found to pick up nested calls.
'          Item is a |-delimited record: cells|instances|sheets|firstSeen|via
Private Function GetLambdaUsage(ByVal wbTarget As Workbook, ByVal dicDefined As Object) As Object

    ' --- CONSTANTS (local to function) ---
    Const REC_CELLS     As Long = 0
    Const REC_INSTANCES As Long = 1
    Const REC_SHEETS    As Long = 2
    Const REC_FIRST     As Long = 3
    Const REC_VIA       As Long = 4

    Dim dicOut          As Object
    Dim dicCallable     As Object
    Dim dicTokens       As Object
    Dim wsScan          As Worksheet
    Dim rngFormulas     As Range
    Dim rngArea         As Range
    Dim rngCell         As Range
    Dim varRec          As Variant
    Dim varToken        As Variant
    Dim varKey          As Variant
    Dim strToken        As String
    Dim strWhere        As String
    Dim blnGrew         As Boolean

    On Error GoTo ErrHandler

    Set dicOut = CreateObject("Scripting.Dictionary")
    dicOut.CompareMode = vbTextCompare
    Set dicCallable = GetCallableLambdas()

    ' --- Pass 1: formula cells ---
    For Each wsScan In wbTarget.Worksheets
        If StrComp(wsScan.Name, m_REPORT_SHEET, vbTextCompare) <> 0 Then

            Set rngFormulas = Nothing
            On Error Resume Next                ' narrow probe: no formulas on the sheet
            Set rngFormulas = wsScan.UsedRange.SpecialCells(xlCellTypeFormulas)
            On Error GoTo ErrHandler

            If Not rngFormulas Is Nothing Then
                For Each rngArea In rngFormulas.Areas
                    For Each rngCell In rngArea.Cells

                        Set dicTokens = GetFunctionTokens(rngCell.Formula2)

                        For Each varToken In dicTokens.Keys
                            strToken = CStr(varToken)
                            If dicCallable.Exists(strToken) Then

                                strWhere = wsScan.Name & "!" & rngCell.Address(False, False)

                                If dicOut.Exists(strToken) Then
                                    varRec = Split(dicOut(strToken), "|")
                                    varRec(REC_CELLS) = CStr(CLng(varRec(REC_CELLS)) + 1)
                                    varRec(REC_INSTANCES) = CStr(CLng(varRec(REC_INSTANCES)) + dicTokens(varToken))
                                    If InStr(1, varRec(REC_SHEETS), wsScan.Name, vbTextCompare) = 0 Then
                                        varRec(REC_SHEETS) = varRec(REC_SHEETS) & ", " & wsScan.Name
                                    End If
                                    dicOut(strToken) = Join(varRec, "|")
                                Else
                                    dicOut(strToken) = Join(Array("1", CStr(dicTokens(varToken)), _
                                                                  wsScan.Name, strWhere, "Sheet"), "|")
                                End If
                            End If
                        Next varToken
                    Next rngCell
                Next rngArea
            End If
        End If
    Next wsScan

    ' --- Pass 2: nested calls inside the definitions of LAMBDAs already in use ---
    If m_INCLUDE_NESTED Then
        Do
            blnGrew = False
            For Each varKey In dicDefined.Keys
                If dicOut.Exists(CStr(varKey)) Then
                    Set dicTokens = GetFunctionTokens(CStr(dicDefined(varKey)))
                    For Each varToken In dicTokens.Keys
                        strToken = CStr(varToken)
                        If dicCallable.Exists(strToken) Then
                            If StrComp(strToken, CStr(varKey), vbTextCompare) <> 0 Then
                                If Not dicOut.Exists(strToken) Then
                                    dicOut(strToken) = Join(Array("0", "0", vbNullString, _
                                                                  "in " & CStr(varKey), "Nested"), "|")
                                    blnGrew = True
                                End If
                            End If
                        End If
                    Next varToken
                End If
            Next varKey
        Loop While blnGrew
    End If

    Set GetLambdaUsage = dicOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    m_strScanErrors = m_strScanErrors & "Scan failed: " & Err.Description & ". "
    NoteError "GetLambdaUsage", Err.Number, Err.Description
    Set GetLambdaUsage = dicOut
End Function


' Purpose: grow the keep-set to every lambda the kept ones CALL, at any depth.
'
'          WHY: a command declares what IT invokes. "Lookup Value by Row" happens to
'          list its whole chain, but that is a hand-written list and the next command
'          will not. Keeping only the declared names deletes the helpers underneath
'          them, and the command then fails with #NAME? on a lambda nobody touched.
'
'          TERMINATION, which is the whole difficulty. A recursive lambda names ITSELF,
'          and two lambdas can name each other, so walking the call graph naively never
'          returns. dicKeep IS the visited set: a name is queued only at the moment it
'          is first ADDED to it, so a self-reference finds itself already present and is
'          not re-queued, and a mutual pair settles on the second visit. Every name
'          therefore enters the queue at most once and the walk is bounded by the number
'          of lambdas in the workbook. MAX_STEPS is a second, independent guard: if the
'          reasoning above is ever wrong, Excel stops rather than hangs.
Private Sub ExpandKeepChain(ByRef dicKeep As Object, ByVal wbTarget As Workbook)

    ' --- CONSTANTS (local to function) ---
    Const MAX_STEPS As Long = 100000

    Dim dicBodies As Object
    Dim dicTokens As Object
    Dim colQueue As Collection
    Dim nmLambda As Name
    Dim varKey As Variant
    Dim varToken As Variant
    Dim strName As String
    Dim strBody As String
    Dim lngSteps As Long
    Dim lngAdded As Long

    On Error GoTo ErrHandler

    Set dicBodies = CreateObject("Scripting.Dictionary")
    dicBodies.CompareMode = vbTextCompare

    ' One pass over the names, so the walk below is pure dictionary work.
    For Each nmLambda In wbTarget.Names
        strBody = vbNullString
        On Error Resume Next
        strBody = nmLambda.RefersTo
        strName = nmLambda.Name
        On Error GoTo ErrHandler
        If Len(strBody) > 0 Then
            If InStr(1, strBody, "LAMBDA(", vbTextCompare) > 0 Then
                dicBodies(NormaliseName(strName)) = strBody
            End If
        End If
    Next nmLambda

    ' Seed with what is already kept and actually exists here.
    Set colQueue = New Collection
    For Each varKey In dicKeep.Keys
        If dicBodies.Exists(CStr(varKey)) Then colQueue.Add CStr(varKey)
    Next varKey

    Do While colQueue.Count > 0
        lngSteps = lngSteps + 1
        If lngSteps > MAX_STEPS Then
            NoteError "ExpandKeepChain", 0, _
                      "walk exceeded " & MAX_STEPS & " steps - stopped early"
            Exit Do
        End If

        strName = colQueue(1)
        colQueue.Remove 1

        Set dicTokens = GetFunctionTokens(CStr(dicBodies(strName)))
        For Each varToken In dicTokens.Keys
            ' Only a name defined HERE can be a lambda call; everything else the
            ' tokeniser returns is a native function such as LET or FILTER.
            If dicBodies.Exists(CStr(varToken)) Then
                If Not dicKeep.Exists(CStr(varToken)) Then
                    dicKeep(CStr(varToken)) = True
                    colQueue.Add CStr(varToken)
                    lngAdded = lngAdded + 1
                End If
            End If
        Next varToken
    Loop

    m_lngChainAdded = lngAdded
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ExpandKeepChain", Err.Number, Err.Description
End Sub


' Purpose: Return a dictionary of FunctionName -> call count found in one formula.
Private Function GetFunctionTokens(ByVal strFormula As String) As Object

    ' --- CONSTANTS (local to function) ---
    Const TOKEN_PATTERN As String = "([A-Za-z_\\][A-Za-z0-9_.\\]*)\s*\("

    Dim dicOut      As Object
    Dim objRegEx    As Object
    Dim objMatch    As Object
    Dim strName     As String

    On Error GoTo ErrHandler

    Set dicOut = CreateObject("Scripting.Dictionary")
    dicOut.CompareMode = vbTextCompare
    If Len(strFormula) = 0 Then GoTo Cleanup

    Set objRegEx = CreateObject("VBScript.RegExp")
    objRegEx.Global = True
    objRegEx.IgnoreCase = True
    objRegEx.Pattern = TOKEN_PATTERN

    For Each objMatch In objRegEx.Execute(StripStringLiterals(strFormula))
        strName = NormaliseName(objMatch.SubMatches(0))
        If Len(strName) > 0 Then dicOut(strName) = CLng(dicOut(strName)) + 1
    Next objMatch

Cleanup:
    Set GetFunctionTokens = dicOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "GetFunctionTokens", Err.Number, Err.Description
    Resume Cleanup
End Function


' Purpose: Blank out "quoted text" so a function name inside a string literal is not counted.
Private Function StripStringLiterals(ByVal strFormula As String) As String

    Dim lngPos      As Long
    Dim blnInQuote  As Boolean
    Dim strChar     As String
    Dim strOut      As String

    For lngPos = 1 To Len(strFormula)
        strChar = Mid$(strFormula, lngPos, 1)
        If strChar = """" Then
            blnInQuote = Not blnInQuote
            strOut = strOut & " "
        ElseIf blnInQuote Then
            strOut = strOut & " "
        Else
            strOut = strOut & strChar
        End If
    Next lngPos

    StripStringLiterals = strOut
End Function


' Purpose: Strip Excel's internal function prefixes and any leading comment backslashes.
Private Function NormaliseName(ByVal strName As String) As String

    Dim varPrefixes As Variant
    Dim varPrefix   As Variant

    varPrefixes = Array("_xlfn.", "_xlws.", "_xleta.", "_xludf.", "_xlpm.")

    For Each varPrefix In varPrefixes
        If InStr(1, strName, CStr(varPrefix), vbTextCompare) = 1 Then
            strName = Mid$(strName, Len(CStr(varPrefix)) + 1)
        End If
    Next varPrefix

    Do While Left$(strName, 1) = "\"
        strName = Mid$(strName, 2)
    Loop

    NormaliseName = strName
End Function


' Purpose: Build (or refresh) the report sheet and write the results in a single assignment.
Private Sub WriteUsageReport(ByVal wbTarget As Workbook, ByVal dicDefined As Object, _
                             ByVal dicCallable As Object, ByVal dicUsage As Object)

    ' --- CONSTANTS (local to function) ---
    Const HEADER_GAP As Long = 3          ' rows between the title block and the header row

    Dim wsReport    As Worksheet
    Dim rngAnchor   As Range
    Dim rngHeader   As Range
    Dim rngBody     As Range
    Dim loReport    As ListObject
    Dim varHeaders  As Variant
    Dim varKeys     As Variant
    Dim varOut      As Variant
    Dim varRec      As Variant
    Dim lngRow      As Long
    Dim lngCols     As Long
    Dim strMeta     As String

    On Error GoTo ErrHandler

    varHeaders = Array("LAMBDA", "Arguments", "Volatile", "Via", "Cells Using", "Instances", "First Seen")
    lngCols = UBound(varHeaders) - LBound(varHeaders) + 1

    Set wsReport = GetSheetByName(wbTarget, m_REPORT_SHEET)
    Set rngAnchor = wsReport.Cells(m_ANCHOR_ROW, m_ANCHOR_COL)

    rngAnchor.Value2 = "LAMBDAs Used"
    rngAnchor.Font.Size = 14
    rngAnchor.Font.Bold = True

    strMeta = wbTarget.Name & "  |  " & Format$(Now, "dd mmm yyyy hh:mm") & _
              "  |  " & dicUsage.Count & " in use  |  " & dicDefined.Count & " defined here"
    If Len(m_strScanErrors) > 0 Then strMeta = strMeta & "  |  INCOMPLETE: " & m_strScanErrors
    rngAnchor.Offset(1, 0).Value2 = strMeta
    rngAnchor.Offset(1, 0).Font.Italic = True
    If Len(m_strScanErrors) > 0 Then rngAnchor.Offset(1, 0).Font.Color = vbRed

    Set rngHeader = rngAnchor.Offset(HEADER_GAP, 0).Resize(1, lngCols)
    rngHeader.Value2 = varHeaders

    If dicUsage.Count = 0 Then
        rngHeader.Offset(1, 0).Cells(1, 1).Value2 = "No LAMBDAs found in use."
        GoTo Cleanup
    End If

    varKeys = SortKeys(dicUsage)
    ReDim varOut(1 To dicUsage.Count, 1 To lngCols)

    For lngRow = 1 To dicUsage.Count
        varRec = Split(dicUsage(varKeys(lngRow - 1)), "|")
        varOut(lngRow, 1) = varKeys(lngRow - 1)
        varOut(lngRow, 2) = GetLambdaArguments(CStr(dicCallable(varKeys(lngRow - 1))))
        varOut(lngRow, 3) = GetVolatileFlag(CStr(dicCallable(varKeys(lngRow - 1))))
        varOut(lngRow, 4) = varRec(4)
        varOut(lngRow, 5) = CLng(varRec(0))
        varOut(lngRow, 6) = CLng(varRec(1))
        varOut(lngRow, 7) = varRec(3)
    Next lngRow

    Set rngBody = rngHeader.Offset(1, 0).Resize(dicUsage.Count, lngCols)
    rngBody.Value2 = varOut

    Set loReport = wsReport.ListObjects.Add(xlSrcRange, Union(rngHeader, rngBody), , xlYes)
    loReport.Name = m_REPORT_TABLE
    loReport.TableStyle = "TableStyleMedium1"
    loReport.ListColumns("Cells Using").DataBodyRange.NumberFormat = "#,##0"
    loReport.ListColumns("Instances").DataBodyRange.NumberFormat = "#,##0"

    ' Rows found only inside another LAMBDA's definition, never called from a cell.
    With loReport.ListColumns("Via").DataBodyRange.FormatConditions
        .Delete
        .Add(xlCellValue, xlEqual, "=""Nested""").Interior.Color = RGB(255, 242, 204)
    End With

Cleanup:
    If Not loReport Is Nothing Then
        loReport.Range.Columns.AutoFit
    Else
        wsReport.Columns(m_ANCHOR_COL).Resize(, lngCols).AutoFit
    End If
    Exit Sub
    
ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "WriteUsageReport", Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: Return the report sheet, cleared and stripped of any previous table, creating it if absent.
Private Function GetReportSheet(ByVal wbTarget As Workbook) As Worksheet

    Dim wsFound As Worksheet
    Dim loOld   As ListObject

    On Error GoTo ErrHandler

    On Error Resume Next                        ' narrow probe: does the sheet exist
    Set wsFound = wbTarget.Worksheets(m_REPORT_SHEET)
    On Error GoTo ErrHandler

    If wsFound Is Nothing Then
        Set wsFound = wbTarget.Worksheets.Add(After:=wbTarget.Worksheets(wbTarget.Worksheets.Count))
        wsFound.Name = m_REPORT_SHEET
    Else
        For Each loOld In wsFound.ListObjects
            loOld.Unlist
        Next loOld
        wsFound.Cells.Clear
        wsFound.Cells.FormatConditions.Delete
    End If

    Set GetReportSheet = wsFound
    HideGridlines wsFound
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "GetReportSheet", Err.Number, Err.Description
    Set GetReportSheet = wsFound
End Function


' Purpose: Dictionary keys, sorted case-insensitively. Insertion sort - the list is short.
Private Function SortKeys(ByVal dicSource As Object) As Variant

    Dim varKeys As Variant
    Dim varSwap As Variant
    Dim lngI    As Long
    Dim lngJ    As Long

    varKeys = dicSource.Keys

    For lngI = LBound(varKeys) To UBound(varKeys) - 1
        For lngJ = lngI + 1 To UBound(varKeys)
            If StrComp(varKeys(lngI), varKeys(lngJ), vbTextCompare) > 0 Then
                varSwap = varKeys(lngI)
                varKeys(lngI) = varKeys(lngJ)
                varKeys(lngJ) = varSwap
            End If
        Next lngJ
    Next lngI

    SortKeys = varKeys
End Function


' Purpose: Every LAMBDA name callable in this session - all open workbooks and add-ins.
'          Item = source workbook name. Replaces list-based built-in exclusion.
Private Function GetCallableLambdas() As Object

    Dim dicOut      As Object
    Dim wbOpen      As Workbook
    Dim nmItem      As Name
    Dim strRefers   As String

    On Error GoTo ErrHandler

    Set dicOut = CreateObject("Scripting.Dictionary")
    dicOut.CompareMode = vbTextCompare

    For Each wbOpen In Application.Workbooks
        For Each nmItem In wbOpen.Names
            strRefers = vbNullString
            On Error Resume Next
            strRefers = nmItem.RefersTo
            On Error GoTo ErrHandler
            If InStr(1, Replace(strRefers, " ", vbNullString), "=LAMBDA(", vbTextCompare) = 1 Then
                If Not dicOut.Exists(NormaliseName(nmItem.Name)) Then
                    dicOut(NormaliseName(nmItem.Name)) = strRefers
                End If
            End If
        Next nmItem
    Next wbOpen

    Set GetCallableLambdas = dicOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "GetCallableLambdas", Err.Number, Err.Description
    Set GetCallableLambdas = dicOut
End Function
' Purpose: Extract the parameter list from a LAMBDA definition, marking optionals with [].
'          Returns "" if the definition cannot be parsed.
Private Function GetLambdaArguments(ByVal strRefers As String) As String

    Dim lngStart    As Long
    Dim lngPos      As Long
    Dim lngDepth    As Long
    Dim strChar     As String
    Dim strArgs     As String
    Dim varParts    As Variant
    Dim lngI        As Long
    Dim strPart     As String
    Dim blnOptional As Boolean

    On Error GoTo ErrHandler

    strRefers = StripStringLiterals(strRefers)
    lngStart = InStr(1, strRefers, "LAMBDA(", vbTextCompare)
    If lngStart = 0 Then Exit Function
    lngStart = lngStart + Len("LAMBDA(")

    ' Walk to the last top-level comma - everything before it is a parameter,
    ' the final segment is the calculation body.
    lngDepth = 0
    For lngPos = lngStart To Len(strRefers)
        strChar = Mid$(strRefers, lngPos, 1)
        Select Case strChar
            Case "(":  lngDepth = lngDepth + 1
            Case ")":  If lngDepth = 0 Then Exit For
                       lngDepth = lngDepth - 1
            Case ",":  If lngDepth = 0 Then strArgs = Mid$(strRefers, lngStart, lngPos - lngStart)
        End Select
    Next lngPos

    If Len(strArgs) = 0 Then
        GetLambdaArguments = "(no arguments)"
        Exit Function
    End If

    ' Optionals are the trailing run - anything after the first one is also optional.
    varParts = Split(strArgs, ",")
    For lngI = UBound(varParts) To LBound(varParts) Step -1
        strPart = Trim$(Replace(Replace(varParts(lngI), Chr$(10), " "), Chr$(13), " "))
        strPart = NormaliseName(strPart)
        If Left$(strPart, 1) = "[" Then blnOptional = True
        If blnOptional And Left$(strPart, 1) <> "[" Then strPart = "[" & strPart & "]"
        varParts(lngI) = strPart
    Next lngI

    GetLambdaArguments = Join(varParts, ", ")
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "GetLambdaArguments", Err.Number, Err.Description
End Function


' Purpose: Flag a definition that calls a volatile function directly.
'          Inherited volatility - via another LAMBDA - is not detected here.
Private Function GetVolatileFlag(ByVal strRefers As String) As String

    Dim varVolatile As Variant
    Dim varFunc     As Variant
    Dim strBody     As String
    Dim strFound    As String

    varVolatile = Array("INDIRECT", "OFFSET", "CELL", "NOW", "TODAY", "RAND", _
                        "RANDBETWEEN", "RANDARRAY", "INFO")

    strBody = StripStringLiterals(strRefers)

    For Each varFunc In varVolatile
        If InStr(1, strBody, CStr(varFunc) & "(", vbTextCompare) > 0 Then
            strFound = strFound & IIf(Len(strFound) > 0, ", ", vbNullString) & CStr(varFunc)
        End If
    Next varFunc

    GetVolatileFlag = strFound
End Function

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Clear Unused Lambdas
' Description:            Clear unused lambdas.
' Macro Expression:       modLambdas.ClearUnusedLambdas()
' Generated:              2026-08-22 09:21 AM
'----------------------------------------------------------------------------------------------------
Public Sub ClearUnusedLambdas(Optional ByVal lngReport As Long = 0)
    On Error GoTo ErrHandler
    VBAInit

    Dim wbTarget    As Workbook
    Dim dicKeep     As Object
    Dim dicRemove   As Object
    Dim nmItem      As Name
    Dim lngI        As Long
    Dim strRefers   As String
    Dim strName     As String
    Dim strRefusal  As String

    Set wbTarget = ActiveWorkbook                 ' resolve once - focus can change mid-run
    If wbTarget Is Nothing Then GoTo Cleanup

    ' Never strip CompBot, the tutorial or the user's own lambda library: in those the
    ' 'unused' lambdas ARE the library. Jaq, 2026-09-24.
    If IsProtectedWorkbook(wbTarget, True) Then
        strRefusal = "Clear Unused Lambdas: refused. " & wbTarget.Name & _
                     " is CompBot or your lambda library, where every lambda is meant to stay."
        GoTo Cleanup
    End If

    ' Always rebuild: a stale table is the one way this deletes a live LAMBDA.
    BuildUsageReport wbTarget

    If Len(m_strScanErrors) > 0 Then
        NoteError "ClearUnusedLambdas", 0, "Usage scan incomplete - refusing to delete."
        GoTo Cleanup
    End If

    Set dicKeep = GetKeepList(wbTarget)
    If dicKeep Is Nothing Then GoTo Cleanup
    If dicKeep.Count = 0 Then
        NoteError "ClearUnusedLambdas", 0, "No LAMBDAs in use - refusing to delete everything."
        GoTo Cleanup
    End If

    Set dicRemove = CreateObject("Scripting.Dictionary")
    dicRemove.CompareMode = vbTextCompare

    ' Backwards: the Names collection re-indexes on delete and a forward loop skips entries.
    For lngI = wbTarget.Names.Count To 1 Step -1
        Set nmItem = wbTarget.Names(lngI)

        strRefers = vbNullString
        On Error Resume Next                      ' narrow probe: unreadable RefersTo
        strRefers = nmItem.RefersTo
        On Error GoTo ErrHandler

        If InStr(1, Replace(strRefers, " ", vbNullString), "=LAMBDA(", vbTextCompare) = 1 Then
            strName = NormaliseName(nmItem.Name)
            If Not dicKeep.Exists(strName) Then
                dicRemove(strName) = strRefers
                If Not m_PREVIEW_ONLY Then nmItem.Delete
            End If
        End If
    Next lngI

    If lngReport = 1 Then
        WriteRemovalReport wbTarget, dicRemove
    Else
        Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | ClearUnusedLambdas | " & _
                    IIf(m_PREVIEW_ONLY, "preview, ", "deleted ") & dicRemove.Count & " unused"
    End If

Cleanup:
    VBAFin
    If Len(strRefusal) > 0 Then Application.StatusBar = Left$(strRefusal, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "ClearUnusedLambdas", Err.Number, Err.Description
    Resume Cleanup
End Sub


' Purpose: Names to preserve, read from the Lambdas Used table. Includes the Nested rows,
'          so a LAMBDA called only by another live LAMBDA is kept.
'          Returns Nothing if the table is absent - never an empty keep-list by accident.
Private Function GetKeepList(ByVal wbTarget As Workbook) As Object

    Dim dicOut      As Object
    Dim loUsed      As ListObject
    Dim wsFind      As Worksheet
    Dim varData     As Variant
    Dim lngI        As Long

    On Error GoTo ErrHandler

    For Each wsFind In wbTarget.Worksheets
        On Error Resume Next                          ' narrow probe: table on this sheet
        Set loUsed = wsFind.ListObjects(m_REPORT_TABLE)
        On Error GoTo ErrHandler
        If Not loUsed Is Nothing Then Exit For
    Next wsFind

    If loUsed Is Nothing Then
        NoteError "GetKeepList", 0, "Table " & m_REPORT_TABLE & " not found - run ListLambdasUsed first."
        Exit Function
    End If

    Set dicOut = CreateObject("Scripting.Dictionary")
    dicOut.CompareMode = vbTextCompare

    If loUsed.ListRows.Count > 0 Then
        varData = loUsed.ListColumns("LAMBDA").DataBodyRange.Value2
        If IsArray(varData) Then
            For lngI = LBound(varData, 1) To UBound(varData, 1)
                If Len(Trim$(CStr(varData(lngI, 1)))) > 0 Then
                    dicOut(NormaliseName(Trim$(CStr(varData(lngI, 1))))) = True
                End If
            Next lngI
        Else
            dicOut(NormaliseName(Trim$(CStr(varData)))) = True
        End If
    End If

    Set GetKeepList = dicOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "GetKeepList", Err.Number, Err.Description
End Function


' Purpose: Record what was removed, with definitions, so a deletion can be undone by
'          pasting the RefersTo back into the Name Manager.
Private Sub WriteRemovalReport(ByVal wbTarget As Workbook, ByVal dicRemove As Object)

    ' --- CONSTANTS (local to function) ---
    Const HEADER_GAP As Long = 3

    Dim wsReport    As Worksheet
    Dim rngAnchor   As Range
    Dim rngHeader   As Range
    Dim rngBody     As Range
    Dim loReport    As ListObject
    Dim varHeaders  As Variant
    Dim varKeys     As Variant
    Dim varOut      As Variant
    Dim lngRow      As Long
    Dim lngCols     As Long
    Dim strMeta     As String
    Dim strAction   As String

    On Error GoTo ErrHandler

    varHeaders = Array("LAMBDA", "Action", "Definition")
    lngCols = UBound(varHeaders) - LBound(varHeaders) + 1
    strAction = IIf(m_PREVIEW_ONLY, "Would delete", "Deleted")

    Set wsReport = GetSheetByName(wbTarget, m_REMOVED_SHEET)
    Set rngAnchor = wsReport.Cells(m_ANCHOR_ROW, m_ANCHOR_COL)

    rngAnchor.Value2 = "LAMBDAs Removed"
    rngAnchor.Font.Size = 14
    rngAnchor.Font.Bold = True

    strMeta = IIf(m_PREVIEW_ONLY, "PREVIEW - nothing deleted. ", vbNullString) & _
              wbTarget.Name & "  |  " & Format$(Now, "dd mmm yyyy hh:mm") & _
              "  |  " & dicRemove.Count & " unused  |  definitions kept below for recovery"
    If Len(m_strScanErrors) > 0 Then strMeta = strMeta & "  |  INCOMPLETE: " & m_strScanErrors
    rngAnchor.Offset(1, 0).Value2 = strMeta
    rngAnchor.Offset(1, 0).Font.Italic = True
    If m_PREVIEW_ONLY Or Len(m_strScanErrors) > 0 Then rngAnchor.Offset(1, 0).Font.Color = vbRed

    Set rngHeader = rngAnchor.Offset(HEADER_GAP, 0).Resize(1, lngCols)
    rngHeader.Value2 = varHeaders

    If dicRemove.Count = 0 Then
        rngHeader.Offset(1, 0).Cells(1, 1).Value2 = "No unused LAMBDAs."
        GoTo Cleanup
    End If

    varKeys = SortKeys(dicRemove)
    ReDim varOut(1 To dicRemove.Count, 1 To lngCols)

    For lngRow = 1 To dicRemove.Count
        varOut(lngRow, 1) = varKeys(lngRow - 1)
        varOut(lngRow, 2) = strAction
        ' Leading apostrophe: the definition starts with "=" and must not be treated as a formula.
        varOut(lngRow, 3) = "'" & CStr(dicRemove(varKeys(lngRow - 1)))
    Next lngRow

    Set rngBody = rngHeader.Offset(1, 0).Resize(dicRemove.Count, lngCols)
    rngBody.Value2 = varOut

    Set loReport = wsReport.ListObjects.Add(xlSrcRange, Union(rngHeader, rngBody), , xlYes)
    loReport.Name = "tblLambdasRemoved"
    loReport.TableStyle = "TableStyleMedium1"

Cleanup:
    If Not loReport Is Nothing Then
        loReport.Range.Columns.AutoFit
        loReport.ListColumns("Definition").Range.ColumnWidth = 80
    End If
    HideGridlines wsReport
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "WriteRemovalReport", Err.Number, Err.Description
    Resume Cleanup
End Sub

' Purpose: Return a report sheet by name, cleared and stripped of any previous table,
'          creating it if absent.
Private Function GetSheetByName(ByVal wbTarget As Workbook, ByVal strSheetName As String) As Worksheet

    Dim wsFound As Worksheet
    Dim loOld   As ListObject

    On Error GoTo ErrHandler

    On Error Resume Next                        ' narrow probe: does the sheet exist
    Set wsFound = wbTarget.Worksheets(strSheetName)
    On Error GoTo ErrHandler

    If wsFound Is Nothing Then
        Set wsFound = wbTarget.Worksheets.Add(After:=wbTarget.Worksheets(wbTarget.Worksheets.Count))
        wsFound.Name = strSheetName
    Else
        For Each loOld In wsFound.ListObjects
            loOld.Unlist
        Next loOld
        wsFound.Cells.Clear
        wsFound.Cells.FormatConditions.Delete
    End If

    HideGridlines wsFound

    Set GetSheetByName = wsFound
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    NoteError "GetSheetByName", Err.Number, Err.Description
    Set GetSheetByName = wsFound
End Function

' Purpose: Record a failure so it surfaces on the report. No logger in this workbook by design.
'          Moved here from modUtilities (2026-09-17): m_strScanErrors is Private to this module, and the
'          report's INCOMPLETE flag and ClearUnusedLambdas' refusal to delete both depend on it being filled.
Private Sub NoteError(ByVal strProc As String, ByVal lngNumber As Long, ByVal strDescription As String)
    m_strScanErrors = m_strScanErrors & strProc & " (" & lngNumber & "): " & strDescription & ". "
    Debug.Print Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & strProc & " | " & _
                lngNumber & " | " & strDescription
End Sub














