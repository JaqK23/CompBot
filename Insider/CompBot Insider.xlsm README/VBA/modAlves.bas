Attribute VB_Name = "modAlves"
Option Explicit

'==============================================================================
' modAlves - CompBot Insider: the Alves Method, after Fredson Alves's workflow.
'
' Build a spill-based solve wide, on the example row, with room for each array;
' then NEST every spilling step into its anchor cell (in-cell arrays), PACK the
' single cells onto the example row, and the whole chain fills down with Save
' From Example like an ordinary row. UNPACK and UNNEST take one packed row back
' out to a scratch area, spills visible again, for debugging.
'
' Needs Excel's in-cell arrays (Microsoft 365 Insider Beta, Build 20520+) and
' Workbook.CompatibilityVersion = 3, which Nest sets on the case workbook.
' Probed 2026-10-11: compbot-insider_2026.10\PLAN-261011.md, section 1.
'   - a nested cell is written with Range.Formula2 = "={...}"; its .Value2 is
'     Empty in VBA, so nothing here reads a nested array through the cell;
'   - a plain reference to a nested cell yields the inner array; a trailing #
'     is unreliable on one, so Nest strips # and Unnest puts it back;
'   - Range.Cut moves a nested cell intact and its dependents follow.
'
' No selection needed (Jaq, 2026-10-11): every command runs from the ANSWER
' cell. The solve region is the answer plus everything it depends on, on the
' same sheet, at or right of the level's working start - the one rule Save From
' Example uses (modCaseSetup.ExampleWorkingStart), so the three agree.
' Status bar only, never a dialog. The CompBot modules this calls: modCaseSetup
' (HeaderRowFor, AnswerColumn, LevelGameRows, ExampleWorkingStart), modUtilities
' (VBAInit, VBAFin), modLog (LogError).
'==============================================================================

' --- MODULE CONSTANTS ---
Private Const m_DEBUG_MODE      As Boolean = False
Private Const m_MAX_STATUS      As Long = 250       ' Application.StatusBar rejects long strings
Private Const m_MIN_BUILD       As Long = 20520     ' first Excel build with in-cell arrays
Private Const m_COMPAT_NESTED   As Long = 3         ' Workbook.CompatibilityVersion that enables them
Private Const m_MAX_PRECEDENTS  As Long = 5000      ' a region bigger than this is not a solve
Private Const m_SCAN_ROWS       As Long = 200       ' how far below the example row the games can run
Private Const m_UNPACK_GAP      As Long = 1         ' empty columns between unpacked steps
Private Const m_ERR_SPILL       As Long = 2045      ' xlErrSpill, the #SPILL! error value

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Alves Nest
' Description:            Nests every spilling step of the solve behind the active answer cell into one cell each
' Macro Expression:       modAlves.Nest()
' Generated:              2026-10-11 (CompBot Insider; after Fredson Alves's workflow, the Alves Method)
'----------------------------------------------------------------------------------------------------
Public Sub Nest()

    Dim strStatus As String

    On Error GoTo ErrHandler
    VBAInit

    If Not StartCell(strStatus, "ALVES NEST") Then GoTo Cleanup
    NestCore ActiveCell, strStatus

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = "ALVES NEST FAILED: " & Err.Number & ": " & Err.Description
    LogError "Nest", Err.Number, Err.Description
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Alves Unnest
' Description:            The reverse of Alves Nest: every nested step behind the active answer cell spills again
' Macro Expression:       modAlves.Unnest()
'----------------------------------------------------------------------------------------------------
Public Sub Unnest()

    Dim strStatus As String

    On Error GoTo ErrHandler
    VBAInit

    If Not StartCell(strStatus, "ALVES UNNEST") Then GoTo Cleanup
    UnnestCore ActiveCell, strStatus

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = "ALVES UNNEST FAILED: " & Err.Number & ": " & Err.Description
    LogError "Unnest", Err.Number, Err.Description
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Alves Pack
' Description:            Moves the nested steps behind the active answer cell onto its row, in order, answer last
' Macro Expression:       modAlves.Pack()
'----------------------------------------------------------------------------------------------------
Public Sub Pack()

    Dim strStatus As String
    Dim rngNew As Range

    On Error GoTo ErrHandler
    VBAInit

    If Not StartCell(strStatus, "ALVES PACK") Then GoTo Cleanup
    If PackCore(ActiveCell, strStatus, rngNew) Then
        ' Landing on the packed answer is the point: Save From Example runs from it.
        rngNew.Select
    End If

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = "ALVES PACK FAILED: " & Err.Number & ": " & Err.Description
    LogError "Pack", Err.Number, Err.Description
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Alves Unpack
' Description:            Copies one packed row's steps to a scratch area with room for each array, still nested
' Macro Expression:       modAlves.Unpack()
'----------------------------------------------------------------------------------------------------
Public Sub Unpack()

    Dim strStatus As String
    Dim rngNew As Range

    On Error GoTo ErrHandler
    VBAInit

    If Not StartCell(strStatus, "ALVES UNPACK") Then GoTo Cleanup
    If UnpackCore(ActiveCell, strStatus, rngNew) Then rngNew.Select

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = "ALVES UNPACK FAILED: " & Err.Number & ": " & Err.Description
    LogError "Unpack", Err.Number, Err.Description
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Alves Nest and Pack
' Description:            Alves Nest then Alves Pack: the solve behind the active answer cell becomes one row
' Macro Expression:       modAlves.NestAndPack()
'----------------------------------------------------------------------------------------------------
Public Sub NestAndPack()

    Dim strStatus As String
    Dim strPack As String
    Dim rngNew As Range

    On Error GoTo ErrHandler
    VBAInit

    If Not StartCell(strStatus, "ALVES NEST AND PACK") Then GoTo Cleanup
    If Not NestCore(ActiveCell, strStatus) Then GoTo Cleanup
    If PackCore(ActiveCell, strPack, rngNew) Then rngNew.Select
    strStatus = strStatus & "  |  " & strPack

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = "ALVES NEST AND PACK FAILED: " & Err.Number & ": " & Err.Description
    LogError "NestAndPack", Err.Number, Err.Description
    Resume Cleanup
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Alves Unpack and Unnest
' Description:            Alves Unpack then Alves Unnest: one packed row laid out with every spill visible
' Macro Expression:       modAlves.UnpackAndUnnest()
'----------------------------------------------------------------------------------------------------
Public Sub UnpackAndUnnest()

    Dim strStatus As String
    Dim strUnnest As String
    Dim rngNew As Range

    On Error GoTo ErrHandler
    VBAInit

    If Not StartCell(strStatus, "ALVES UNPACK AND UNNEST") Then GoTo Cleanup
    If Not UnpackCore(ActiveCell, strStatus, rngNew) Then GoTo Cleanup
    UnnestCore rngNew, strUnnest
    strStatus = strStatus & "  |  " & strUnnest
    rngNew.Select

Cleanup:
    VBAFin
    Application.StatusBar = Left$(strStatus, m_MAX_STATUS)
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = "ALVES UNPACK AND UNNEST FAILED: " & Err.Number & ": " & Err.Description
    LogError "UnpackAndUnnest", Err.Number, Err.Description
    Resume Cleanup
End Sub

'==============================================================================
' The four cores. Each returns True on success and always fills strStatus.
'==============================================================================

' Purpose: Nest. Every spilling anchor in the solve region becomes ={X}; every formula that refers
'          to one of those anchors with a trailing # (in the region, and every dependent of the
'          anchor on the sheet, so an answer cell further left keeps working) loses the #. The #
'          is stripped first, while the graph is intact, then the anchors are wrapped. Sets the
'          workbook's CompatibilityVersion to 3 when it is lower, and says so.
Private Function NestCore(ByVal rngAnswer As Range, ByRef strStatus As String) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const CMD As String = "ALVES NEST"

    Dim ws          As Worksheet
    Dim rngRegion   As Range
    Dim rngDeps     As Range
    Dim rngCell     As Range
    Dim dicAnchors  As Object
    Dim varKey      As Variant
    Dim strF        As String
    Dim strNew      As String
    Dim lngStart    As Long
    Dim lngN        As Long
    Dim lngNested   As Long
    Dim lngRefs     As Long
    Dim lngCells    As Long
    Dim strWhy      As String
    Dim strCompat   As String

    On Error GoTo ErrHandler

    strWhy = FeatureMissing()
    If Len(strWhy) > 0 Then
        strStatus = CMD & " FAILED: " & strWhy
        GoTo Cleanup
    End If
    Set rngRegion = SolveRegion(rngAnswer, lngStart, strWhy)
    If rngRegion Is Nothing Then
        strStatus = CMD & " FAILED: " & strWhy
        GoTo Cleanup
    End If
    Set ws = rngAnswer.Worksheet

    ' The anchors to nest: region cells that spill.
    Set dicAnchors = CreateObject("Scripting.Dictionary")
    For Each rngCell In rngRegion.Cells
        If rngCell.HasSpill Then
            If rngCell.SpillParent.Address = rngCell.Address Then
                dicAnchors(rngCell.Address(False, False)) = True
            End If
        End If
    Next rngCell
    If dicAnchors.Count = 0 Then
        strStatus = "Alves Nest: nothing to nest - no spilling formula feeds " & rngAnswer.Address(False, False) & _
                    " (region " & rngRegion.Address(False, False) & ")."
        NestCore = True
        GoTo Cleanup
    End If

    strWhy = EnableNestedArrays(ws.Parent, strCompat)
    If Len(strWhy) > 0 Then
        strStatus = CMD & " FAILED: " & strWhy
        GoTo Cleanup
    End If

    ' Every formula that may refer to an anchor with #: the region plus each anchor's dependents,
    ' collected before anything changes.
    Set rngDeps = rngRegion
    For Each varKey In dicAnchors.Keys
        Set rngDeps = UnionDependents(rngDeps, ws.Range(varKey))
    Next varKey
    For Each rngCell In rngDeps.Cells
        If rngCell.HasFormula Then
            strF = rngCell.Formula2
            strNew = StripHash(strF, dicAnchors, ws.Name, lngN)
            If lngN > 0 Then
                rngCell.Formula2 = strNew
                lngRefs = lngRefs + lngN
                lngCells = lngCells + 1
            End If
        End If
    Next rngCell

    ' Wrap the anchors, then calculate: calculation is manual while a command runs, and Pack reads
    ' HasSpill, which only changes on a calculation.
    For Each varKey In dicAnchors.Keys
        Set rngCell = ws.Range(varKey)
        strF = rngCell.Formula2
        If Not IsNested(strF) Then
            rngCell.Formula2 = "={" & Mid$(strF, 2) & "}"
            lngNested = lngNested + 1
        End If
    Next varKey
    ws.Calculate

    strStatus = "Alves Nest: " & lngNested & " step(s) nested in " & rngRegion.Address(False, False) & _
                ", " & lngRefs & " reference(s) updated in " & lngCells & " cell(s)." & strCompat
    NestCore = True

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = CMD & " FAILED: " & Err.Number & ": " & Err.Description
    LogError "NestCore", Err.Number, Err.Description
    Resume Cleanup
End Function

' Purpose: Unnest. Every nested cell in the solve region becomes =X again (checked one at a time:
'          a step with no room to spill is put back and the command stops, saying which), then every
'          dependent of an unnested anchor that refers to it gets its # back.
Private Function UnnestCore(ByVal rngAnswer As Range, ByRef strStatus As String) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const CMD As String = "ALVES UNNEST"

    Dim ws          As Worksheet
    Dim rngRegion   As Range
    Dim rngDeps     As Range
    Dim rngCell     As Range
    Dim dicNested   As Object
    Dim varKey      As Variant
    Dim strF        As String
    Dim strNew      As String
    Dim lngStart    As Long
    Dim lngN        As Long
    Dim lngDone     As Long
    Dim lngRefs     As Long
    Dim lngCells    As Long
    Dim strWhy      As String

    On Error GoTo ErrHandler

    Set rngRegion = SolveRegion(rngAnswer, lngStart, strWhy)
    If rngRegion Is Nothing Then
        strStatus = CMD & " FAILED: " & strWhy
        GoTo Cleanup
    End If
    Set ws = rngAnswer.Worksheet

    Set dicNested = CreateObject("Scripting.Dictionary")
    For Each rngCell In rngRegion.Cells
        If IsNested(rngCell.Formula2) Then dicNested(rngCell.Address(False, False)) = True
    Next rngCell
    If dicNested.Count = 0 Then
        strStatus = "Alves Unnest: nothing to unnest - no nested step feeds " & rngAnswer.Address(False, False) & _
                    " (region " & rngRegion.Address(False, False) & ")."
        UnnestCore = True
        GoTo Cleanup
    End If

    ' Dependents first, while every reference still points at a single cell.
    Set rngDeps = rngRegion
    For Each varKey In dicNested.Keys
        Set rngDeps = UnionDependents(rngDeps, ws.Range(varKey))
    Next varKey

    ' Unnest one at a time, so a step with no room is caught and put back before the next.
    For Each varKey In dicNested.Keys
        Set rngCell = ws.Range(varKey)
        strF = rngCell.Formula2
        rngCell.Formula2 = "=" & Mid$(strF, 3, Len(strF) - 3)
        rngCell.Calculate
        If IsSpillError(rngCell) Then
            rngCell.Formula2 = strF
            strStatus = CMD & " FAILED: " & rngCell.Address(False, False) & " has no room to spill - clear the cells below and right of it. " & _
                        lngDone & " step(s) were unnested before it."
            GoTo Cleanup
        End If
        lngDone = lngDone + 1
    Next varKey

    ' Put the # back on every reference to an unnested anchor.
    For Each rngCell In rngDeps.Cells
        If rngCell.HasFormula Then
            strF = rngCell.Formula2
            strNew = AddHash(strF, dicNested, ws.Name, lngN)
            If lngN > 0 Then
                rngCell.Formula2 = strNew
                lngRefs = lngRefs + lngN
                lngCells = lngCells + 1
            End If
        End If
    Next rngCell
    ws.Calculate

    strStatus = "Alves Unnest: " & lngDone & " step(s) spill again in " & rngRegion.Address(False, False) & _
                ", " & lngRefs & " reference(s) updated in " & lngCells & " cell(s)."
    UnnestCore = True

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = CMD & " FAILED: " & Err.Number & ": " & Err.Description
    LogError "UnnestCore", Err.Number, Err.Description
    Resume Cleanup
End Function

' Purpose: Pack. After Nest every step is one cell. Move them (Range.Cut: dependents follow, the
'          moved cell's own references stay put) onto the answer's row from the working start, in
'          dependency order, answer last, so Save From Example can take the row. A target that holds
'          another step is moved aside first; a target holding anything else stops the command
'          before anything moves. rngNewAnswer returns where the answer ended up.
Private Function PackCore(ByVal rngAnswer As Range, ByRef strStatus As String, ByRef rngNewAnswer As Range) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const CMD As String = "ALVES PACK"

    Dim ws          As Worksheet
    Dim rngRegion   As Range
    Dim rngCell     As Range
    Dim rngSrc      As Range
    Dim rngTgt      As Range
    Dim varOrder    As Variant
    Dim dicWhere    As Object       ' original address -> current address, as cells move
    Dim dicAtCell   As Object       ' current address -> original address
    Dim lngStart    As Long
    Dim lngRow      As Long
    Dim lngI        As Long
    Dim lngStage    As Long
    Dim lngMoved    As Long
    Dim strWhy      As String
    Dim strOrig     As String
    Dim strTgt      As String
    Dim strAnswer   As String

    On Error GoTo ErrHandler

    Set rngRegion = SolveRegion(rngAnswer, lngStart, strWhy)
    If rngRegion Is Nothing Then
        strStatus = CMD & " FAILED: " & strWhy
        GoTo Cleanup
    End If
    Set ws = rngAnswer.Worksheet
    lngRow = rngAnswer.Row
    ' A Range object follows its cell when the cell is Cut, so the answer's ORIGINAL address is kept
    ' as text: the dictionaries are keyed by it (first run, 2026-10-11: Range failed on the moved address).
    strAnswer = rngAnswer.Address(False, False)

    For Each rngCell In rngRegion.Cells
        If rngCell.HasSpill Then
            strStatus = CMD & " FAILED: " & rngCell.Address(False, False) & " still spills - run Alves Nest first."
            GoTo Cleanup
        End If
    Next rngCell

    varOrder = DependencyOrder(rngRegion, rngAnswer)

    ' Every target must be empty or hold a step; checked before anything moves.
    Set dicWhere = CreateObject("Scripting.Dictionary")
    Set dicAtCell = CreateObject("Scripting.Dictionary")
    For lngI = LBound(varOrder) To UBound(varOrder)
        dicWhere(varOrder(lngI)) = varOrder(lngI)
        dicAtCell(varOrder(lngI)) = varOrder(lngI)
    Next lngI
    For lngI = LBound(varOrder) To UBound(varOrder)
        Set rngTgt = ws.Cells(lngRow, lngStart + lngI - LBound(varOrder))
        strTgt = rngTgt.Address(False, False)
        If Len(CStr(rngTgt.Formula)) > 0 And Not dicAtCell.Exists(strTgt) Then
            strStatus = CMD & " FAILED: nothing moved - " & strTgt & " is in the way of the packed row."
            GoTo Cleanup
        End If
    Next lngI

    ' Staging column for a step that sits on another step's target: right of everything on the row.
    lngStage = ws.Cells(lngRow, ws.Columns.Count).End(xlToLeft).Column + 2
    If lngStage < lngStart + (UBound(varOrder) - LBound(varOrder) + 1) + 1 Then
        lngStage = lngStart + (UBound(varOrder) - LBound(varOrder) + 1) + 1
    End If

    For lngI = LBound(varOrder) To UBound(varOrder)
        strOrig = varOrder(lngI)
        Set rngSrc = ws.Range(dicWhere(strOrig))
        Set rngTgt = ws.Cells(lngRow, lngStart + lngI - LBound(varOrder))
        strTgt = rngTgt.Address(False, False)
        If rngSrc.Address <> rngTgt.Address Then
            If dicAtCell.Exists(strTgt) Then
                ' Another step is on this target: park it to the right, it moves later.
                MoveStep ws, dicWhere, dicAtCell, dicAtCell(strTgt), ws.Cells(lngRow, lngStage).Address(False, False)
                lngStage = lngStage + 1
            End If
            MoveStep ws, dicWhere, dicAtCell, strOrig, strTgt
            lngMoved = lngMoved + 1
        End If
    Next lngI
    Application.CutCopyMode = False

    Set rngNewAnswer = ws.Range(dicWhere(strAnswer))
    strStatus = "Alves Pack: " & (UBound(varOrder) - LBound(varOrder) + 1) & " step(s) on row " & lngRow & " from " & _
                ws.Cells(lngRow, lngStart).Address(False, False) & " to " & rngNewAnswer.Address(False, False) & _
                " (" & lngMoved & " moved); answer in " & rngNewAnswer.Address(False, False) & " - run Save From Example there."
    PackCore = True

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    Application.CutCopyMode = False
    strStatus = CMD & " FAILED: " & Err.Number & ": " & Err.Description
    LogError "PackCore", Err.Number, Err.Description
    Resume Cleanup
End Function

' Purpose: Unpack. Copy one packed row's steps, in dependency order, to a scratch area on the same
'          row right of the sheet's used range, still nested, each given the width its array needs
'          (measured by unnesting it there for a moment) plus a gap, so Alves Unnest can run on the
'          copy with room. References between steps are rewritten to the copies; references to the
'          inputs are unchanged. The packed original is untouched. rngNewAnswer is the answer's copy.
Private Function UnpackCore(ByVal rngAnswer As Range, ByRef strStatus As String, ByRef rngNewAnswer As Range) As Boolean

    ' --- CONSTANTS (local to function) ---
    Const CMD As String = "ALVES UNPACK"

    Dim ws          As Worksheet
    Dim rngRegion   As Range
    Dim rngCell     As Range
    Dim rngNew      As Range
    Dim varOrder    As Variant
    Dim dicMap      As Object       ' original address -> copy address
    Dim lngStart    As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim lngFirst    As Long
    Dim lngI        As Long
    Dim lngN        As Long
    Dim lngWidth    As Long
    Dim strF        As String
    Dim strNew      As String
    Dim strWhy      As String

    On Error GoTo ErrHandler

    Set rngRegion = SolveRegion(rngAnswer, lngStart, strWhy)
    If rngRegion Is Nothing Then
        strStatus = CMD & " FAILED: " & strWhy
        GoTo Cleanup
    End If
    Set ws = rngAnswer.Worksheet
    lngRow = rngAnswer.Row
    For Each rngCell In rngRegion.Cells
        If rngCell.HasSpill Then
            strStatus = CMD & " FAILED: " & rngCell.Address(False, False) & " spills - this takes a PACKED row (nested steps)."
            GoTo Cleanup
        End If
    Next rngCell

    varOrder = DependencyOrder(rngRegion, rngAnswer)

    ' Scratch area: one empty column right of everything on the sheet, on the answer's row.
    With ws.UsedRange
        lngCol = .Column + .Columns.Count + 1
    End With
    lngFirst = lngCol

    Set dicMap = CreateObject("Scripting.Dictionary")
    For lngI = LBound(varOrder) To UBound(varOrder)
        Set rngCell = ws.Range(varOrder(lngI))
        Set rngNew = ws.Cells(lngRow, lngCol)
        strF = rngCell.Formula2
        strNew = RewriteRefs(strF, dicMap, ws.Name, lngN)
        dicMap(varOrder(lngI)) = rngNew.Address(False, False)
        lngWidth = 1
        If IsNested(strNew) Then
            ' Measure: unnest it here for a moment, read the spill width, nest it again.
            rngNew.Formula2 = "=" & Mid$(strNew, 3, Len(strNew) - 3)
            rngNew.Calculate
            If rngNew.HasSpill Then lngWidth = rngNew.SpillingToRange.Columns.Count
            rngNew.Formula2 = strNew
        Else
            rngNew.Formula2 = strNew
        End If
        rngNew.Calculate
        lngCol = lngCol + lngWidth + m_UNPACK_GAP
    Next lngI

    Set rngNewAnswer = ws.Range(dicMap(rngAnswer.Address(False, False)))
    strStatus = "Alves Unpack: " & (UBound(varOrder) - LBound(varOrder) + 1) & " step(s) copied to " & _
                ws.Cells(lngRow, lngFirst).Address(False, False) & ":" & rngNewAnswer.Address(False, False) & _
                " with room to spill; the packed row is untouched. Run Alves Unnest on " & rngNewAnswer.Address(False, False) & "."
    UnpackCore = True

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = CMD & " FAILED: " & Err.Number & ": " & Err.Description
    LogError "UnpackCore", Err.Number, Err.Description
    Resume Cleanup
End Function

'==============================================================================
' Finding the solve region
'==============================================================================

' Purpose: The active cell is a usable start: one cell, holding a formula. Fills strStatus when not.
Private Function StartCell(ByRef strStatus As String, ByVal strCmd As String) As Boolean
    On Error GoTo ErrHandler
    If TypeName(Selection) <> "Range" Then
        strStatus = strCmd & " FAILED: stand on your answer formula (one cell on the example row)."
    ElseIf Selection.Cells.CountLarge <> 1 Then
        strStatus = strCmd & " FAILED: select ONE cell - your answer formula on the example row."
    ElseIf Not ActiveCell.HasFormula Then
        strStatus = strCmd & " FAILED: " & ActiveCell.Address(False, False) & " is not a formula - stand on your answer formula."
    Else
        StartCell = True
    End If
    Exit Function
ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strStatus = strCmd & " FAILED: " & Err.Description
    StartCell = False
End Function

' Purpose: "" when this Excel has in-cell arrays, else why not. The build number alone is not enough:
'          Jaq's laptop (Build 20527, Beta Channel, 2026-10-11) passed it and had no FLATTEN, and setting
'          CompatibilityVersion = 3 there raised error 5. So the feature itself is probed: FLATTEN must
'          be a known function (Evaluate gives #NAME? otherwise). The rollout is gradual per machine.
Private Function FeatureMissing() As String

    ' --- CONSTANTS (local to function) ---
    Const ERR_NAME As Long = 2029     ' xlErrName, what Evaluate returns for an unknown function

    Dim varProbe As Variant

    On Error GoTo ErrHandler

    If Application.Build < m_MIN_BUILD Then
        FeatureMissing = "needs Excel Beta Channel Build " & m_MIN_BUILD & " or later with in-cell arrays (this is Build " & _
                         Application.Build & ")."
        Exit Function
    End If
    varProbe = Application.Evaluate("FLATTEN(1)")
    If IsError(varProbe) Then
        If CLng(varProbe) = ERR_NAME Then
            FeatureMissing = "in-cell arrays are not in this Excel yet (FLATTEN is unknown; Build " & Application.Build & _
                             ", the rollout is gradual) - try File > Account > Update Now."
        End If
    End If
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    FeatureMissing = "could not check this Excel for in-cell arrays: " & Err.Description
End Function

' Purpose: Switch the workbook to CompatibilityVersion 3 (in-cell arrays on) when it is lower. Returns ""
'          when it is 3 afterwards, else why not: an Excel without the feature rejects the write with
'          error 5 (Jaq's laptop, 2026-10-11), so the write is probed rather than trusted.
Private Function EnableNestedArrays(ByVal wb As Workbook, ByRef strNote As String) As String

    Dim lngErr As Long

    On Error Resume Next
    If wb.CompatibilityVersion < m_COMPAT_NESTED Then
        wb.CompatibilityVersion = m_COMPAT_NESTED
        lngErr = Err.Number
        Err.Clear
        If lngErr = 0 And wb.CompatibilityVersion = m_COMPAT_NESTED Then
            strNote = " Workbook compatibility version set to " & m_COMPAT_NESTED & " (in-cell arrays on)."
        Else
            EnableNestedArrays = "this Excel would not switch the workbook to compatibility version " & m_COMPAT_NESTED & _
                                 " (error " & lngErr & "): in-cell arrays are not available on it yet."
        End If
    End If
    On Error GoTo 0
End Function

' Purpose: The solve region for the answer cell: the answer itself plus every formula cell it depends
'          on, directly or through other steps, on the same sheet, at or right of the level's working
'          start (modCaseSetup.ExampleWorkingStart: right of the inputs, the rule Save From Example
'          uses). A spilled child among the precedents counts as its anchor. Nothing, with strWhy,
'          when the answer is not on a level's example row (no Answer header above it) or the region
'          is absurdly large. lngStartCol returns the working start. Never raises.
Private Function SolveRegion(ByVal rngAnswer As Range, ByRef lngStartCol As Long, ByRef strWhy As String) As Range

    Dim ws          As Worksheet
    Dim rngPrec     As Range
    Dim rngCell     As Range
    Dim rngAnchor   As Range
    Dim rngOut      As Range
    Dim colGames    As Collection
    Dim lngAnsCol   As Long
    Dim lngHdRow    As Long
    Dim lngStopRow  As Long
    Dim lngFirstRow As Long
    Dim lngLastRow  As Long

    On Error GoTo ErrHandler

    Set ws = rngAnswer.Worksheet
    If Not rngAnswer.HasFormula Then
        strWhy = rngAnswer.Address(False, False) & " is not a formula - stand on your answer formula."
        GoTo Cleanup
    End If

    ' Where the working starts: the inputs end by the rule Save From Example uses. The level's Answer
    ' header is the nearest row above holding one, up to m_SCAN_ROWS up: HeaderRowFor looks only four
    ' rows up, which serves the example row, but Unpack and Unnest run on a GAME row further down
    ' (found on the video take, 2026-10-11: row 9 failed with "no Answer header").
    lngAnsCol = 0
    lngStopRow = rngAnswer.Row - m_SCAN_ROWS
    If lngStopRow < 1 Then lngStopRow = 1
    For lngHdRow = rngAnswer.Row - 1 To lngStopRow Step -1
        lngAnsCol = AnswerColumn(ws, lngHdRow)
        If lngAnsCol > 0 Then Exit For
    Next lngHdRow
    If lngAnsCol = 0 Then
        strWhy = "no Answer header found above row " & rngAnswer.Row & " - stand on the answer formula of a level's example row."
        GoTo Cleanup
    End If
    Set colGames = LevelGameRows(ws, 2, rngAnswer.Row, rngAnswer.Row + m_SCAN_ROWS)
    If colGames.Count > 0 Then
        lngFirstRow = colGames(1)
        lngLastRow = colGames(colGames.Count)
    Else
        lngFirstRow = rngAnswer.Row + 1
        lngLastRow = rngAnswer.Row + 1
    End If
    lngStartCol = ExampleWorkingStart(ws, rngAnswer.Row, rngAnswer.Column, lngAnsCol, lngFirstRow, lngLastRow)

    Set rngOut = rngAnswer
    ' Probe: Precedents raises when there are none.
    On Error Resume Next
    Set rngPrec = rngAnswer.Precedents
    On Error GoTo ErrHandler
    If Not rngPrec Is Nothing Then
        If rngPrec.Cells.CountLarge > m_MAX_PRECEDENTS Then
            strWhy = rngAnswer.Address(False, False) & " depends on " & rngPrec.Cells.CountLarge & " cells - that is not a solve region."
            GoTo Cleanup
        End If
        For Each rngCell In rngPrec.Cells
            If rngCell.Column >= lngStartCol Then
                Set rngAnchor = Nothing
                If rngCell.HasFormula Then
                    Set rngAnchor = rngCell
                Else
                    ' Probe: a spilled child answers with its anchor; a plain value raises.
                    On Error Resume Next
                    Set rngAnchor = rngCell.SpillParent
                    On Error GoTo ErrHandler
                End If
                If Not rngAnchor Is Nothing Then
                    If Intersect(rngOut, rngAnchor) Is Nothing Then Set rngOut = Union(rngOut, rngAnchor)
                End If
            End If
        Next rngCell
    End If
    Set SolveRegion = rngOut

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    strWhy = "could not read the solve region: " & Err.Description
    LogError "SolveRegion", Err.Number, Err.Description
    Set SolveRegion = Nothing
    Resume Cleanup
End Function

' Purpose: rngBase plus every dependent of rngCell on its sheet (direct and indirect). Probe: Dependents
'          raises when there are none. Never raises.
Private Function UnionDependents(ByVal rngBase As Range, ByVal rngCell As Range) As Range

    Dim rngDeps As Range

    On Error Resume Next
    Set rngDeps = rngCell.Dependents
    On Error GoTo 0
    If rngDeps Is Nothing Then
        Set UnionDependents = rngBase
    Else
        Set UnionDependents = Union(rngBase, rngDeps)
    End If
End Function

' Purpose: The region's cells in dependency order (a step before every step that uses it), ties
'          broken left to right then top to bottom, the answer last. Addresses without $. A cycle
'          falls back to left-to-right order. Never raises.
Private Function DependencyOrder(ByVal rngRegion As Range, ByVal rngAnswer As Range) As Variant

    Dim rngCell     As Range
    Dim rngPrec     As Range
    Dim rngP        As Range
    Dim dicIndex    As Object
    Dim dicDeps     As Object       ' address -> Collection of addresses it depends on (inside the region)
    Dim dicDone     As Object
    Dim colDeps     As Collection
    Dim varAddr     As Variant
    Dim varD        As Variant
    Dim strBest     As String
    Dim strOut()    As String
    Dim lngN        As Long
    Dim lngCount    As Long
    Dim blnReady    As Boolean
    Dim blnProgress As Boolean
    Dim strAnswer   As String

    On Error GoTo ErrHandler

    strAnswer = rngAnswer.Address(False, False)
    Set dicIndex = CreateObject("Scripting.Dictionary")
    Set dicDeps = CreateObject("Scripting.Dictionary")
    Set dicDone = CreateObject("Scripting.Dictionary")
    For Each rngCell In rngRegion.Cells
        dicIndex(rngCell.Address(False, False)) = rngCell.Column * 10000000 + rngCell.Row   ' sort key: column, then row
    Next rngCell
    For Each rngCell In rngRegion.Cells
        Set colDeps = New Collection
        Set rngPrec = Nothing
        On Error Resume Next
        Set rngPrec = rngCell.DirectPrecedents
        On Error GoTo ErrHandler
        If Not rngPrec Is Nothing Then
            For Each rngP In rngPrec.Cells
                If dicIndex.Exists(rngP.Address(False, False)) And rngP.Address <> rngCell.Address Then
                    colDeps.Add rngP.Address(False, False)
                End If
            Next rngP
        End If
        Set dicDeps(rngCell.Address(False, False)) = colDeps
    Next rngCell

    lngCount = dicIndex.Count
    ReDim strOut(1 To lngCount)
    lngN = 0
    Do While lngN < lngCount
        ' The leftmost (then topmost) cell whose in-region precedents are all placed; the answer waits.
        strBest = vbNullString
        blnProgress = False
        For Each varAddr In dicIndex.Keys
            If Not dicDone.Exists(varAddr) Then
                If varAddr <> strAnswer Or lngN = lngCount - 1 Then
                    blnReady = True
                    For Each varD In dicDeps(varAddr)
                        If Not dicDone.Exists(varD) Then blnReady = False: Exit For
                    Next varD
                    If blnReady Then
                        If Len(strBest) = 0 Then
                            strBest = varAddr
                        ElseIf dicIndex(varAddr) < dicIndex(strBest) Then
                            strBest = varAddr
                        End If
                    End If
                End If
            End If
        Next varAddr
        If Len(strBest) = 0 Then
            ' A cycle (circular references): take what is left, left to right.
            For Each varAddr In dicIndex.Keys
                If Not dicDone.Exists(varAddr) Then
                    If Len(strBest) = 0 Then
                        strBest = varAddr
                    ElseIf dicIndex(varAddr) < dicIndex(strBest) Then
                        strBest = varAddr
                    End If
                End If
            Next varAddr
        End If
        lngN = lngN + 1
        strOut(lngN) = strBest
        dicDone(strBest) = True
    Loop
    DependencyOrder = strOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LogError "DependencyOrder", Err.Number, Err.Description
    ReDim strOut(1 To 1)
    strOut(1) = strAnswer
    DependencyOrder = strOut
End Function

' Purpose: Cut the step whose ORIGINAL address is strOrig to strTo, keeping both dictionaries current.
Private Sub MoveStep(ByVal ws As Worksheet, ByVal dicWhere As Object, ByVal dicAtCell As Object, _
                     ByVal strOrig As String, ByVal strTo As String)
    Dim strFrom As String
    strFrom = dicWhere(strOrig)
    ws.Range(strFrom).Cut ws.Range(strTo)
    dicWhere(strOrig) = strTo
    dicAtCell.Remove strFrom
    dicAtCell(strTo) = strOrig
End Sub

'==============================================================================
' Formula text
'==============================================================================

' Purpose: True for a nested formula: "={" ... "}" with the first brace matching the last.
Private Function IsNested(ByVal strF As String) As Boolean
    Dim lngI As Long
    Dim lngDepth As Long
    Dim blnQuote As Boolean
    Dim strC As String
    If Left$(strF, 2) <> "={" Or Right$(strF, 1) <> "}" Then Exit Function
    For lngI = 2 To Len(strF)
        strC = Mid$(strF, lngI, 1)
        If strC = """" Then
            blnQuote = Not blnQuote
        ElseIf Not blnQuote Then
            If strC = "{" Then lngDepth = lngDepth + 1
            If strC = "}" Then
                lngDepth = lngDepth - 1
                If lngDepth = 0 Then
                    IsNested = (lngI = Len(strF))
                    Exit Function
                End If
            End If
        End If
    Next lngI
End Function

' Purpose: True when the cell shows #SPILL!.
Private Function IsSpillError(ByVal rngCell As Range) As Boolean
    Dim varV As Variant
    On Error Resume Next
    varV = rngCell.Value2
    If IsError(varV) Then IsSpillError = (CLng(varV) = m_ERR_SPILL)
End Function

' Purpose: True for a plain A1-style cell address (optional $ signs), e.g. N305, $N$305.
Private Function IsCellAddress(ByVal strTok As String) As Boolean
    Dim strT As String
    Dim lngI As Long
    Dim lngLetters As Long
    strT = UCase$(Replace(strTok, "$", ""))
    For lngI = 1 To Len(strT)
        If Mid$(strT, lngI, 1) Like "[A-Z]" Then
            If lngLetters <> lngI - 1 Then Exit Function      ' letters first, then digits only
            lngLetters = lngI
        ElseIf Not Mid$(strT, lngI, 1) Like "[0-9]" Then
            Exit Function
        End If
    Next lngI
    IsCellAddress = (lngLetters >= 1 And lngLetters <= 3 And Len(strT) > lngLetters And Len(strT) - lngLetters <= 7)
End Function

' Purpose: The sheet a reference token at position lngPos is qualified with, or "" when unqualified:
'          the text before a "!" just before the token, quotes removed. Used to keep sheet-qualified
'          references to OTHER sheets alone.
Private Function QualifierBefore(ByVal strF As String, ByVal lngPos As Long) As String
    Dim lngJ As Long
    Dim lngQ As Long
    If lngPos < 2 Then Exit Function
    If Mid$(strF, lngPos - 1, 1) <> "!" Then Exit Function
    lngJ = lngPos - 2
    If lngJ < 1 Then Exit Function
    If Mid$(strF, lngJ, 1) = "'" Then
        lngQ = lngJ - 1
        Do While lngQ >= 1
            If Mid$(strF, lngQ, 1) = "'" Then
                If lngQ > 1 Then
                    If Mid$(strF, lngQ - 1, 1) = "'" Then lngQ = lngQ - 2: GoTo NextQ    ' '' inside a name
                End If
                Exit Do
            End If
NextQ:
            lngQ = lngQ - 1
        Loop
        If lngQ < 1 Then lngQ = 1
        QualifierBefore = Replace(Mid$(strF, lngQ + 1, lngJ - lngQ - 1), "''", "'")
    Else
        Do While lngJ >= 1
            If Not Mid$(strF, lngJ, 1) Like "[A-Za-z0-9_.]" Then Exit Do
            lngJ = lngJ - 1
        Loop
        QualifierBefore = Mid$(strF, lngJ + 1, lngPos - 2 - lngJ)
    End If
End Function

' Purpose: strF with the trailing # removed from every reference to an anchor in dicAnchors (keys:
'          addresses without $, this sheet). Text inside quotes is left alone, so is a reference
'          qualified with another sheet's name. lngCount returns how many were removed.
Private Function StripHash(ByVal strF As String, ByVal dicAnchors As Object, ByVal strSheet As String, _
                           ByRef lngCount As Long) As String

    Dim lngI    As Long
    Dim lngJ    As Long
    Dim strTok  As String
    Dim strQual As String
    Dim blnQuote As Boolean

    lngCount = 0
    lngI = 1
    Do While lngI <= Len(strF)
        If Mid$(strF, lngI, 1) = """" Then
            blnQuote = Not blnQuote
        ElseIf Not blnQuote And Mid$(strF, lngI, 1) = "#" Then
            lngJ = lngI - 1
            Do While lngJ >= 1
                If Not Mid$(strF, lngJ, 1) Like "[A-Za-z0-9$]" Then Exit Do
                lngJ = lngJ - 1
            Loop
            strTok = Mid$(strF, lngJ + 1, lngI - lngJ - 1)
            If IsCellAddress(strTok) Then
                strQual = QualifierBefore(strF, lngJ + 1)
                If Len(strQual) = 0 Or StrComp(strQual, strSheet, vbTextCompare) = 0 Then
                    If dicAnchors.Exists(UCase$(Replace(strTok, "$", ""))) Then
                        strF = Left$(strF, lngI - 1) & Mid$(strF, lngI + 1)
                        lngCount = lngCount + 1
                        lngI = lngI - 1
                    End If
                End If
            End If
        End If
        lngI = lngI + 1
    Loop
    StripHash = strF
End Function

' Purpose: strF with a # appended to every bare reference to an anchor in dicAnchors (this sheet,
'          outside quotes, not a function name, not part of a range, not already #). lngCount returns
'          how many were added.
Private Function AddHash(ByVal strF As String, ByVal dicAnchors As Object, ByVal strSheet As String, _
                         ByRef lngCount As Long) As String

    Dim lngI    As Long
    Dim lngJ    As Long
    Dim strTok  As String
    Dim strQual As String
    Dim strNext As String
    Dim strPrev As String
    Dim blnQuote As Boolean

    lngCount = 0
    lngI = 1
    Do While lngI <= Len(strF)
        If Mid$(strF, lngI, 1) = """" Then
            blnQuote = Not blnQuote
            lngI = lngI + 1
        ElseIf Not blnQuote And Mid$(strF, lngI, 1) Like "[A-Za-z$]" Then
            lngJ = lngI
            Do While lngJ <= Len(strF)
                If Not Mid$(strF, lngJ, 1) Like "[A-Za-z0-9$_.]" Then Exit Do
                lngJ = lngJ + 1
            Loop
            strTok = Mid$(strF, lngI, lngJ - lngI)
            strNext = Mid$(strF, lngJ, 1)
            If lngI > 1 Then strPrev = Mid$(strF, lngI - 1, 1) Else strPrev = vbNullString
            If IsCellAddress(strTok) And strNext <> "(" And strNext <> "#" And strNext <> ":" And strPrev <> ":" Then
                strQual = QualifierBefore(strF, lngI)
                If Len(strQual) = 0 Or StrComp(strQual, strSheet, vbTextCompare) = 0 Then
                    If dicAnchors.Exists(UCase$(Replace(strTok, "$", ""))) Then
                        strF = Left$(strF, lngJ - 1) & "#" & Mid$(strF, lngJ)
                        lngCount = lngCount + 1
                        lngJ = lngJ + 1
                    End If
                End If
            End If
            lngI = lngJ
        Else
            lngI = lngI + 1
        End If
    Loop
    AddHash = strF
End Function

' Purpose: strF with every bare reference to a key of dicMap (an original step address, this sheet)
'          replaced by its value (the copy's address). References to anything else are unchanged.
'          lngCount returns how many were rewritten.
Private Function RewriteRefs(ByVal strF As String, ByVal dicMap As Object, ByVal strSheet As String, _
                             ByRef lngCount As Long) As String

    Dim lngI    As Long
    Dim lngJ    As Long
    Dim strTok  As String
    Dim strKey  As String
    Dim strQual As String
    Dim strNext As String
    Dim strNew  As String
    Dim blnQuote As Boolean

    lngCount = 0
    lngI = 1
    Do While lngI <= Len(strF)
        If Mid$(strF, lngI, 1) = """" Then
            blnQuote = Not blnQuote
            lngI = lngI + 1
        ElseIf Not blnQuote And Mid$(strF, lngI, 1) Like "[A-Za-z$]" Then
            lngJ = lngI
            Do While lngJ <= Len(strF)
                If Not Mid$(strF, lngJ, 1) Like "[A-Za-z0-9$_.]" Then Exit Do
                lngJ = lngJ + 1
            Loop
            strTok = Mid$(strF, lngI, lngJ - lngI)
            strNext = Mid$(strF, lngJ, 1)
            If IsCellAddress(strTok) And strNext <> "(" Then
                strQual = QualifierBefore(strF, lngI)
                If Len(strQual) = 0 Or StrComp(strQual, strSheet, vbTextCompare) = 0 Then
                    strKey = UCase$(Replace(strTok, "$", ""))
                    If dicMap.Exists(strKey) Then
                        strNew = dicMap(strKey)
                        strF = Left$(strF, lngI - 1) & strNew & Mid$(strF, lngJ)
                        lngJ = lngI + Len(strNew)
                        lngCount = lngCount + 1
                    End If
                End If
            End If
            lngI = lngJ
        Else
            lngI = lngI + 1
        End If
    Loop
    RewriteRefs = strF
End Function
