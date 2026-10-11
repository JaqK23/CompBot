Attribute VB_Name = "modWalk"
Option Explicit

'==============================================================================
' modWalk - "Record Walk Route"
'
' Jaq's spec, 2026-09-22: start on a cell; every ORTHOGONAL move - arrow,
' Ctrl+arrow, or an orthogonal CLICK - appends the whole run from the anchor to
' the new cell and re-anchors there. A DIAGONAL click ends the mode and emits
' the route. The point is to capture a path that exists only visually.
'
' HER RULINGS, 2026-09-22 - these are decisions, not open questions:
'   1. An orthogonal CLICK is a move, not an exit. Clicking is often easier than
'      Ctrl+arrow when there is no content to Ctrl+arrow over. Arrow, Ctrl+arrow
'      and click are all just SelectionChange, so no gesture detection is needed
'      - only an orthogonality test against the anchor.
'   2. The exit gesture is a DIAGONAL click. The exit cell does NOT join the route.
'   3. A self-crossing route keeps EVERY visit, in order, with repeats flagged.
'      Not deduped - order matters more than uniqueness for a route.
'   4. NO vbaInit / vbaFin on this command. The mode depends on events firing and
'      that wrapper turns them off. THIS IS A DELIBERATE, RULED EXCEPTION TO AN
'      OTHERWISE ABSOLUTE HOUSE RULE - do not "fix" it back, or the command goes
'      deaf and silently records nothing. The previous EnableEvents state is
'      captured and restored, including on the error path.
'   5. DO NOT COUNT THE CORNER TWICE. Walk right to X then up from X: X ends one
'      run and starts the next. AppendRun starts at offset 1 for exactly this.
'
' Open, chosen default: a click-and-DRAG raises SelectionChange with a multi-cell
' selection. We take the LAST cell of the selection as the move target.
'==============================================================================

' --- CONSTANTS (module) ---
Private Const m_OUT_COLS As Long = 13

' --- module state ---
Private m_blnActive         As Boolean
Private m_blnPrevEvents     As Boolean
' A VBA Boolean defaults to False, so restoring m_blnPrevEvents without having
' captured it would turn events OFF and leave the session deaf. This flag means
' "we actually took a reading"; without it, ResetState must not touch events.
Private m_blnEventsCaptured As Boolean
Private m_colRoute      As Collection
Private m_rngAnchor     As Range
Private m_objWatcher    As clsWalkWatcher
' True for Record Walk Route Here: the route goes to the right of the current sheet's
' contents instead of onto a new sheet (Jaq, 2026-09-28).
Private m_blnHere       As Boolean


'------------------------------------------------------------------------------
' THE COMMAND. Select the starting cell, run this, then walk.
'------------------------------------------------------------------------------
Public Sub StartWalkRoute()

    On Error GoTo ErrHandler

    If m_blnActive Then Exit Sub

    Set m_colRoute = New Collection
    Set m_rngAnchor = ActiveCell
    m_colRoute.Add m_rngAnchor

    m_blnPrevEvents = Application.EnableEvents
    m_blnEventsCaptured = True
    Application.EnableEvents = True

    Set m_objWatcher = New clsWalkWatcher
    m_blnActive = True

    Application.StatusBar = "Walk route recording - " & m_rngAnchor.Address(False, False) & _
                            ". Move orthogonally to extend. Diagonal click to finish."
    Exit Sub

ErrHandler:
    ResetState

End Sub


'------------------------------------------------------------------------------
' THE IN-SHEET VARIANT (Record Walk Route Here, RWH). Same walk; the route table
' is written to the right of this sheet's contents rather than to a new sheet.
'------------------------------------------------------------------------------
Public Sub StartWalkRouteHere()
    If m_blnActive Then Exit Sub
    StartWalkRoute
    If m_blnActive Then m_blnHere = True
End Sub


'------------------------------------------------------------------------------
' Manual abort, for when a diagonal click is awkward to reach.
'------------------------------------------------------------------------------
Public Sub StopWalkRoute()
    If m_blnActive Then FinishWalk
End Sub


'------------------------------------------------------------------------------
' Called by clsWalkWatcher on every selection change. Public so tests can drive
' it directly without a real user moving around the sheet.
'------------------------------------------------------------------------------
Public Sub HandleSelection(ByVal Target As Range)

    Dim rngNew As Range

    If Not m_blnActive Then Exit Sub
    If m_rngAnchor Is Nothing Then Exit Sub

    On Error GoTo ErrHandler

    ' Leaving the sheet ends the route.
    If Target.Worksheet.Name <> m_rngAnchor.Worksheet.Name Then
        FinishWalk
        Exit Sub
    End If

    ' A drag selects several cells; take the last one as the target.
    Set rngNew = Target.Cells(Target.Cells.Count)

    ' Re-selecting the anchor is a no-op, not an exit.
    If rngNew.Row = m_rngAnchor.Row And rngNew.Column = m_rngAnchor.Column Then Exit Sub

    If rngNew.Row = m_rngAnchor.Row Or rngNew.Column = m_rngAnchor.Column Then
        AppendRun m_rngAnchor, rngNew
        Set m_rngAnchor = rngNew
        Application.StatusBar = "Walk route recording - " & m_colRoute.Count & _
                                " cells. Diagonal click to finish."
    Else
        FinishWalk
    End If
    Exit Sub

ErrHandler:
    FinishWalk

End Sub


'------------------------------------------------------------------------------
' Append every cell from rngFrom to rngTo EXCLUDING rngFrom itself.
' Excluding the start is what stops the corner being counted twice.
'------------------------------------------------------------------------------
Private Sub AppendRun(ByVal rngFrom As Range, ByVal rngTo As Range)

    Dim wsHere   As Worksheet
    Dim lngStep  As Long
    Dim lngCount As Long
    Dim lngI     As Long

    Set wsHere = rngFrom.Worksheet

    If rngFrom.Row = rngTo.Row Then
        lngCount = Abs(rngTo.Column - rngFrom.Column)
        If rngTo.Column > rngFrom.Column Then lngStep = 1 Else lngStep = -1
        For lngI = 1 To lngCount
            m_colRoute.Add wsHere.Cells(rngFrom.Row, rngFrom.Column + lngI * lngStep)
        Next lngI
    Else
        lngCount = Abs(rngTo.Row - rngFrom.Row)
        If rngTo.Row > rngFrom.Row Then lngStep = 1 Else lngStep = -1
        For lngI = 1 To lngCount
            m_colRoute.Add wsHere.Cells(rngFrom.Row + lngI * lngStep, rngFrom.Column)
        Next lngI
    End If

End Sub


'------------------------------------------------------------------------------
' End the mode and write the route out to a new sheet.
'------------------------------------------------------------------------------
Private Sub FinishWalk()

    Dim varOut As Variant
    Dim wsOut  As Worksheet
    Dim wbHost As Workbook
    Dim rngTop As Range

    On Error GoTo ErrHandler

    If m_colRoute Is Nothing Then GoTo Cleanup
    If m_colRoute.Count = 0 Then GoTo Cleanup

    ' Stop listening BEFORE writing or selecting anything. The Select below raises SelectionChange; with
    ' the mode still active, HandleSelection read it as a diagonal click and called FinishWalk again,
    ' which wrote another table further right and selected it, until Excel ran out of stack and crashed
    ' (Jaq, 2026-10-05, Record Walk Route Here on the Case sheet).
    m_blnActive = False
    Set m_objWatcher = Nothing

    varOut = BuildRouteArray()

    If m_blnHere Then
        ' To the right of the walk's START cell, level with it: the first position from two columns past
        ' the start where a block the size of the table is completely clear (no values, no fill, no
        ' merges) - Jaq, 2026-10-05. Only if there is none: one column after the sheet's contents.
        Set wsOut = m_rngAnchor.Worksheet
        Set rngTop = ClearSpotRightOf(m_colRoute(1), UBound(varOut, 1), m_OUT_COLS)
        If rngTop Is Nothing Then
            With wsOut.UsedRange
                Set rngTop = wsOut.Cells(.Row, .Column + .Columns.Count + 1)
            End With
        End If
        rngTop.Resize(UBound(varOut, 1), m_OUT_COLS).value = varOut
        rngTop.Select
        modCaseSetup.ShowRange rngTop
    Else
        Set wbHost = m_rngAnchor.Worksheet.Parent
        Set wsOut = wbHost.Worksheets.Add
        wsOut.Range("A1").Resize(UBound(varOut, 1), m_OUT_COLS).value = varOut
        wsOut.Activate
    End If

Cleanup:
    ResetState
    Exit Sub

ErrHandler:
    Resume Cleanup

End Sub


'------------------------------------------------------------------------------
' Build the route table. Same shape as the flattened list, plus Step and Repeat.
' Public so a test can inspect it without a sheet being created.
'------------------------------------------------------------------------------
Public Function BuildRouteArray() As Variant

    Dim arrOut()  As Variant
    Dim dicSeen   As Object
    Dim rngCell   As Range
    Dim lngI      As Long
    Dim strAddr   As String

    Set dicSeen = CreateObject("Scripting.Dictionary")

    ReDim arrOut(1 To m_colRoute.Count + 1, 1 To m_OUT_COLS)

    arrOut(1, 1) = "Step"
    arrOut(1, 2) = "Items"
    arrOut(1, 3) = "Row #"
    arrOut(1, 4) = "Column #"
    arrOut(1, 5) = "Address"
    arrOut(1, 6) = "RCRef"
    arrOut(1, 7) = "Fill"
    arrOut(1, 8) = "Font"
    arrOut(1, 9) = "BdrT"
    arrOut(1, 10) = "BdrR"
    arrOut(1, 11) = "BdrB"
    arrOut(1, 12) = "BdrL"
    arrOut(1, 13) = "Repeat"

    lngI = 1
    For Each rngCell In m_colRoute
        lngI = lngI + 1
        strAddr = rngCell.Address(False, False)

        arrOut(lngI, 1) = lngI - 1
        arrOut(lngI, 2) = rngCell.Value2
        arrOut(lngI, 3) = rngCell.Row
        arrOut(lngI, 4) = rngCell.Column
        arrOut(lngI, 5) = strAddr
        arrOut(lngI, 6) = 1000000# * CDbl(rngCell.Row) + CDbl(rngCell.Column)
        arrOut(lngI, 7) = modFlatten.FillHex(rngCell)
        arrOut(lngI, 8) = modFlatten.FontHex(rngCell)
        arrOut(lngI, 9) = modFlatten.BorderFlag(rngCell, xlEdgeTop)
        arrOut(lngI, 10) = modFlatten.BorderFlag(rngCell, xlEdgeRight)
        arrOut(lngI, 11) = modFlatten.BorderFlag(rngCell, xlEdgeBottom)
        arrOut(lngI, 12) = modFlatten.BorderFlag(rngCell, xlEdgeLeft)

        If dicSeen.Exists(strAddr) Then
            arrOut(lngI, 13) = 1
        Else
            arrOut(lngI, 13) = 0
            dicSeen.Add strAddr, 1
        End If
    Next rngCell

    BuildRouteArray = arrOut

End Function


'------------------------------------------------------------------------------
' The first lngRows x lngCols block, level with rngStart and starting two columns
' to its right, that holds no values, no fill and no merged cells. Nothing if the
' row runs out first.
'------------------------------------------------------------------------------
Private Function ClearSpotRightOf(ByVal rngStart As Range, ByVal lngRows As Long, _
                                  ByVal lngCols As Long) As Range

    Dim wsHere  As Worksheet
    Dim lngCol  As Long
    Dim rngTry  As Range
    Dim varMrg  As Variant

    On Error GoTo ErrHandler
    Set wsHere = rngStart.Worksheet
    If rngStart.Row + lngRows - 1 > wsHere.Rows.Count Then Exit Function
    For lngCol = rngStart.Column + 2 To wsHere.Columns.Count - lngCols + 1
        Set rngTry = wsHere.Cells(rngStart.Row, lngCol).Resize(lngRows, lngCols)
        If Application.WorksheetFunction.CountA(rngTry) = 0 Then
            varMrg = rngTry.MergeCells
            If Not IsNull(varMrg) Then
                If varMrg = False And rngTry.Interior.ColorIndex = xlColorIndexNone Then
                    Set ClearSpotRightOf = rngTry.Cells(1, 1)
                    Exit Function
                End If
            End If
        End If
    Next lngCol
    Exit Function

ErrHandler:
    Set ClearSpotRightOf = Nothing

End Function


'------------------------------------------------------------------------------
' Drop the event hook and put EnableEvents back exactly as we found it.
'------------------------------------------------------------------------------
Private Sub ResetState()

    On Error Resume Next
    Set m_objWatcher = Nothing
    Set m_rngAnchor = Nothing
    Set m_colRoute = Nothing
    m_blnActive = False
    m_blnHere = False
    ' Only ever put back a reading we actually took. Restoring an uncaptured
    ' default would silently disable events for the whole session.
    If m_blnEventsCaptured Then
        Application.EnableEvents = m_blnPrevEvents
        m_blnEventsCaptured = False
    End If
    Application.StatusBar = False
    On Error GoTo 0

End Sub


'------------------------------------------------------------------------------
' Test hooks - let a test drive the state machine without a live user.
'------------------------------------------------------------------------------
Public Sub TestBeginAt(ByVal rngStart As Range)
    ResetState
    Set m_colRoute = New Collection
    Set m_rngAnchor = rngStart
    m_colRoute.Add rngStart
    ' Mirror StartWalkRoute exactly, including the events capture, or the test
    ' exercises a different state machine from the real one.
    m_blnPrevEvents = Application.EnableEvents
    m_blnEventsCaptured = True
    Application.EnableEvents = True
    m_blnActive = True
End Sub

Public Function TestRouteCount() As Long
    If m_colRoute Is Nothing Then Exit Function
    TestRouteCount = m_colRoute.Count
End Function

Public Function TestIsActive() As Boolean
    TestIsActive = m_blnActive
End Function

Public Sub TestReset()
    ResetState
End Sub







