Attribute VB_Name = "modFlatten"
Option Explicit

'==============================================================================
' modFlatten - "Paste Flattened List With Formatting"
'
' Jaq's spec, 2026-09-22:
'   copy a range -> select a target cell -> run the command -> the flattened
'   list is written from that cell, using VBA, reading the copied range.
'
' Output columns, header row first:
'   Items | Row # | Column # | Address | RCRef | Fill | Font | BdrT | BdrR | BdrB | BdrL
'
' Design rules applied (see commandreview\PROGRESS.md):
'   - Colours as #RRGGBB hex, never the raw Interior.Color Long (which is BGR
'     and unreadable). Hex is XLOOKUP-able against a legend.
'   - BLANK CELLS ARE INCLUDED. A blank cell with a fill is the whole point on a
'     colour-coded map. This is the OPPOSITE of GridToCol's own default.
'   - One array write, not cell-by-cell. Reading formatting must be per-cell
'     (Interior.Color has no array form) but the write is a single assignment.
'   - The write is GUARDED like a paste: the target cell may be overwritten, the
'     rest of the block must be empty and unmerged. Otherwise the output goes to a
'     new sheet and the target cell links to it (Jaq, 2026-10-05).
'   - RCRef is packed 10^6 * row + column, matching GridToCol's convention.
'   - A shared border is ONE line with TWO possible owners: A1's right edge and
'     B1's left edge are the same wall. Every edge test ORs both sides, or walls
'     drawn "from the other side" vanish.
'==============================================================================

' --- CONSTANTS (module) ---
Private Const m_OUT_COLS   As Long = 11
Private Const m_CLIP_LINK  As String = "Link"

#If VBA7 Then
    Private Declare PtrSafe Function OpenClipboard Lib "user32" (ByVal hwnd As LongPtr) As Long
    Private Declare PtrSafe Function CloseClipboard Lib "user32" () As Long
    Private Declare PtrSafe Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As LongPtr
    Private Declare PtrSafe Function RegisterClipboardFormatA Lib "user32" (ByVal lpString As String) As Long
    Private Declare PtrSafe Function GlobalLock Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Function GlobalUnlock Lib "kernel32" (ByVal hMem As LongPtr) As Long
    Private Declare PtrSafe Function GlobalSize Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As LongPtr, ByVal Length As LongPtr)
#Else
    Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function CloseClipboard Lib "user32" () As Long
    Private Declare Function GetClipboardData Lib "user32" (ByVal wFormat As Long) As Long
    Private Declare Function RegisterClipboardFormatA Lib "user32" (ByVal lpString As String) As Long
    Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Long, ByVal Length As Long)
#End If


'------------------------------------------------------------------------------
' THE COMMAND. Copy a range, select a target cell, run this.
'------------------------------------------------------------------------------
Public Sub PasteFlattenedListWithFormatting()

    Dim rngSource   As Range
    Dim rngTarget   As Range

    On Error GoTo ErrHandler

    Set rngTarget = ActiveCell
    Set rngSource = CopiedRange()

    If rngSource Is Nothing Then
        rngTarget.value = "FLATTEN FAILED - no copied range found. Copy a range first, then run this."
        GoTo Cleanup
    End If

    PasteFlattenedAt rngSource, rngTarget

Cleanup:
    Exit Sub

ErrHandler:
    On Error Resume Next
    If rngTarget Is Nothing Then Set rngTarget = ActiveCell
    rngTarget.value = "FLATTEN ERROR " & Err.Number & ": " & Err.Description
    Resume Cleanup

End Sub


'------------------------------------------------------------------------------
' Write the flattened list of rngSource from rngTarget, like a paste: the target
' cell itself may be overwritten, the rest of the block must be empty (Jaq,
' 2026-10-05). If it does not fit, the list goes to a new sheet next to this
' one and the target cell becomes a link to it; the user stays where they were.
'------------------------------------------------------------------------------
Private Sub PasteFlattenedAt(ByVal rngSource As Range, ByVal rngTarget As Range)

    Dim varOut  As Variant
    Dim lngRows As Long
    Dim wsHere  As Worksheet
    Dim wsOut   As Worksheet

    varOut = BuildFlattened(rngSource)
    lngRows = UBound(varOut, 1)
    Set wsHere = rngTarget.Worksheet

    If TargetIsClear(rngTarget, lngRows, m_OUT_COLS) Then
        rngTarget.Resize(lngRows, m_OUT_COLS).value = varOut
    Else
        Set wsOut = wsHere.Parent.Worksheets.Add(After:=wsHere)
        wsOut.Range("A1").Resize(lngRows, m_OUT_COLS).value = varOut
        rngTarget.ClearContents
        wsHere.Hyperlinks.Add Anchor:=rngTarget, Address:="", _
            SubAddress:="'" & Replace(wsOut.Name, "'", "''") & "'!A1", _
            TextToDisplay:="Unable to fit here: see " & wsOut.Name
        wsHere.Activate
        rngTarget.Select
    End If

End Sub


'------------------------------------------------------------------------------
' Test hook: the command without the clipboard. Addresses on the active sheet.
'------------------------------------------------------------------------------
Public Sub TestPasteFlattened(ByVal strSource As String, ByVal strTarget As String)
    PasteFlattenedAt ActiveSheet.Range(strSource), ActiveSheet.Range(strTarget)
End Sub


'------------------------------------------------------------------------------
' Build the output array. Header row plus one row per cell, in row-major order.
'------------------------------------------------------------------------------
Public Function BuildFlattened(ByVal rngSource As Range) As Variant

    Dim arrOut()  As Variant
    Dim rngCell   As Range
    Dim lngN      As Long
    Dim lngI      As Long

    lngN = rngSource.Cells.Count
    ReDim arrOut(1 To lngN + 1, 1 To m_OUT_COLS)

    arrOut(1, 1) = "Items"
    arrOut(1, 2) = "Row #"
    arrOut(1, 3) = "Column #"
    arrOut(1, 4) = "Address"
    arrOut(1, 5) = "RCRef"
    arrOut(1, 6) = "Fill"
    arrOut(1, 7) = "Font"
    arrOut(1, 8) = "BdrT"
    arrOut(1, 9) = "BdrR"
    arrOut(1, 10) = "BdrB"
    arrOut(1, 11) = "BdrL"

    lngI = 1
    For Each rngCell In rngSource.Cells
        lngI = lngI + 1
        arrOut(lngI, 1) = rngCell.Value2
        arrOut(lngI, 2) = rngCell.Row
        arrOut(lngI, 3) = rngCell.Column
        arrOut(lngI, 4) = rngCell.Address(False, False)
        arrOut(lngI, 5) = 1000000# * CDbl(rngCell.Row) + CDbl(rngCell.Column)
        arrOut(lngI, 6) = FillHex(rngCell)
        arrOut(lngI, 7) = FontHex(rngCell)
        arrOut(lngI, 8) = BorderFlag(rngCell, xlEdgeTop)
        arrOut(lngI, 9) = BorderFlag(rngCell, xlEdgeRight)
        arrOut(lngI, 10) = BorderFlag(rngCell, xlEdgeBottom)
        arrOut(lngI, 11) = BorderFlag(rngCell, xlEdgeLeft)
    Next rngCell

    BuildFlattened = arrOut

End Function


'------------------------------------------------------------------------------
' Fill colour as #RRGGBB, or "" when there is no fill.
'------------------------------------------------------------------------------
Public Function FillHex(ByVal rngCell As Range) As String

    On Error GoTo NoFill

    If rngCell.Interior.ColorIndex = xlColorIndexNone Then
        FillHex = vbNullString
    Else
        FillHex = LongToHex(rngCell.Interior.Color)
    End If
    Exit Function

NoFill:
    FillHex = vbNullString

End Function


'------------------------------------------------------------------------------
' Font colour as #RRGGBB. Automatic reports as "".
'------------------------------------------------------------------------------
Public Function FontHex(ByVal rngCell As Range) As String

    On Error GoTo NoFont

    If rngCell.Font.ColorIndex = xlColorIndexAutomatic Then
        FontHex = vbNullString
    Else
        FontHex = LongToHex(rngCell.Font.Color)
    End If
    Exit Function

NoFont:
    FontHex = vbNullString

End Function


'------------------------------------------------------------------------------
' Excel stores colour as BGR in a Long. Unpack to #RRGGBB.
'------------------------------------------------------------------------------
Public Function LongToHex(ByVal lngColor As Long) As String

    Dim lngR As Long
    Dim lngG As Long
    Dim lngB As Long

    lngR = lngColor Mod 256
    lngG = (lngColor \ 256) Mod 256
    lngB = (lngColor \ 65536) Mod 256

    LongToHex = "#" & Right$("0" & Hex$(lngR), 2) _
                    & Right$("0" & Hex$(lngG), 2) _
                    & Right$("0" & Hex$(lngB), 2)

End Function


'------------------------------------------------------------------------------
' 1 if this edge carries a border, counting a border owned by the NEIGHBOUR.
' A1's right edge and B1's left edge are the same wall.
'------------------------------------------------------------------------------
Public Function BorderFlag(ByVal rngCell As Range, ByVal lngEdge As Long) As Long

    Dim rngNb As Range

    If HasBorder(rngCell, lngEdge) Then
        BorderFlag = 1
        Exit Function
    End If

    Set rngNb = NeighbourCell(rngCell, lngEdge)
    If Not rngNb Is Nothing Then
        If HasBorder(rngNb, OppositeEdge(lngEdge)) Then
            BorderFlag = 1
            Exit Function
        End If
    End If

    BorderFlag = 0

End Function


Private Function HasBorder(ByVal rngCell As Range, ByVal lngEdge As Long) As Boolean
    On Error Resume Next
    HasBorder = (rngCell.Borders(lngEdge).LineStyle <> xlLineStyleNone)
    On Error GoTo 0
End Function


Private Function NeighbourCell(ByVal rngCell As Range, ByVal lngEdge As Long) As Range

    Dim lngR As Long
    Dim lngC As Long

    lngR = rngCell.Row
    lngC = rngCell.Column

    Select Case lngEdge
        Case xlEdgeTop:    lngR = lngR - 1
        Case xlEdgeBottom: lngR = lngR + 1
        Case xlEdgeLeft:   lngC = lngC - 1
        Case xlEdgeRight:  lngC = lngC + 1
    End Select

    If lngR < 1 Or lngC < 1 Then Exit Function
    If lngR > rngCell.Worksheet.Rows.Count Then Exit Function
    If lngC > rngCell.Worksheet.Columns.Count Then Exit Function

    Set NeighbourCell = rngCell.Worksheet.Cells(lngR, lngC)

End Function


Private Function OppositeEdge(ByVal lngEdge As Long) As Long
    Select Case lngEdge
        Case xlEdgeTop:    OppositeEdge = xlEdgeBottom
        Case xlEdgeBottom: OppositeEdge = xlEdgeTop
        Case xlEdgeLeft:   OppositeEdge = xlEdgeRight
        Case xlEdgeRight:  OppositeEdge = xlEdgeLeft
    End Select
End Function


'------------------------------------------------------------------------------
' True when the output block is empty APART FROM the target cell itself (a paste
' overwrites the cell you paste into) and holds no merged cells, so writing
' cannot destroy anything the user did not point at.
'------------------------------------------------------------------------------
Private Function TargetIsClear(ByVal rngTarget As Range, _
                               ByVal lngRows As Long, _
                               ByVal lngCols As Long) As Boolean

    Dim rngBlock As Range
    Dim lngUsed  As Long
    Dim varMrg   As Variant

    On Error GoTo NotClear

    If rngTarget.Row + lngRows - 1 > rngTarget.Worksheet.Rows.Count Then GoTo NotClear
    If rngTarget.Column + lngCols - 1 > rngTarget.Worksheet.Columns.Count Then GoTo NotClear

    Set rngBlock = rngTarget.Resize(lngRows, lngCols)
    varMrg = rngBlock.MergeCells
    If IsNull(varMrg) Then GoTo NotClear
    If varMrg Then GoTo NotClear

    lngUsed = Application.WorksheetFunction.CountA(rngBlock)
    If Not IsEmpty(rngTarget.Cells(1, 1).Formula) Then
        If Len(rngTarget.Cells(1, 1).Formula) > 0 Then lngUsed = lngUsed - 1
    End If
    TargetIsClear = (lngUsed = 0)
    Exit Function

NotClear:
    TargetIsClear = False

End Function


'==============================================================================
' Resolving the COPIED range.
'
' VBA has no API for "the range currently on the copy marquee". Excel does
' however publish a "Link" clipboard format holding the source reference, as
' NUL-separated parts:  Excel <NUL> [Book1.xlsx]Sheet1 <NUL> R1C1:R5C5 <NUL>
' That is what we read here. Clipboard TEXT would lose the formatting that is
' the entire point of this command, so it has to resolve to a real Range.
'==============================================================================
Public Function CopiedRange() As Range

    Dim strLink      As String
    Dim arrParts()   As String
    Dim strBookSheet As String
    Dim strR1C1      As String
    Dim strBook      As String
    Dim strSheet     As String
    Dim strA1        As String
    Dim lngOpen      As Long
    Dim lngClose     As Long
    Dim wbSrc        As Workbook
    Dim wsSrc        As Worksheet

    On Error GoTo ErrHandler

    If Application.CutCopyMode = 0 Then Exit Function

    strLink = ClipboardLink()
    If Len(strLink) = 0 Then Exit Function

    arrParts = Split(strLink, vbNullChar)
    If UBound(arrParts) < 2 Then Exit Function

    strBookSheet = arrParts(1)
    strR1C1 = arrParts(2)
    If Len(strR1C1) = 0 Then Exit Function

    lngOpen = InStr(1, strBookSheet, "[")
    lngClose = InStr(1, strBookSheet, "]")
    If lngOpen > 0 And lngClose > lngOpen Then
        strBook = Mid$(strBookSheet, lngOpen + 1, lngClose - lngOpen - 1)
        strSheet = Mid$(strBookSheet, lngClose + 1)
    Else
        strSheet = strBookSheet
    End If

    ' A sheet name carrying spaces arrives wrapped in single quotes.
    If Left$(strSheet, 1) = "'" Then strSheet = Mid$(strSheet, 2)
    If Right$(strSheet, 1) = "'" Then strSheet = Left$(strSheet, Len(strSheet) - 1)

    If Len(strBook) > 0 Then
        Set wbSrc = Workbooks(strBook)
    Else
        Set wbSrc = ActiveWorkbook
    End If
    Set wsSrc = wbSrc.Worksheets(strSheet)

    strA1 = Application.ConvertFormula(strR1C1, xlR1C1, xlA1)
    Set CopiedRange = wsSrc.Range(strA1)
    Exit Function

ErrHandler:
    Set CopiedRange = Nothing

End Function


Private Function ClipboardLink() As String

    Dim lngFmt   As Long
    Dim lngSize  As Long
    Dim arrBytes() As Byte
    Dim lngI     As Long
    Dim strOut   As String
#If VBA7 Then
    Dim hMem As LongPtr
    Dim pMem As LongPtr
#Else
    Dim hMem As Long
    Dim pMem As Long
#End If

    On Error GoTo ErrHandler

    lngFmt = RegisterClipboardFormatA(m_CLIP_LINK)
    If lngFmt = 0 Then Exit Function

    If OpenClipboard(0) = 0 Then Exit Function

    hMem = GetClipboardData(lngFmt)
    If hMem <> 0 Then
        lngSize = CLng(GlobalSize(hMem))
        If lngSize > 0 Then
            pMem = GlobalLock(hMem)
            If pMem <> 0 Then
                ReDim arrBytes(0 To lngSize - 1)
                CopyMemory arrBytes(0), pMem, lngSize
                GlobalUnlock hMem
                For lngI = 0 To lngSize - 1
                    strOut = strOut & Chr$(arrBytes(lngI))
                Next lngI
            End If
        End If
    End If

    CloseClipboard
    ClipboardLink = strOut
    Exit Function

ErrHandler:
    On Error Resume Next
    CloseClipboard
    ClipboardLink = vbNullString

End Function

'==============================================================================
' The test harness for this module lives in a separate development workbook
'   FlattenTools.xlsm  (modTests)
' and is deliberately NOT shipped here - its fixture builder clears a worksheet.
' Run modTests.RunAllTests there after changing anything in this module.
'==============================================================================


