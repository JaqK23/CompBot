Attribute VB_Name = "modStackSheets"
Option Explicit

'==============================================================================
' modStackSheets - "Stack Sheets From Clipboard"
'
' Jaq's spec, 2026-09-22: drive it from a list of sheet names sitting in the
' CLIPBOARD, which may be EITHER
'    two cells  - the FIRST and LAST sheet name, taking everything between them
'                 in tab order (so 64 Area sheets need two cells, not 64)
'    N cells    - the exact sheets to stack, in the order given
' and the clipboard list may be a ROW or a COLUMN. Both are handled.
'
' It then works out the BOUNDING USED RANGE across those sheets - the smallest
' block that covers every sheet's data - and writes a StackSheets formula into
' the active cell.
'
' StackSheets_byJaqKennedy(blocks, names, [TopLeft], [HideBlanks]) returns
'    Sheet | Value | Address | Row | Col | RCRef
' the same shape as GridToCol and Paste Flattened List With Formatting, so the
' three outputs are interchangeable downstream.
'
' NOTE ON [HideBlanks], measured during the lambda review: it hides blanks AND
' ZEROS - the filter is (val<>"")*(val<>0). On a numeric grid a real zero
' disappears. That is why the plain version is the default command and the
' hiding one is separate and says so.
'==============================================================================

' --- CONSTANTS (module) ---
Private Const m_DEBUG_MODE As Boolean = False


'------------------------------------------------------------------------------
' Copy the sheet-name list, select the output cell, run this.
'------------------------------------------------------------------------------
Public Sub StackSheetsFromClipboard()
    WriteStackFormula False
End Sub


'------------------------------------------------------------------------------
' Same, but hides blank cells. NOTE it hides zeros too - see the module header.
'------------------------------------------------------------------------------
Public Sub StackSheetsFromClipboardHideBlanks()
    WriteStackFormula True
End Sub


'------------------------------------------------------------------------------
' Shared builder.
'------------------------------------------------------------------------------
Private Sub WriteStackFormula(ByVal blnHideBlanks As Boolean)

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "StackSheetsFromClipboard"

    Dim wb          As Workbook
    Dim rngNames    As Range
    Dim colSheets   As Collection
    Dim strFormula  As String
    Dim strNames    As String
    Dim strBlock    As String
    Dim strTopLeft  As String
    Dim rngTarget   As Range
    Dim lngErrNum   As Long
    Dim strErrDesc  As String

    On Error GoTo ErrHandler

    Set rngTarget = ActiveCell
    Set wb = rngTarget.Worksheet.Parent

    ' The sheet names come from the clipboard, reusing the same copied-range
    ' resolver that Paste Flattened List With Formatting uses.
    Set rngNames = modFlatten.CopiedRange()
    If rngNames Is Nothing Then
        rngTarget.value = "STACK SHEETS: copy the sheet-name list first, then select a cell and run this."
        GoTo Cleanup
    End If

    Set colSheets = SheetsFromNames(wb, rngNames)
    If colSheets Is Nothing Then GoTo Cleanup
    If colSheets.Count = 0 Then
        rngTarget.value = "STACK SHEETS: none of the copied names matched a sheet in this workbook."
        GoTo Cleanup
    End If

    strBlock = BoundingBlock(colSheets, strTopLeft)
    If Len(strBlock) = 0 Then
        rngTarget.value = "STACK SHEETS: those sheets have no used range to stack."
        GoTo Cleanup
    End If

    strNames = NamesArrayLiteral(colSheets)

    ' A 3D reference (First:Last!Block) is what VSTACK expands into one tall block.
    strFormula = "=StackSheets_byJaqKennedy(VSTACK('" & colSheets(1).Name & ":" & _
                 colSheets(colSheets.Count).Name & "'!" & strBlock & ")," & _
                 strNames & ",""" & strTopLeft & """"
    If blnHideBlanks Then strFormula = strFormula & ",1"
    strFormula = strFormula & ")"

    rngTarget.Formula2 = strFormula
    Application.CutCopyMode = False

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    On Error Resume Next
    ActiveCell.value = "STACK SHEETS failed: " & lngErrNum & " - " & strErrDesc
    LogError PROC_NAME, lngErrNum, strErrDesc
    Resume Cleanup

End Sub


'------------------------------------------------------------------------------
' Turn the copied cells into the ordered list of worksheets to stack.
' Exactly two names  -> everything from the first to the last in TAB ORDER.
' Anything else      -> those sheets, in the order copied.
' Works whether the clipboard list is a row or a column.
'------------------------------------------------------------------------------
Private Function SheetsFromNames(ByVal wb As Workbook, ByVal rngNames As Range) As Collection

    Dim colOut   As Collection
    Dim colRaw   As Collection
    Dim rngCell  As Range
    Dim strName  As String
    Dim lngFrom  As Long
    Dim lngTo    As Long
    Dim lngI     As Long
    Dim ws       As Worksheet

    On Error GoTo ErrHandler

    Set colRaw = New Collection
    ' For Each walks a row or a column the same way, so no orientation test is needed.
    For Each rngCell In rngNames.Cells
        strName = Trim$(CStr(rngCell.Value2 & ""))
        If Len(strName) > 0 Then colRaw.Add strName
    Next rngCell

    Set colOut = New Collection

    If colRaw.Count = 2 Then
        lngFrom = SheetIndex(wb, colRaw(1))
        lngTo = SheetIndex(wb, colRaw(2))
        If lngFrom = 0 Or lngTo = 0 Then
            Set SheetsFromNames = colOut
            Exit Function
        End If
        If lngTo < lngFrom Then
            lngI = lngFrom
            lngFrom = lngTo
            lngTo = lngI
        End If
        For lngI = lngFrom To lngTo
            Set ws = wb.Worksheets(lngI)
            If ws.Visible = xlSheetVisible Then colOut.Add ws
        Next lngI
    Else
        For lngI = 1 To colRaw.Count
            On Error Resume Next
            Set ws = Nothing
            Set ws = wb.Worksheets(colRaw(lngI))
            On Error GoTo ErrHandler
            If Not ws Is Nothing Then colOut.Add ws
        Next lngI
    End If

    Set SheetsFromNames = colOut
    Exit Function

ErrHandler:
    Set SheetsFromNames = New Collection

End Function


Private Function SheetIndex(ByVal wb As Workbook, ByVal strName As String) As Long

    Dim lngI As Long

    For lngI = 1 To wb.Worksheets.Count
        If StrComp(wb.Worksheets(lngI).Name, strName, vbTextCompare) = 0 Then
            SheetIndex = lngI
            Exit Function
        End If
    Next lngI

End Function


'------------------------------------------------------------------------------
' The smallest block covering every listed sheet's used range, as "B2:D10".
' strTopLeft is returned too, because StackSheets needs it to rebuild addresses.
'------------------------------------------------------------------------------
Private Function BoundingBlock(ByVal colSheets As Collection, _
                               ByRef strTopLeft As String) As String

    Dim ws       As Worksheet
    Dim lngTop   As Long
    Dim lngLeft  As Long
    Dim lngBot   As Long
    Dim lngRight As Long
    Dim rngUsed  As Range

    On Error GoTo ErrHandler

    lngTop = 0

    For Each ws In colSheets
        Set rngUsed = Nothing
        On Error Resume Next
        Set rngUsed = ws.UsedRange
        On Error GoTo ErrHandler
        If Not rngUsed Is Nothing Then
            If lngTop = 0 Then
                lngTop = rngUsed.Row
                lngLeft = rngUsed.Column
                lngBot = rngUsed.Row + rngUsed.Rows.Count - 1
                lngRight = rngUsed.Column + rngUsed.Columns.Count - 1
            Else
                If rngUsed.Row < lngTop Then lngTop = rngUsed.Row
                If rngUsed.Column < lngLeft Then lngLeft = rngUsed.Column
                If rngUsed.Row + rngUsed.Rows.Count - 1 > lngBot Then _
                    lngBot = rngUsed.Row + rngUsed.Rows.Count - 1
                If rngUsed.Column + rngUsed.Columns.Count - 1 > lngRight Then _
                    lngRight = rngUsed.Column + rngUsed.Columns.Count - 1
            End If
        End If
    Next ws

    If lngTop = 0 Then Exit Function

    strTopLeft = Cells(lngTop, lngLeft).Address(False, False)
    BoundingBlock = strTopLeft & ":" & Cells(lngBot, lngRight).Address(False, False)
    Exit Function

ErrHandler:
    BoundingBlock = vbNullString

End Function


'------------------------------------------------------------------------------
' The sheet names as an Excel array literal, {"A";"B";"C"}, so the formula does
' not depend on the clipboard range still existing afterwards.
'------------------------------------------------------------------------------
Private Function NamesArrayLiteral(ByVal colSheets As Collection) As String

    Dim ws     As Worksheet
    Dim strOut As String

    For Each ws In colSheets
        If Len(strOut) > 0 Then strOut = strOut & ";"
        strOut = strOut & """" & Replace(ws.Name, """", """""") & """"
    Next ws

    NamesArrayLiteral = "{" & strOut & "}"

End Function


