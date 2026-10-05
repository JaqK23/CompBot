Attribute VB_Name = "modGoToSpecial"
Option Explicit

Private Const m_DEBUG_MODE As Boolean = False

'--------------------------------------------< OA Robot >--------------------------------------------
' Function:             GotoSimilarBackgroundColor
' Created By:           Erik Oehm
' Source:               https://github.com/ExcelRobot/MEWC-Robot/blob/main/MEWC%20Robot.xlsm
'----------------------------------------------------------------------------------------------------
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Goto Similar Background Color
' Description:            Goto similar background color.
' Macro Expression:       modGoToSpecial.GotoSimilarBackgroundColor()
' Generated:              01/08/2025 01:28 PM
'----------------------------------------------------------------------------------------------------
Sub GotoSimilarBackgroundColor()

    Dim nColorIndex As Long
    Dim dTintShade As Double
    Dim nCtr As Long
    Dim rngOldActive As Range
    Dim rngOldSelection As Range
    Dim rngNewSelection As Range
    
    Set rngOldActive = ActiveCell
    Set rngOldSelection = SelectionOrUsedRange(Selection)
    nColorIndex = ActiveCell.Interior.ColorIndex
    dTintShade = Round(ActiveCell.Interior.TintAndShade, 3)
    
    Dim rngArea As Range
    For Each rngArea In rngOldSelection.Areas
        For nCtr = 1 To rngArea.Cells.Count
            If rngArea.Cells(nCtr).Interior.ColorIndex = nColorIndex And Round(rngArea.Cells(nCtr).Interior.TintAndShade, 3) = dTintShade Then
                If rngNewSelection Is Nothing Then
                    Set rngNewSelection = rngArea.Cells(nCtr)
                Else
                    Set rngNewSelection = Union(rngNewSelection, rngArea.Cells(nCtr))
                End If
            End If
        Next nCtr
    Next rngArea
    
    If Not rngNewSelection Is Nothing Then
        rngNewSelection.Select
        If Not Intersect(rngNewSelection, rngOldActive) Is Nothing Then
            rngOldActive.Activate
        End If
    End If

End Sub

Private Function SelectionOrUsedRange(vSelection As Variant) As Range
    If TypeName(vSelection) <> "Range" Then
        Set SelectionOrUsedRange = ActiveSheet.UsedRange
    ElseIf vSelection.Cells.Count = 1 Then
        Set SelectionOrUsedRange = vSelection.Parent.UsedRange
    ElseIf Not Intersect(vSelection, vSelection.Parent.UsedRange) Is Nothing Then
        Set SelectionOrUsedRange = Intersect(vSelection, vSelection.Parent.UsedRange)
    Else
        Set SelectionOrUsedRange = vSelection.Parent.UsedRange
    End If
End Function

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Fill Similar Background Color
' Description:            Fill similar background color.
' Macro Expression:       modGoToSpecial.FillSimilarBackgroundColor()
' Generated:              2026-07-18 07:26 PM
'----------------------------------------------------------------------------------------------------
' 2026-09-25 (Jaq): the SELECTED cells are the keys, one per colour, and each key's colour is
' searched for over the whole used range. It used to fall back to the used range as the KEYS when
' one cell was selected, so every cell of the sheet was run as a key: on a 59x85 map that hung
' Excel for minutes and wrote one neighbour's value over every colour group.
Sub FillSimilarBackgroundColor()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "FillSimilarBackgroundColor"
    Const MAX_KEYS  As Long = 50

    Dim rngKeys     As Range
    Dim rngCell     As Range
    Dim strSelected As String
    Dim lngErrNum   As Long
    Dim strErrDesc  As String

    On Error GoTo ErrHandler

    If TypeName(Selection) <> "Range" Then
        Application.StatusBar = "Fill Similar Background Color: select one key cell per color first"
        GoTo Cleanup
    End If
    Set rngKeys = Selection
    If rngKeys.Cells.CountLarge > MAX_KEYS Then
        Application.StatusBar = "Fill Similar Background Color: " & rngKeys.Cells.CountLarge & _
                                " cells selected - select one key cell per color (" & MAX_KEYS & " max)"
        GoTo Cleanup
    End If

    For Each rngCell In rngKeys.Cells
        strSelected = CStr(rngCell.Value2)
        If strSelected = "" Then strSelected = CStr(rngCell.Offset(0, 1).Value2)
        ' One cell selected: GotoSimilarBackgroundColor searches the whole used range.
        rngCell.Select
        Call GotoSimilarBackgroundColor
        Selection.Value2 = strSelected
    Next rngCell

    rngKeys.Select
    Application.StatusBar = "Fill Similar Background Color: " & rngKeys.Cells.Count & " color(s) filled"

Cleanup:
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    Application.StatusBar = "Fill Similar Background Color failed: " & lngErrNum & " - " & strErrDesc
    LogError PROC_NAME, lngErrNum, strErrDesc
    Resume Cleanup

End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Copy Map And Fill From Legend
' Description:            From a colored legend cell: copies the sheet, fills every map cell on the copy
'                         with its colour's legend value, and adds clr_ names for the copy's ranges.
' Macro Expression:       modGoToSpecial.CopyMapFillFromLegend()
'----------------------------------------------------------------------------------------------------
' 2026-10-05 (Jaq): the usual first step on a colour-coded map, in one run. The legend is the
' unbroken run of coloured cells through the active cell (down the column, else along the row), or
' the selection if several cells are selected. A legend cell's value is its key, or the value to its
' right when the swatch itself is empty. White and no-fill are NEVER keys, so it cannot fill the
' whole sheet. Colours are matched exactly (Interior.Color) in one pass over the copy's used range.
' Every workbook-level name on the original sheet gets a clr_ twin pointing at the copy.
Sub CopyMapFillFromLegend()

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "CopyMapFillFromLegend"
    Const MAX_KEYS  As Long = 50
    Const PREFIX    As String = "clr_"

    Dim wsSrc        As Worksheet
    Dim wsNew        As Worksheet
    Dim rngKeys      As Range
    Dim rngCell      As Range
    Dim objKeys      As Object
    Dim objHits      As Object
    Dim varKey       As Variant
    Dim varVal       As Variant
    Dim lngColour    As Long
    Dim lngFilled    As Long
    Dim lngNames     As Long
    Dim lngIdx       As Long
    Dim strAddr      As String
    Dim strNames()   As String
    Dim strAddrs()   As String
    Dim nmItem       As Name
    Dim blnScreen    As Boolean
    Dim lngErrNum    As Long
    Dim strErrDesc   As String

    On Error GoTo ErrHandler
    blnScreen = Application.ScreenUpdating

    If TypeName(Selection) <> "Range" Then
        Application.StatusBar = "Copy Map And Fill From Legend: stand on a colored legend cell first"
        GoTo Cleanup
    End If
    Set wsSrc = ActiveSheet
    If Selection.Cells.CountLarge > 1 Then
        Set rngKeys = Selection
    Else
        Set rngKeys = LegendKeys(ActiveCell)
    End If
    If rngKeys Is Nothing Then
        Application.StatusBar = "Copy Map And Fill From Legend: stand on a colored legend cell first"
        GoTo Cleanup
    End If
    If rngKeys.Cells.CountLarge > MAX_KEYS Then
        Application.StatusBar = "Copy Map And Fill From Legend: " & rngKeys.Cells.CountLarge & _
                                " legend cells - select the legend only (" & MAX_KEYS & " max)"
        GoTo Cleanup
    End If

    ' colour -> legend value; white and no-fill never become keys
    Set objKeys = CreateObject("Scripting.Dictionary")
    For Each rngCell In rngKeys.Cells
        If IsKeyColour(rngCell) Then
            lngColour = rngCell.Interior.Color
            varVal = rngCell.Value2
            If IsEmpty(varVal) Then varVal = rngCell.Offset(0, 1).Value2
            If Not objKeys.Exists(lngColour) Then objKeys.Add lngColour, varVal
        End If
    Next rngCell
    If objKeys.Count = 0 Then
        Application.StatusBar = "Copy Map And Fill From Legend: no colored legend cells found"
        GoTo Cleanup
    End If

    ' names on the original sheet, collected before anything is added
    ReDim strNames(0 To 0)
    ReDim strAddrs(0 To 0)
    For Each nmItem In wsSrc.Parent.Names
        strAddr = vbNullString
        If TypeName(nmItem.Parent) = "Workbook" Then
            On Error Resume Next
            If nmItem.RefersToRange.Worksheet.Name = wsSrc.Name Then strAddr = nmItem.RefersToRange.Address
            On Error GoTo ErrHandler
        End If
        If Len(strAddr) > 0 And Left$(nmItem.Name, Len(PREFIX)) <> PREFIX Then
            lngIdx = UBound(strNames) + 1
            ReDim Preserve strNames(0 To lngIdx)
            ReDim Preserve strAddrs(0 To lngIdx)
            strNames(lngIdx) = nmItem.Name
            strAddrs(lngIdx) = strAddr
        End If
    Next nmItem

    Application.ScreenUpdating = False
    wsSrc.Copy After:=wsSrc
    Set wsNew = wsSrc.Parent.Worksheets(wsSrc.Index + 1)
    wsNew.Name = UniqueSheetName(wsSrc.Parent, PREFIX & wsSrc.Name)

    ' one pass over the copy: group cells by colour, then write each group once
    Set objHits = CreateObject("Scripting.Dictionary")
    For Each rngCell In wsNew.UsedRange.Cells
        If rngCell.Interior.Pattern <> xlNone Then
            lngColour = rngCell.Interior.Color
            If objKeys.Exists(lngColour) Then
                If objHits.Exists(lngColour) Then
                    Set objHits(lngColour) = Union(objHits(lngColour), rngCell)
                Else
                    objHits.Add lngColour, rngCell
                End If
            End If
        End If
    Next rngCell
    For Each varKey In objHits.Keys
        objHits(varKey).Value2 = objKeys(varKey)
        lngFilled = lngFilled + objHits(varKey).Cells.Count
    Next varKey

    For lngIdx = 1 To UBound(strNames)
        wsSrc.Parent.Names.Add Name:=PREFIX & strNames(lngIdx), _
            RefersTo:="='" & Replace(wsNew.Name, "'", "''") & "'!" & strAddrs(lngIdx)
        lngNames = lngNames + 1
    Next lngIdx

    wsNew.Activate
    Application.StatusBar = "Copy Map And Fill From Legend: " & wsNew.Name & ", " & objKeys.Count & _
                            " color(s), " & lngFilled & " cells filled, " & lngNames & " clr_ name(s)"

Cleanup:
    Application.ScreenUpdating = blnScreen
    Exit Sub

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    Application.StatusBar = "Copy Map And Fill From Legend failed: " & lngErrNum & " - " & strErrDesc
    LogError PROC_NAME, lngErrNum, strErrDesc
    Resume Cleanup

End Sub

' Purpose: TRUE for a cell with a real fill colour; white and no-fill are never legend keys.
Private Function IsKeyColour(ByVal rngCell As Range) As Boolean
    On Error GoTo ErrHandler
    IsKeyColour = (rngCell.Interior.Pattern <> xlNone) And (rngCell.Interior.Color <> RGB(255, 255, 255))
    Exit Function
ErrHandler:
    IsKeyColour = False
End Function

' Purpose: the unbroken run of coloured cells through rngStart, down its column; if that is the cell
'          alone, along its row. Nothing when rngStart itself is not coloured.
Private Function LegendKeys(ByVal rngStart As Range) As Range

    Dim rngFirst As Range
    Dim rngLast  As Range

    On Error GoTo ErrHandler
    If Not IsKeyColour(rngStart) Then Exit Function
    Set rngFirst = rngStart
    Do While rngFirst.Row > 1
        If Not IsKeyColour(rngFirst.Offset(-1, 0)) Then Exit Do
        Set rngFirst = rngFirst.Offset(-1, 0)
    Loop
    Set rngLast = rngStart
    Do While rngLast.Row < rngStart.Worksheet.Rows.Count
        If Not IsKeyColour(rngLast.Offset(1, 0)) Then Exit Do
        Set rngLast = rngLast.Offset(1, 0)
    Loop
    If rngFirst.Row = rngLast.Row Then
        Do While rngFirst.Column > 1
            If Not IsKeyColour(rngFirst.Offset(0, -1)) Then Exit Do
            Set rngFirst = rngFirst.Offset(0, -1)
        Loop
        Do While rngLast.Column < rngStart.Worksheet.Columns.Count
            If Not IsKeyColour(rngLast.Offset(0, 1)) Then Exit Do
            Set rngLast = rngLast.Offset(0, 1)
        Loop
    End If
    Set LegendKeys = rngStart.Worksheet.Range(rngFirst, rngLast)
    Exit Function
ErrHandler:
    Set LegendKeys = Nothing
End Function

' Purpose: strName, or strName (2), (3)... if a sheet of that name exists; at most 31 characters.
Private Function UniqueSheetName(ByVal wb As Workbook, ByVal strName As String) As String

    Dim strTry  As String
    Dim lngN    As Long
    Dim wsProbe As Object

    strTry = Left$(strName, 31)
    lngN = 1
    Do
        Set wsProbe = Nothing
        On Error Resume Next
        Set wsProbe = wb.Sheets(strTry)
        On Error GoTo 0
        If wsProbe Is Nothing Then Exit Do
        lngN = lngN + 1
        strTry = Left$(strName, 31 - Len(" (" & lngN & ")")) & " (" & lngN & ")"
    Loop
    UniqueSheetName = strTry
End Function


