Attribute VB_Name = "modCaseSetup"
Option Explicit
Option Base 1

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
Sub Setup()
    'backup all sheets
    Call Backup
    Call CreateLevelSheets
    Call CreateCaseInputsSheet("detailed")
End Sub

'backup all sheets (to prevent overwrite errors)
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Backup sheets in workbook
' Description:            Copy all sheets in workbook
' Macro Expression:       modCaseSetup.Backup()
' Generated:              11/15/2024 08:44 PM
'----------------------------------------------------------------------------------------------------
Sub Backup()
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim intCnt As Integer
    Dim intCurr As Integer
    Dim strName As String
    
    Set WB = ActiveWorkbook
    intCnt = WB.Sheets.Count
    'only cycle through existing sheets
    For intCurr = 1 To intCnt
        'get sheet name and copy to end
        strName = WB.Worksheets(intCurr).Name
        WB.Worksheets(intCurr).Copy After:=WB.Worksheets(WB.Worksheets.Count)
        'rename new sheet as BU version
        Set WS = WB.Worksheets(WB.Worksheets.Count)
        If Len(strName) < 30 Then
            WS.Name = strName & "BU"
        Else
            WS.Name = Left(strName, 29) & "BU"
        End If
    Next intCurr
    
End Sub

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create level sheets
' Description:            Creates a sheet for each level in the Case sheet
' Macro Expression:       modCaseSetup.CreateLevelSheets()
' Generated:              11/15/2024 09:59 PM
'----------------------------------------------------------------------------------------------------
Sub CreateLevelSheets()
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim WSNew As Worksheet
    Dim strWS As String
    Dim intLevels As Integer
    Dim intCurr As Integer
    Dim intRow As Integer
    Dim intBegRow As Integer
    Dim intEndRow As Integer
    Dim colLRows As Collection
    Dim rng As Range
    Dim rngC2 As Range
    Dim rngAnswers As Range
    Dim rngNumbers As Range
    Dim objCell As Object
    Dim rngScore As Range
    
    VBAInit
    Set WB = ActiveWorkbook
    On Error Resume Next
    Set WS = WB.Worksheets("Case")
    On Error GoTo 0
    
    'check if sheet exists, if not, prompt for a sheet name,
    'if this doesn't exist, exit sub
    If WS Is Nothing Then
        strWS = InputBox("What is the case sheet called?")
        On Error Resume Next
        Set WS = WB.Worksheets(strWS)
        On Error GoTo 0
        If WS Is Nothing Then
            MsgBox ("Sheet not found")
            VBAFin
            Exit Sub
        End If
    End If
    
    '___________________worksheet exists_____________________
    'get current score cell
    On Error Resume Next
    Set rngScore = WS.UsedRange.Find("Current Score").Offset(1, 0)
    On Error GoTo 0
    
    'cycle through column B looking for "Level *" cells
    Set colLRows = New Collection
    
    Set rng = WS.UsedRange
    Set rngC2 = WS.Cells(1, 2).Resize(rng.Rows.Count, 1)
    For Each objCell In rngC2
        If Not IsEmpty(objCell.Value) Then
            If objCell.Value Like "Level *" And objCell.Value <> "Level Code" Then
                colLRows.Add objCell.Row
            End If
        End If
    Next objCell
    
    'count levels
    intLevels = colLRows.Count
    
    'loop through the levels to create new sheets
    For intCurr = 1 To intLevels
        'create new worksheet for the level
        Set WSNew = WB.Sheets.Add(After:=WB.Sheets(WS.Index + intCurr - 1))
        WSNew.Name = "L" & intCurr
        
        'get row numbers
        intBegRow = colLRows(intCurr)
        If intCurr < intLevels Then
            intEndRow = colLRows(intCurr + 1) - 1
        Else
            'find last row in column B
            intEndRow = WS.Cells(WS.Rows.Count, 2).End(xlUp).Row
        End If
        
        'copy rows to new worksheet
        WS.Rows(intBegRow & ":" & intEndRow).Copy Destination:=WSNew.Rows(1)
        
        'link answer cells to new cells in new sheet
        For intRow = intBegRow To intEndRow
            Set rngNumbers = WS.Cells(intRow, 2)
            Set rngAnswers = WS.Cells(intRow, 5)
            If Not IsEmpty(rngNumbers) And IsNumeric(rngNumbers) And IsEmpty(rngAnswers.Value) Then
                rngAnswers.Formula = "='" & WSNew.Name & "'!" & rngAnswers.Offset(1 - intBegRow).Address
            End If
            On Error GoTo SkipStep
            If WSNew.Cells(intRow - intBegRow + 1, 6) <> "" Then
                WSNew.Cells(intRow - intBegRow + 1, 6).Formula = "='Case'!" & rngAnswers.Offset(0, 1).Address
            End If
SkipStep:
            
        Next intRow
        'for workbooks with answers (for training), add the score at the top of each sheet
        If Not rngScore Is Nothing Then
            WSNew.Cells(1, 1).Formula = "='Case'!" & rngScore.Address
        End If
        
    Next intCurr
    VBAFin
End Sub



'creates a data table for the current level
'Assumes either in a level sheet named in the format "L#" or in a cell with value "Example#"
'use [NewTableTargetToRight] to point at target range from command
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Data Table
' Description:            Creates a data table on the current sheet for level sheets or activecell = "Example#"
' Macro Expression:       modCaseSetup.CreateDataTable([[NewTableTargetToRight]])
' Generated:              01/05/2025 04:54 PM
'----------------------------------------------------------------------------------------------------
Sub CreateDataTable(rngTarget As Range)
    Dim rngAC As Range
    Dim strAC As String
    Dim WS As Worksheet
    Dim strWS As String
    Dim strCV As String
    
    'required info for table data
    Dim intLevel As Integer
    Dim intQs() As Integer
    
    'Level sheet setup
    Dim booLS As Boolean
    Dim intEndRow As Integer
    Dim intRow As Integer
    Dim intQCnt As Integer
    Dim intCurrQ As Integer
    Dim intEndCol As Integer
    Dim intCol As Integer
    Dim strCol As String
    Dim strHdr As String
    Dim intExRow As Integer
    Dim intHdRow As Integer
    Dim intACol As Integer
    Dim strVal As String
    Dim intInCnt As Integer
    Dim intBlnk As Integer
    
    VBAInit
    
    Set rngAC = ActiveCell
    strAC = rngAC.Value2
    Set WS = ActiveSheet
    strWS = WS.Name
    booLS = False
    
    'make sure target row is row 2 or higher
    If rngTarget.Row = 1 Then
        Set rngTarget = WS.Cells(2, rngTarget.Column)
    End If
    
    'get question list
    'if this is a level sheet, get level name
    If Len(strWS) <= 3 And Left(strWS, 1) = "L" Then
        booLS = True
        intLevel = CInt(Right(strWS, Len(strWS) - 1))
        
        'find last row in column B
        intEndRow = WS.Cells(WS.Rows.Count, 2).End(xlUp).Row
        
        intQCnt = 0
        'cycle through column B to get question numbers
        For intRow = 1 To intEndRow
            strCV = WS.Cells(intRow, 2).Value
            If IsNumeric(strCV) Then
                intQCnt = intQCnt + 1
                ReDim Preserve intQs(intQCnt)
                intQs(intQCnt) = strCV
            End If
        Next intRow
    Else
        'if not a level sheet, check first cell is "Example*" and generate information from there
        If Left(strAC, 7) = "Example" Then
            On Error Resume Next
            intLevel = CInt(Trim(Replace(strAC, "Example", "")))
            If Err.Number <> 0 Then
                MsgBox ("Err " & Err.Number & ": " & Err.Description & vbCrLf & "Unable to setup data table.")
                VBAFin
                Exit Sub
            End If
            On Error GoTo 0
            'cycle through cells below "Example" cell for question numbers
            intRow = 1
            intQCnt = 0
            strCV = WS.Cells(rngAC.Row + intRow, rngAC.Column).Value2
            Do Until Left(strCV, 7) = "Example" Or strCV = "Level Code" Or strCV = "Game #" Or intRow = 200
                If IsNumeric(strCV) Then
                    intQCnt = intQCnt + 1
                    ReDim Preserve intQs(intQCnt)
                    intQs(intQCnt) = strCV
                End If
                intRow = intRow + 1
                strCV = WS.Cells(rngAC.Row + intRow, rngAC.Column).Value2
            Loop
        Else
            MsgBox ("Please select an 'Example#' cell or a level sheet as 'L#'" & vbCrLf & "Unable to setup data table.")
            VBAFin
            Exit Sub
        End If
    End If
    
    'generate data table
    
    'find example row and get last column - assumes example row is always in column B
    'cycle through to find "Example" cell for input data
    If booLS Then
        intRow = 1
        strCV = WS.Cells(intRow, 2).Value2
        Do Until strCV = "Example" & intLevel Or intRow = 1000
            intRow = intRow + 1
            strCV = WS.Cells(intRow, 2).Value2
        Loop
        If intRow = 1000 Then
            MsgBox ("Error finding Example cell in column B" & vbCrLf & "Unable to setup data table.")
            VBAFin
            Exit Sub
        End If
        intExRow = intRow
    Else
        intExRow = rngAC.Row
    End If
    intHdRow = intExRow - 2
    intEndCol = WS.Cells(intExRow, 16300).End(xlToLeft).Column
    
    'add InputCell and Data loads
    intACol = 2
    intBlnk = 0
    rngTarget.Value = "Example" & intLevel
    For intCol = 3 To intEndCol
        strHdr = WS.Cells(intHdRow, intCol)
        strVal = WS.Cells(intExRow, intCol)
        If strHdr = "Answer" Then intACol = intCol
        If strHdr <> "Level" And strHdr <> "Points" And strHdr <> "Answer" And strHdr <> "Game #" Then
            If strVal <> "" Then
                intBlnk = 0
                strCol = Replace(Replace(Cells(1, intCol).Address, "1", ""), "$", "")
                intInCnt = intInCnt + 1
                'header
                WS.Cells(intHdRow, intCol).Copy
                WS.Cells(rngTarget.Row - 1, rngTarget.Column + intInCnt).PasteSpecial xlPasteAll
                Application.CutCopyMode = False
                'lookup data
                WS.Cells(rngTarget.Row, rngTarget.Column + intInCnt).Formula = "=XLOOKUP(" & _
                rngTarget.Address(1, 1) & ",B:B," & strCol & ":" & strCol & ",,0)"
                WS.Cells(intExRow, intCol).Copy
                WS.Cells(rngTarget.Row, rngTarget.Column + intInCnt).PasteSpecial xlPasteFormats
                Application.CutCopyMode = False
            Else
                'if the column is blank, check how many blank columns have occurred
                'if there are 5 or more, assume there are calcs to the right and stop linking data columns
                intBlnk = intBlnk + 1
                If intBlnk >= 5 Then GoTo NoMoreInputs
            End If
        End If
    Next intCol
NoMoreInputs:
    
    'set up formula and question numbers
    WS.Cells(rngTarget.Row + 3, rngTarget.Column + 1).Formula = "=" & rngTarget.Address
    For intRow = LBound(intQs) To UBound(intQs)
        WS.Cells(rngTarget.Row + 3 + intRow + (1 - LBound(intQs)), rngTarget.Column).Value = intQs(intRow)
    Next intRow
    
    'set up table
    Range(WS.Cells(rngTarget.Row + 3, rngTarget.Column), _
    WS.Cells(rngTarget.Row + 3 + UBound(intQs) + (1 - LBound(intQs)), rngTarget.Column + 1)).Table ColumnInput:=rngTarget
    
    'points answers at table
    If intACol > 2 Then
        'cycle through column B to get question numbers
        intCurrQ = 0
        For intRow = 1 To intEndRow
            strCV = WS.Cells(intRow, 2).Value
            If IsNumeric(strCV) Then
                intCurrQ = intCurrQ + 1
                WS.Cells(intRow, intACol).Formula = "=" & WS.Cells(rngTarget.Row + 3 + intCurrQ, rngTarget.Column + 1).Address(0, 0)
            End If
        Next intRow
    End If
    
    Erase intQs
    VBAFin
End Sub

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
Sub CreateCaseInputsSheet(Optional strDetailed As String = "")
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim WSNew As Worksheet
    Dim strWS As String
    Dim booDetailed As Boolean
    
    Set WB = ActiveWorkbook
    On Error Resume Next
    Set WS = WB.Worksheets("Case")
    On Error GoTo 0
    booDetailed = True
    If IsMissing(strDetailed) Or strDetailed = "" Then booDetailed = False
    
    'check if sheet exists
    If WS Is Nothing Then
        strWS = InputBox("What is the case sheet called?")
        On Error Resume Next
        Set WS = WB.Worksheets(strWS)
        On Error GoTo 0
        If WS Is Nothing Then
            MsgBox ("Sheet not found")
            Exit Sub
        End If
    End If
    
    'worksheet exists
    Set WSNew = WB.Sheets.Add(After:=WB.Sheets(WS.Index + 1))
    On Error Resume Next
    WSNew.Name = "CaseInputs"
    On Error GoTo 0
    'freeze panes below header
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 2
        .SplitRow = 2
        .FreezePanes = True
    End With
    If Not booDetailed Then
        WSNew.Cells(2, 2).Formula2 = "=CaseInputs()"
    Else
        Call DetailedInputs(WS, WSNew)
    End If
End Sub

'more detailed breakdown of inputs for non-standard layouts
'this will run
Sub DetailedInputs(WS As Worksheet, WSNew As Worksheet)
    Dim rngTarget As Range
    Dim rngHdr As Range
    Dim intIncCols() As Integer
    Dim intInc As Integer
    Dim strHeaders() As String
    Dim intCols() As Integer
    
    'for cycling through WB
    Dim intEndRow As Integer
    Dim intEndCol As Integer
    Dim intCurrExRow As Integer
    Dim strCV As String
    Dim intHdr As Integer
    Dim booHdr As Boolean
    Dim intRowOut As Integer
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strHdr As String
    
    Set rngHdr = WSNew.Cells(2, 3)
    Set rngTarget = WSNew.Cells(3, 3)
    intInc = 0
    intRowOut = 0
    
    'find last row in column B
    intEndRow = WS.Cells(WS.Rows.Count, 2).End(xlUp).Row
    intEndCol = WS.UsedRange.Columns.Count + 3
    
    'cycle through rows in input sheet
    For intRow = 1 To intEndRow
        strCV = WS.Cells(intRow, 2).Value
        If Left(strCV, 7) = "Example" Then
            intCurrExRow = intRow
            'example row - check current headers list vs existing
            'reset current column numbers
            If intInc > 0 Then
                ReDim intIncCols(1 To intInc)
                For intHdr = 1 To intInc
                    intIncCols(intHdr) = 0
                Next intHdr
            End If
            For intCol = 3 To intEndCol
                strHdr = WS.Cells(intCurrExRow - 2, intCol).Value
                strCV = WS.Cells(intCurrExRow, intCol).Value
                If strCV <> "" And strHdr <> "Level" And strHdr <> "Points" And strHdr <> "Answer" Then
                    'first header of first example
                    If intInc = 0 Then
                        intInc = intInc + 1
                        ReDim strHeaders(1 To intInc)
                        strHeaders(intInc) = strHdr
                        'add header to output sheet
                        If strHdr <> "" Then WSNew.Cells(rngHdr.Row, rngHdr.Column + intInc - 1).Value = strHdr
                        ReDim intIncCols(1 To intInc)
                        intIncCols(intInc) = intCol
                    Else
                        'find current header match and update column to output from
                        booHdr = False
                        For intHdr = 1 To intInc
                            If strHeaders(intHdr) = strHdr And intIncCols(intHdr) = 0 Then
                                intIncCols(intHdr) = intCol
                                booHdr = True
                                Exit For
                            End If
                        Next intHdr
                        If Not booHdr Then
                            intInc = intInc + 1
                            ReDim Preserve strHeaders(intInc)
                            ReDim Preserve intIncCols(intInc)
                            strHeaders(intInc) = strHdr
                            intIncCols(intInc) = intCol
                            'add header to output sheet
                            If strHdr <> "" Then WSNew.Cells(rngHdr.Row, rngHdr.Column + intInc - 1).Value = strHdr
                        End If
                    End If
                End If
            Next intCol
            
        ElseIf IsNumeric(strCV) Then
            'load values into current output sheet
            intRowOut = intRowOut + 1
            WSNew.Cells(rngTarget.Row + intRowOut - 1, rngTarget.Column - 1).Value = strCV
            For intHdr = 1 To intInc
                If intIncCols(intHdr) <> 0 Then
                    WSNew.Cells(rngTarget.Row + intRowOut - 1, rngTarget.Column + intHdr - 1).Value = WS.Cells(intRow, intIncCols(intHdr)).Value
                End If
            Next intHdr
        End If
    Next intRow
    WSNew.UsedRange.Columns.AutoFit
    
End Sub
