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
    Call RenameSht                          'rename sheets with multiple words for easy reference
    Call Backup                             'backup all sheets
    Call NameAllUsedRanges                  'set up named ranges on input sheets
    Call CreateLevelSheets                  'create a sheet for each level
    Call CreateBonusSheet                   'create a bonus sheet
    VBAInit
    Call CreateCaseInputsSheet("detailed")  'create a sheet with the inputs in "Case" sheet
    VBAFin
End Sub

'rename all sheets (to shorter names)
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Rename Sheets
' Description:            Rename multi-word sheet names
' Macro Expression:       modCaseSetup.RenameSht()
' Generated:              01/25/2025 07:41 PM
'----------------------------------------------------------------------------------------------------
Sub RenameSht()
    Dim WB As Workbook
    Dim WS As Worksheet
    Dim intCnt As Integer
    Dim intCurr As Integer
    Dim strName As String
    Dim strNewName As String
    Dim strWords() As String
    Dim lngI
    Set WB = ActiveWorkbook
    intCnt = WB.Sheets.Count
    'only cycle through existing sheets
    For intCurr = 1 To intCnt
        'get sheet name and copy to end
        Set WS = WB.Worksheets(intCurr)
        strName = WS.Name
        If strName <> "Case" And strName <> "Case-Varsity" And strName <> "Answers" And InStr(strName, " ") + InStr(strName, "_") > 0 Then
            strWords = Split(Replace(strName, "_", " "), " ")
            strNewName = ""
            For lngI = LBound(strWords) To UBound(strWords)
                If Len(strWords(lngI)) > 0 Then
                    strNewName = strNewName & UCase(Left(strWords(lngI), 1))
                End If
            Next lngI
            On Error Resume Next
            WS.Name = strNewName
            On Error GoTo 0
        End If
    Next intCurr
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
    Dim intACol As Integer
    Dim dicExcHdr As Object
    Dim booHdrs As Boolean
    Dim colColList As Collection
    Dim intCol As Integer
    Dim intEndCol As Integer
    Dim rngHdrRow As Range
    Dim rngCell As Range
    Dim excludedHeaders As Object
    Dim rngRow As Range
    Dim strHeader As String
    Dim varCol As Variant
    Dim strCurrVal As String
    Dim shp As Shape

    'Initialize
    VBAInit
    ' Initialize excluded headers as a dictionary
    Set excludedHeaders = CreateObject("Scripting.Dictionary")
    excludedHeaders.Add "Answer", True
    excludedHeaders.Add "Level", True
    excludedHeaders.Add "Points", True
    'set up workbook and worksheet
    Set WB = ActiveWorkbook
    On Error Resume Next
    Set WS = WB.Worksheets("Case")
    If Err.Number <> 0 Then
        On Error Resume Next
        Set WS = WB.Worksheets("Case-Varsity")
    End If
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
    
    'cycle through column B of case sheet to find each level
    'add level start row to collection colLRows
    Set rng = WS.UsedRange
    Set rngC2 = WS.Cells(1, 2).Resize(rng.Rows.Count, 1)
    For Each objCell In rngC2
        strCurrVal = objCell.Value
        If Not IsEmpty(objCell.Value) Then
            If (objCell.Value Like "Level *" Or objCell.Value Like "Section *") _
            And strCurrVal <> "Level Code" And _
            IsNumeric(Trim(Replace(Replace(strCurrVal, "Level", ""), "Section", ""))) Then
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
        WSNew.Name = "L0" & intCurr
        
        'get row numbers
        intBegRow = colLRows(intCurr)
        If intCurr < intLevels Then
            intEndRow = colLRows(intCurr + 1) - 1
        Else
            'find last row in column B
            intEndRow = WS.Cells(WS.Rows.Count, 2).End(xlUp).Row
        End If
        
        '____________________COPY DATA________________________________________________
        'copy rows to new worksheet
        WS.Rows(intBegRow & ":" & intEndRow).Copy Destination:=WSNew.Rows(1)
        WSNew.Calculate
        
        '____________________COLUMN SETUP WITHIN SHEETS_______________________________
        ' link answer cells to new cells in new sheet and mark input cells
        ' get column numbers for answers and inputs
        
        ' Find the strHeader row with "Answer", "Level", or "Points"
        For Each rngRow In WSNew.Rows(1 & ":" & intEndRow - intBegRow + 1)
            Set rngHdrRow = rngRow.Find(What:="Answer", LookIn:=xlValues, LookAt:=xlWhole)
            If Not rngHdrRow Is Nothing Then Exit For
            Set rngHdrRow = rngRow.Find(What:="Level", LookIn:=xlValues, LookAt:=xlWhole)
            If Not rngHdrRow Is Nothing Then Exit For
            Set rngHdrRow = rngRow.Find(What:="Points", LookIn:=xlValues, LookAt:=xlWhole)
            If Not rngHdrRow Is Nothing Then Exit For
        Next rngRow
        
        ' If no strHeader row is found, tag boolean booHdrs
        booHdrs = True
        If rngHdrRow Is Nothing Then
            booHdrs = False
        End If
        ' Identify non-matching columns after column B
        Set colColList = New Collection
        intACol = 5     'default value of 5 unless overwritten
        intEndCol = 0
        If booHdrs Then
            For intCol = 3 To WSNew.UsedRange.Columns.Count + 2 ' Start after column B
                strHeader = WSNew.Cells(rngHdrRow.Row, intCol).Value
                If strHeader = "Answer" Then intACol = intCol
                If Not excludedHeaders.Exists(strHeader) And strHeader <> "" Then
                    colColList.Add intCol
                    If intCol > intEndCol Then
                        intEndCol = intCol
                    End If
                End If
            Next intCol
        End If
        If intEndCol = 0 Then intEndCol = WSNew.UsedRange.Columns.Count + 2
        
        'loop through rows in level sheet and mark
        For intRow = intBegRow To intEndRow
            'set answer lookups in Case sheet
            Set rngNumbers = WS.Cells(intRow, 2)
            Set rngAnswers = WS.Cells(intRow, intACol)
            If IsError(WSNew.Cells(intRow - intBegRow + 1, intACol)) Then
                WSNew.Cells(intRow - intBegRow + 1, intACol).Formula = ""
            End If
            If IsError(rngAnswers) Then GoTo LinkCase
            If (Not IsEmpty(rngNumbers) And IsNumeric(rngNumbers) And _
            IsEmpty(rngAnswers.Value)) Or InStr(rngAnswers.Formula, "#REF") > 0 Then
LinkCase:
                rngAnswers.Formula = "='" & WSNew.Name & "'!" & rngAnswers.Offset(1 - intBegRow).Address
            End If
            
            'set score lookups - assumed this is the column after "Answer"
            On Error GoTo SkipStep

            If IsError(WSNew.Cells(intRow - intBegRow + 1, intACol + 1)) Then
                WSNew.Cells(intRow - intBegRow + 1, intACol + 1).Formula = "='Case'!" & rngAnswers.Offset(0, 1).Address
            ElseIf WSNew.Cells(intRow - intBegRow + 1, intACol + 1) <> "" Then
                WSNew.Cells(intRow - intBegRow + 1, intACol + 1).Formula = "='Case'!" & rngAnswers.Offset(0, 1).Address
            End If
SkipStep:
            On Error GoTo 0
            'mark as input cells for non-standard columns with values
            If intRow - intBegRow + 1 > rngHdrRow.Row Then
                For Each varCol In colColList
                    Set rngCell = WSNew.Cells(intRow - intBegRow + 1, varCol)
                    If Not IsEmpty(rngCell.Value) And rngCell.HasFormula = False Then
                        Call MarkAsInputCells(rngCell, False)
                    End If
                Next varCol
            End If

        Next intRow
        
        'for workbooks with answers (for training), add the score at the top of each sheet
        If Not rngScore Is Nothing Then
            WSNew.Cells(1, 1).Formula = "='Case'!" & rngScore.Address
        End If
        
        'clear shapes
        For Each shp In WSNew.Shapes
            shp.Delete
        Next shp

    Next intCurr
    VBAFin
    Application.Calculate
End Sub

'creates a bonus sheet with just the bonus info on
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Create Bonus Sheet
' Description:            Creates bonus sheet "B" with bonus questions
' Macro Expression:       modCaseSetup.CreateBonusSheet()
' Generated:              01/08/2025 04:37 PM
'----------------------------------------------------------------------------------------------------
Sub CreateBonusSheet()
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
    Dim intACol As Integer
    Dim dicExcHdr As Object
    Dim booHdrs As Boolean
    Dim colColList As Collection
    Dim intCol As Integer
    Dim intEndCol As Integer
    Dim rngHdrRow As Range
    Dim rngCell As Range
    Dim excludedHeaders As Object
    Dim rngRow As Range
    Dim strHeader As String
    Dim varCol As Variant
    Dim strCurrVal As String
    
    Dim intFound As Integer
    
    'Initialize
    VBAInit
    'set up workbook and worksheet
    Set WB = ActiveWorkbook
    On Error Resume Next
    Set WS = WB.Worksheets("Case")
    If Err.Number <> 0 Then
        On Error Resume Next
        Set WS = WB.Worksheets("Case-Varsity")
    End If
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
    
    'cycle through column B looking for "Bonus Questions" cell
    'cycle through column B of case sheet to find first level
    
    Set rng = WS.UsedRange
    Set rngC2 = WS.Cells(1, 2).Resize(rng.Rows.Count, 1)
    intFound = 0
    For Each objCell In rngC2
        strCurrVal = objCell.Value
        If Not IsEmpty(objCell.Value) Then
            If Trim(strCurrVal) = "Bonus Questions" Then
                intBegRow = objCell.Row
                intFound = intFound + 1
            ElseIf (intFound = 1 And (strCurrVal = "Questions" Or strCurrVal = "Levels")) Or _
            (objCell.Value Like "Level *" And strCurrVal <> "Level Code" And _
            IsNumeric(Trim(Replace(strCurrVal, "Level", "")))) Then
                intEndRow = objCell.Row
                intFound = intFound + 1
                If intFound = 2 Then Exit For
            End If
        End If
    Next objCell
    If intFound < 2 Then intBegRow = 1
    
    'create new worksheet for bonuses
    Set WSNew = WB.Sheets.Add(After:=WB.Sheets(WS.Index))
    WSNew.Name = "B"
    
    
    '____________________COPY DATA________________________________________________
    'copy rows to new worksheet
    If intBegRow <= 0 Then intBegRow = 1
    WS.Rows(intBegRow & ":" & intEndRow).Copy Destination:=WSNew.Rows(1)
    WSNew.Calculate
    
    '____________________COLUMN SETUP WITHIN SHEETS_______________________________
    ' link answer cells to new cells in new sheet and mark input cells
    ' get column numbers for answers and inputs
    
    ' Find the strHeader row with "Answer", "Level", or "Points"
    For Each rngRow In WSNew.Rows(1 & ":" & intEndRow - intBegRow + 1)
        Set rngHdrRow = rngRow.Find(What:="Answer", LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngHdrRow Is Nothing Then Exit For
        Set rngHdrRow = rngRow.Find(What:="Level", LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngHdrRow Is Nothing Then Exit For
        Set rngHdrRow = rngRow.Find(What:="Points", LookIn:=xlValues, LookAt:=xlWhole)
        If Not rngHdrRow Is Nothing Then Exit For
    Next rngRow
    
    ' If no strHeader row is found, tag boolean booHdrs
    booHdrs = True
    If rngHdrRow Is Nothing Then
        booHdrs = False
    End If
    'identify answer column and unmarked score column
    intACol = 5     'default value of 5 unless overwritten
    If booHdrs Then
        For intCol = 3 To WSNew.UsedRange.Columns.Count + 2 ' Start after column B
            strHeader = WSNew.Cells(rngHdrRow.Row, intCol).Value
            If strHeader = "Answer" Then intACol = intCol
        Next intCol
    End If
    intEndCol = WSNew.UsedRange.Columns.Count + 2
    
    'loop through rows in level sheet and mark
    For intRow = intBegRow To intEndRow
        'set answer lookups in Case sheet
        Set rngNumbers = WS.Cells(intRow, 2)
        Set rngAnswers = WS.Cells(intRow, intACol)
        If IsError(WSNew.Cells(intRow - intBegRow + 1, intACol)) Then
            WSNew.Cells(intRow - intBegRow + 1, intACol).Formula = ""
        End If
        If IsError(rngAnswers) Then GoTo LinkCase
        If (Not IsEmpty(rngNumbers) And IsNumeric(Trim(Replace(rngNumbers, "Bonus", ""))) And _
        IsEmpty(rngAnswers.Value)) Or InStr(rngAnswers.Formula, "#REF") > 0 Then
LinkCase:
            rngAnswers.Formula = "='" & WSNew.Name & "'!" & rngAnswers.Offset(1 - intBegRow).Address
        End If
        
        'set score lookups - assumed this is the column after "Answer"
        On Error GoTo SkipStep

        If IsError(WSNew.Cells(intRow - intBegRow + 1, intACol + 1)) Then
            WSNew.Cells(intRow - intBegRow + 1, intACol + 1).Formula = "='Case'!" & rngAnswers.Offset(0, 1).Address
        ElseIf WSNew.Cells(intRow - intBegRow + 1, intACol + 1) <> "" Then
            WSNew.Cells(intRow - intBegRow + 1, intACol + 1).Formula = "='Case'!" & rngAnswers.Offset(0, 1).Address
        End If
SkipStep:
        On Error GoTo 0
        
    Next intRow
    
    'for workbooks with answers (for training), add the score at the top of each sheet
    If Not rngScore Is Nothing Then
        WSNew.Cells(1, 1).Formula = "='Case'!" & rngScore.Address
    End If

    VBAFin
    Application.Calculate
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
                'strHeader
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
    
    If Not Application.Iteration Then
        Call ToggleIterativeCalculation
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
    If Err.Number <> 0 Then
        On Error Resume Next
        Set WS = WB.Worksheets("Case-Varsity")
    End If
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
    Set WSNew = WB.Sheets.Add(After:=WB.Sheets(WS.Index))
    On Error Resume Next
    WSNew.Name = "CaseInputs"
    On Error GoTo 0
    'freeze panes below strHeader
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
    Dim booVal As Boolean
    Dim intHdrOffset As Integer
    
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
        If Left(strCV, 7) = "Example" And Len(strCV) <= 10 Then
            'get header offset
            intHdrOffset = -1
            intCurrExRow = intRow
            If intCurrExRow + intHdrOffset > 0 Then
                Do While WS.Cells(intCurrExRow + intHdrOffset, 2).Value = ""
                    intHdrOffset = intHdrOffset - 1
                Loop
            End If
            'example row - check current headers list vs existing
            'reset current column numbers
            If intInc > 0 Then
                ReDim intIncCols(1 To intInc)
                For intHdr = 1 To intInc
                    intIncCols(intHdr) = 0
                Next intHdr
            End If
            For intCol = 3 To intEndCol
                strHdr = WS.Cells(intCurrExRow + intHdrOffset, intCol).Value
                strCV = WS.Cells(intCurrExRow, intCol).Value
                booVal = Not WS.Cells(intCurrExRow, intCol).HasFormula
                If booVal And strCV <> "" And strHdr <> "Level" And strHdr <> "Points" And strHdr <> "Answer" Then
                    'first strHeader of first example
                    If intInc = 0 Then
                        intInc = intInc + 1
                        ReDim strHeaders(1 To intInc)
                        strHeaders(intInc) = strHdr
                        'add strHeader to output sheet
                        If strHdr <> "" Then WSNew.Cells(rngHdr.Row, rngHdr.Column + intInc - 1).Value = strHdr
                        ReDim intIncCols(1 To intInc)
                        intIncCols(intInc) = intCol
                    Else
                        'find current strHeader match and update column to output from
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
                            'add strHeader to output sheet
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
                    WSNew.Cells(rngTarget.Row + intRowOut - 1, rngTarget.Column + intHdr - 1).Value = "'" & WS.Cells(intRow, intIncCols(intHdr)).Value
                End If
            Next intHdr
        End If
    Next intRow
    WSNew.UsedRange.Columns.AutoFit
    
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
Sub SaveAnswersToLeft()
    Dim cell As Range
    Dim greenCol As Integer
    Dim dest As Range
    
    ' MEWC Answer Cell Green
    Const MEWC_GREEN As Long = 3631104
    
    ' Find the green cell
    For Each cell In Intersect(ActiveCell.EntireRow, ActiveSheet.UsedRange)
        If cell.Interior.Color = MEWC_GREEN Then ' mewc answer cell green
            greenCol = cell.Column
            Exit For
        End If
    Next
    
    If greenCol <> 0 Then
        Dim calcMode As Integer
        On Error Resume Next
        calcMode = Application.Calculation
        Application.Calculation = xlCalculationManual
        For Each cell In Selection
            If Cells(cell.Row, greenCol).Interior.Color = MEWC_GREEN Then
                Cells(cell.Row, greenCol).Formula = "=" & cell.Address(False, False)
                If dest Is Nothing Then
                    Set dest = Cells(cell.Row, greenCol)
                Else
                    Set dest = Union(dest, Cells(cell.Row, greenCol))
                End If
            End If
        Next
        Application.Calculation = calcMode
        
        ' if some green cells were saved to, select them and copy either those answers or the formula below.
        If Not dest Is Nothing Then
            dest.Select
            If Left(dest(1).Offset(dest.Rows.Count + 1).Formula, 1) = "=" Then
                dest(1).Offset(dest.Rows.Count + 1).Select
            Else
                dest.Select
            End If
            Selection.Copy
        End If
    End If
End Sub

'sub to allow editing and save a working copy of the active workbook
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Enable Editing And Save Copy
' Description:            Enable editing and save copy of file with suffix based on active cell (otherwise Working)
' Macro Expression:       modCaseSetup.EnableEditingAndSaveCopy([[ActiveCell]])
' Generated:              01/08/2025 05:11 PM
'----------------------------------------------------------------------------------------------------
Sub SaveCopy(Optional strSuff As String = "Working")
    Dim WB As Workbook
    Dim strPath As String

    ' Set WB to the active workbook
    Set WB = ActiveWorkbook
    If IsError(strSuff) Then
        strSuff = "Working"
    ElseIf Len(strSuff) = 0 Then
        strSuff = "Working"
    End If
    
    ' Check if the workbook is open and writable
    If WB Is Nothing Then
        MsgBox "No active workbook found!", vbExclamation
        Exit Sub
    End If

    ' Allow editing if the workbook is protected
    If WB.ProtectStructure Then
        WB.Unprotect ' Unprotect the workbook (password may be required if protected with one)
    End If

    ' Specify the path to save the copy
    strPath = WB.Path & "\" & Replace(WB.Name, ".xlsx", "_" & strSuff & ".xlsx")
    
    ' Save a copy of the workbook
    WB.SaveAs strPath
    
End Sub

