Attribute VB_Name = "modInfoSheets"
Option Explicit

'==============================================================================
' modInfoSheets - rebuild CompBot's own Info sheets from the live collection.
'
' WHY THIS EXISTS: Info_Robot Commands had drifted to 42 rows while the
' collection held 84 commands, and the Lambdas List sheet was still showing a
' pre-lambda-review list including names that had since been deleted or renamed.
' Hand-maintained mirrors of generated data always drift; these two routines
' make the mirrors rebuildable in one step.
'
' HOW IT IS DRIVEN: commandtools\dump_compbot_info.py reads the collection JSON
' straight out of the saved .xlsm (a zip - no Excel, no COM) and writes a TSV.
' RefreshCommandsInfo reads that TSV back into the CompBotCommands table.
' Two steps rather than one because VBA cannot conveniently parse the custom XML
' part, and Python cannot write into a workbook Excel has open.
'==============================================================================

' --- CONSTANTS (module) ---
Private Const m_DEBUG_MODE As Boolean = False
Private Const m_CMD_TABLE  As String = "CompBotCommands"


'------------------------------------------------------------------------------
' Reload the CompBotCommands table from a TSV of
'   Name <tab> Creator <tab> Description <tab> Parameters <tab> Launch Code
' Returns a short status string rather than raising, so it can be run from the
' MCP and read as a result.
'------------------------------------------------------------------------------
Public Function RefreshCommandsInfo(ByVal strTsvPath As String) As String

    ' --- CONSTANTS (local to function) ---
    Const PROC_NAME As String = "RefreshCommandsInfo"
    Const COLS      As Long = 5

    Dim lo          As ListObject
    Dim varRows     As Variant
    Dim arrOut()    As Variant
    Dim strAll      As String
    Dim varLines    As Variant
    Dim varCells    As Variant
    Dim lngFile     As Long
    Dim lngRow      As Long
    Dim lngCol      As Long
    Dim lngCount    As Long
    Dim lngErrNum   As Long
    Dim strErrDesc  As String

    On Error GoTo ErrHandler

    If Len(Dir$(strTsvPath)) = 0 Then
        RefreshCommandsInfo = "TSV not found: " & strTsvPath
        GoTo Cleanup
    End If

    ' Read the whole file in one go. Written UTF-8 by the Python side; the names
    ' and descriptions are plain ASCII so a byte read is safe here.
    lngFile = FreeFile
    Open strTsvPath For Input As #lngFile
    strAll = Input$(LOF(lngFile), #lngFile)
    Close #lngFile

    strAll = Replace(strAll, vbCrLf, vbLf)
    varLines = Split(strAll, vbLf)

    ' Count real lines first so the output array is the right size.
    For lngRow = LBound(varLines) To UBound(varLines)
        If Len(Trim$(varLines(lngRow))) > 0 Then lngCount = lngCount + 1
    Next lngRow

    If lngCount = 0 Then
        RefreshCommandsInfo = "TSV held no rows"
        GoTo Cleanup
    End If

    ReDim arrOut(1 To lngCount, 1 To COLS)
    lngCount = 0
    For lngRow = LBound(varLines) To UBound(varLines)
        If Len(Trim$(varLines(lngRow))) > 0 Then
            lngCount = lngCount + 1
            varCells = Split(varLines(lngRow), vbTab)
            For lngCol = 1 To COLS
                If UBound(varCells) >= lngCol - 1 Then
                    arrOut(lngCount, lngCol) = varCells(lngCol - 1)
                Else
                    arrOut(lngCount, lngCol) = vbNullString
                End If
            Next lngCol
        End If
    Next lngRow

    Set lo = FindTable(m_CMD_TABLE)
    If lo Is Nothing Then
        RefreshCommandsInfo = "table not found: " & m_CMD_TABLE
        GoTo Cleanup
    End If

    ' Resize the table to the new row count, then write the body in ONE hit.
    ' Deleting and re-adding rows would break anything pointing at the table.
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    lo.Resize lo.Range.Resize(lngCount + 1, COLS)
    lo.DataBodyRange.value = arrOut

    RefreshCommandsInfo = "CompBotCommands reloaded: " & lngCount & " commands"

Cleanup:
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    lngErrNum = Err.Number
    strErrDesc = Err.Description
    On Error Resume Next
    Close #lngFile
    LogError PROC_NAME, lngErrNum, strErrDesc
    RefreshCommandsInfo = "FAILED " & lngErrNum & ": " & strErrDesc
    Resume Cleanup

End Function


'------------------------------------------------------------------------------
' Every workbook-scoped name whose RefersTo is a LAMBDA, sorted, as a 1-column
' array. This is the LIVE inventory - the thing the old Lambdas List sheet was
' showing a frozen copy of.
'------------------------------------------------------------------------------
Public Function LiveLambdaNames() As Variant

    Dim nm       As Name
    Dim colNames As Collection
    Dim arrOut() As Variant
    Dim varTmp   As Variant
    Dim lngI     As Long
    Dim lngJ     As Long
    Dim strRef   As String

    On Error GoTo ErrHandler

    Set colNames = New Collection

    For Each nm In ThisWorkbook.Names
        strRef = vbNullString
        On Error Resume Next
        strRef = nm.RefersTo
        On Error GoTo ErrHandler
        If InStr(1, strRef, "LAMBDA(", vbTextCompare) > 0 Then
            If Left$(nm.Name, 1) <> "_" Then colNames.Add nm.Name
        End If
    Next nm

    If colNames.Count = 0 Then Exit Function

    ReDim arrOut(1 To colNames.Count, 1 To 1)
    For lngI = 1 To colNames.Count
        arrOut(lngI, 1) = colNames(lngI)
    Next lngI

    ' Simple insertion sort - the list is short and this keeps it dependency-free.
    For lngI = 2 To UBound(arrOut, 1)
        varTmp = arrOut(lngI, 1)
        lngJ = lngI - 1
        Do While lngJ >= 1
            If StrComp(CStr(arrOut(lngJ, 1)), CStr(varTmp), vbTextCompare) <= 0 Then Exit Do
            arrOut(lngJ + 1, 1) = arrOut(lngJ, 1)
            lngJ = lngJ - 1
        Loop
        arrOut(lngJ + 1, 1) = varTmp
    Next lngI

    LiveLambdaNames = arrOut
    Exit Function

ErrHandler:
    If m_DEBUG_MODE Then Stop: Resume
    LiveLambdaNames = Empty

End Function


Public Function CountLiveLambdas() As Long
    Dim varNames As Variant
    varNames = LiveLambdaNames()
    If IsArray(varNames) Then CountLiveLambdas = UBound(varNames, 1)
End Function


Private Function FindTable(ByVal strName As String) As ListObject

    Dim ws As Worksheet
    Dim lo As ListObject

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, strName, vbTextCompare) = 0 Then
                Set FindTable = lo
                Exit Function
            End If
        Next lo
    Next ws

End Function

