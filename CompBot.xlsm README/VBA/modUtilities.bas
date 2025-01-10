Attribute VB_Name = "modUtilities"
Option Explicit
Option Base 1
Private varCalc As Variant

Sub VBAInit()
    Application.ScreenUpdating = False
    varCalc = Application.Calculation
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Sub

Sub VBAFin()
    Application.ScreenUpdating = True
    Application.Calculation = varCalc
    Application.DisplayAlerts = True
    Application.StatusBar = False
End Sub

'function to check if a sheet exists in the active workbook
Function SheetExists(strSht As String) As Boolean
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    SheetExists = False
    For Each WS In WB.Worksheets
        If WS.Name = strSht Then
            SheetExists = True
            Exit For
        End If
    Next WS
End Function

' --------------------------------------------< OA Robot >--------------------------------------------
'  Name:                SanitizeRangeName
'  Description:         Function to sanitize names for named ranges
'  Credit:              Erik Oehm
'  Source:              https://github.com/ExcelRobot/MEWC-Robot/blob/main/MEWC%20Robot.xlsm
' ----------------------------------------------------------------------------------------------------
Function SanitizeRangeName(proposedName As String) As String
    Dim result As String
    Dim i As Long
    Dim char As String
    Dim validFirstChars As String
    Dim validChars As String
    
    ' If empty string, return default name
    If Len(proposedName) = 0 Then
        SanitizeRangeName = "Range1"
        Exit Function
    End If
    
    ' Initialize working string
    result = proposedName
    
    ' Replace spaces with underscores
    result = Replace(result, " ", "_")
    
    ' Define valid characters
    validFirstChars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ_\"
    validChars = validFirstChars & "0123456789."
    
    ' Check and fix first character
    char = UCase(Left(result, 1))
    If InStr(validFirstChars, char) = 0 Then
        ' If first char is a number, prepend "N_"
        If IsNumeric(char) Then
            result = "N_" & result
        Else
            ' Replace invalid first char with underscore
            result = "_" & Mid(result, 2)
        End If
    End If
    
    ' Clean remaining characters
    Dim sanitized As String
    sanitized = Left(result, 1)
    For i = 2 To Len(result)
        char = Mid(result, i, 1)
        ' Keep only valid characters
        If InStr(validChars, UCase(char)) > 0 Then
            sanitized = sanitized & char
        Else
            ' Replace invalid chars with underscore
            sanitized = sanitized & "_"
        End If
    Next i
    
    ' Trim to 255 characters if needed
    If Len(sanitized) > 255 Then
        sanitized = Left(sanitized, 255)
    End If
    
    ' Handle reserved words and cell references
    result = sanitized
    If Not IsValidRangeName(result) Then
        ' Add prefix for reserved words or cell references
        result = "RNG_" & result
    End If
    
    ' Verify final result
    If Not IsValidRangeName(result) Then
        ' If still invalid, use a safe default
        result = "Range_" & Format(Now, "yyyymmddhhnnss")
    End If
    
    SanitizeRangeName = result
End Function

' --------------------------------------------< OA Robot >--------------------------------------------
'  Name:                MarkAsInputCells
'  Description:         Mark as input cells
'  Credit:              Erik Oehm
'  Source:              Lambda robot
'  Notes:               JK 2025-01-08: Adjusted to not output to robot logs and set colors
' ----------------------------------------------------------------------------------------------------
Public Sub MarkAsInputCells(ByVal GivenRange As Range, Optional ByVal InteriorOnly As Boolean = True)

    Dim INPUT_CELL_BACKGROUND_COLOR As Long
    Dim INPUT_CELL_FONT_COLOR As Long
    
    INPUT_CELL_BACKGROUND_COLOR = 13434879
    INPUT_CELL_FONT_COLOR = 16711680
    
    GivenRange.Interior.Color = INPUT_CELL_BACKGROUND_COLOR
    If Not InteriorOnly Then
        GivenRange.Font.Color = INPUT_CELL_FONT_COLOR
    End If
    
End Sub

'Clears named ranges with RefersTo="=#NAME?"
Sub ClearNamedRangeErrors()
    Dim WB As Workbook
    Dim nmName As Name
    Dim intDel As Integer
    
    Set WB = ActiveWorkbook
    intDel = 0
    
    For Each nmName In WB.Names
        If nmName.RefersTo = "=#NAME?" Then
            intDel = intDel + 1
            Debug.Print "Deletion number " & CStr(intDel) & ": " & nmName.Name
            On Error Resume Next
            nmName.Delete
            If Err.Number <> 0 Then Debug.Print "Deletion failed"
            On Error GoTo 0
        End If
    Next nmName
    
End Sub
