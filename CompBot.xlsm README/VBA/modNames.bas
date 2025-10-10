Attribute VB_Name = "modNames"
Option Explicit


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Name Used Ranges
' Description:            Names used range in all sheets starting with the prefix as provided in the active cell
' Macro Expression:       modNames.NameUsedRanges([[ActiveCell]])
' Generated:              01/08/2025 01:17 PM
'----------------------------------------------------------------------------------------------------
Sub NameUsedRanges(strShtPref As String)
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    For Each WS In WB.Worksheets
        If LCase(Mid(WS.Name, 1, Len(strShtPref))) = LCase(strShtPref) Then
            WS.UsedRange.Name = "UR_" & SanitizeRangeName(WS.Name)
        End If
    Next WS
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Name All Used Ranges
' Description:            Renames used range in all sheets
' Macro Expression:       modNames.NameAllUsedRanges()
' Generated:              01/08/2025 01:16 PM
'----------------------------------------------------------------------------------------------------
Sub NameAllUsedRanges()
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    For Each WS In WB.Worksheets
        On Error Resume Next
        WS.UsedRange.Name = "UR_" & SanitizeRangeName(WS.Name)
        On Error GoTo 0
    Next WS
End Sub
