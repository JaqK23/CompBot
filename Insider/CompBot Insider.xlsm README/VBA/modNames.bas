Attribute VB_Name = "modNames"
Option Explicit


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Name Used Ranges
' Description:            Names used range in all sheets starting with the prefix as provided in the active cell
' Macro Expression:       modNames.NameUsedRanges([[ActiveCell]])
' Generated:              01/08/2025 01:17 PM
'----------------------------------------------------------------------------------------------------
Sub NameUsedRanges(strShtPref As String)
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        If LCase(Mid(ws.Name, 1, Len(strShtPref))) = LCase(strShtPref) Then
            ws.UsedRange.Name = "Sht_" & SanitizeRangeName(ws.Name)
        End If
    Next ws
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Name All Used Ranges
' Description:            Renames used range in all sheets
' Macro Expression:       modNames.NameAllUsedRanges()
' Generated:              01/08/2025 01:16 PM
'----------------------------------------------------------------------------------------------------
Sub NameAllUsedRanges()
    Dim wb As Workbook
    Dim ws As Worksheet
    
    Set wb = ActiveWorkbook
    For Each ws In wb.Worksheets
        On Error Resume Next
        ws.UsedRange.Name = "Sht_" & SanitizeRangeName(ws.Name)
        On Error GoTo 0
    Next ws
End Sub
