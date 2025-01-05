Attribute VB_Name = "modNames"
Option Explicit

'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Name Used Ranges of Sheets with suffix
' Description:            Name used ranges.
' Macro Expression:       modRangeNames.NameUsedRanges(<strShtSuff As String>)
' Generated:              11/29/2024 11:05 PM
' NOT FUNCTIONING - NEEDS PARAMETER INPUT AND I STILL NEED TO FIGURE THAT OUT
'----------------------------------------------------------------------------------------------------
Sub NameUsedRanges(strShtSuff As String)
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    For Each WS In WB.Worksheets
        If LCase(Mid(WS.Name, 1, Len(strShtSuff))) = LCase(strShtSuff) Then
            WS.UsedRange.Name = WS.Name
        End If
    Next WS
End Sub


'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Name All Used Ranges
' Description:            Name all used ranges.
' Macro Expression:       modRangeNames.NameAllUsedRanges()
' Generated:              11/29/2024 11:07 PM
'----------------------------------------------------------------------------------------------------
Sub NameAllUsedRanges()
    Dim WB As Workbook
    Dim WS As Worksheet
    
    Set WB = ActiveWorkbook
    For Each WS In WB.Worksheets
        On Error Resume Next
        WS.UsedRange.Name = WS.Name
        On Error GoTo 0
    Next WS
End Sub
