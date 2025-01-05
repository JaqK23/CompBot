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
