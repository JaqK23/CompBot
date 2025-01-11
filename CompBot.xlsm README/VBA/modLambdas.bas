Attribute VB_Name = "modLambdas"
Option Explicit

'function to list lambdas in this workbook
Function ListLambdas() As Variant
    Dim nm As Name
    Dim strLambdaNames() As String
    Dim lngLambdaCount As Long
    Dim result As Variant

    ' Initialize array for lambda names
    lngLambdaCount = 0

    ' Loop through all named ranges in the workbook
    For Each nm In ThisWorkbook.Names
        ' Check if the name refers to a LAMBDA function
        If InStr(1, nm.RefersTo, "=LAMBDA(", vbTextCompare) > 0 Then
            lngLambdaCount = lngLambdaCount + 1
            ReDim Preserve strLambdaNames(1 To lngLambdaCount)
            strLambdaNames(lngLambdaCount) = nm.Name
        End If
    Next nm

    ' Handle case where no lambdas are found
    If lngLambdaCount = 0 Then
        ListLambdas = "No LAMBDA functions found."
    Else
        ' Return the list of lambda names
        ListLambdas = strLambdaNames
    End If
End Function

'function to import lambdas from a specified workbook
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Import Lambdas From
' Description:            Import lambdas from a specified workbook
' Macro Expression:       modLambdas.ImportLambdasFrom()
' Generated:              01/03/2025 09:28 PM
'----------------------------------------------------------------------------------------------------
Sub ImportLambdasFrom()
    Dim sourceWorkbook As Workbook
    Dim currentWorkbook As Workbook
    Dim filePath As String
    Dim nm As Name
    Dim importedCount As Long
    Dim WS As Worksheet

    ' Initialize variables
    importedCount = 0
    Set currentWorkbook = ThisWorkbook

    ' Prompt user to select the file
    filePath = Application.GetOpenFilename("Excel Files (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , "Select a workbook to import LAMBDA functions")
    
    ' Exit if no file selected
    If filePath = "False" Then Exit Sub

    ' Open the source workbook
    Application.ScreenUpdating = False
    Set sourceWorkbook = Workbooks.Open(filePath, ReadOnly:=True)

    ' Loop through named ranges in the source workbook
    For Each nm In sourceWorkbook.Names
        ' Check if the named range is a LAMBDA function
        If InStr(1, Replace(nm.RefersTo, " ", ""), "LAMBDA(", vbTextCompare) > 0 Then
            On Error Resume Next
            ' Add the LAMBDA to the active workbook
            currentWorkbook.Names.Add Name:=nm.Name, RefersTo:=nm.RefersTo
            On Error GoTo 0
            importedCount = importedCount + 1
        End If
    Next nm

    ' Close the source workbook without saving
    sourceWorkbook.Close SaveChanges:=False
    Application.ScreenUpdating = True

    ' Notify the user
    If importedCount > 0 Then
        MsgBox importedCount & " LAMBDA functions have been imported successfully!", vbInformation
    Else
        MsgBox "No LAMBDA functions were found in the selected file.", vbExclamation
    End If
End Sub

'function to clear lambdas to only those in lambda table
'--------------------------------------------< OA Robot >--------------------------------------------
' Command Name:           Clear Lambdas
' Description:            Clear lambdas that aren't in Lambdas table - for Comp Bot maintenance
' Macro Expression:       modLambdas.ClearLambdas()
' Generated:              01/10/2025 11:00 PM
'----------------------------------------------------------------------------------------------------
Sub ClearLambdas()
    Dim WB As Workbook
    Dim LOB As ListObject
    Dim strName As String
    Dim dicName As Object
    Dim nmLambda As Name
    Dim cell As Range
    Dim booFound As Boolean

    ' Set workbook to the current workbook
    Set WB = ThisWorkbook
    
    ' Set the ListObject to the Lambdas table (replace "Lambdas" with the actual range name if needed)
    Set LOB = Range("Lambdas").ListObject
    
    ' Create a dictionary to store the names from the "Name" column of the table
    Set dicName = CreateObject("Scripting.Dictionary")
    
    ' Populate the dictionary with names from the "Name" column of the Lambdas table
    For Each cell In LOB.ListColumns("Name").DataBodyRange
        dicName(cell.Value) = True
    Next cell
    
    ' Loop through all named ranges in the workbook
    For Each nmLambda In WB.Names
        ' Check if the named range refers to a Lambda (based on the name containing "LAMBDA")
        If InStr(1, nmLambda.RefersTo, "LAMBDA(") > 0 Then
            strName = nmLambda.Name
            booFound = dicName.Exists(strName) ' Check if the name is in the dictionary
            
            ' If the name isn't found in the dictionary, delete the Lambda
            If Not booFound Then
                nmLambda.Delete
            End If
        End If
    Next nmLambda

    

End Sub
