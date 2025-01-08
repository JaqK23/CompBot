Attribute VB_Name = "modGoToSpecial"
Option Explicit

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
