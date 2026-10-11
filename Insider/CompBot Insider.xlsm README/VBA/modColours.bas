Attribute VB_Name = "modColours"
Option Explicit

'from Hadyn Wiseman
Sub ColourFill_FAST()
    
    Dim c As Range
    ' Store the current calculation state to restore it later
    Dim originalCalcState As Long
    originalCalcState = Application.Calculation

    '--- 1. Optimization Flags ---
    ' Stop Excel from redrawing the screen
    Application.ScreenUpdating = False
    ' Stop automatic calculation
    Application.Calculation = xlCalculationManual

    On Error GoTo Cleanup ' Ensure cleanup runs even if an error occurs

    '--- 2. Core Logic ---
    For Each c In Selection.Cells
        ' The core task: set color based on cell's value
        ' It's best practice to ensure the value is a number before assignment
        If IsNumeric(c.value) Then
            c.Interior.Color = c.value
        End If
    Next c

Cleanup:
    '--- 3. Clean-up (ALWAYS re-enable application features) ---
    Application.ScreenUpdating = True
    ' Restore the original calculation state
    Application.Calculation = originalCalcState

End Sub

'from Hadyn Wiseman
Sub OverlayColorCode()
    
    Dim c As Range
    
    '--- 1. Optimization Flags (Maximizes speed) ---
    ' Store original calculation state to restore it later
    Dim originalCalcState As Long
    originalCalcState = Application.Calculation
    
    ' Turn off screen updating and set calculation to manual
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Cleanup ' Essential for error handling and cleanup

    '--- 2. Core Logic ---
    For Each c In Selection.Cells
        ' Check if the cell has an interior color applied
        If c.Interior.ColorIndex <> xlNone Then
            ' Set the cell's value to its Interior Color value (Long integer)
            ' This integer is the BGR value (Blue*256^2 + Green*256 + Red)
            c.value = c.Interior.Color
        Else
            ' If no interior color is set, you might want to replace it
            ' with a specific value (e.g., 0 for black or -1 for 'No Fill').
            ' Using -1 (xlNone) might be confusing, so 0 (Black) is often safer
            ' or you can leave the value as is.
            ' We'll use 0 to represent 'No Fill' for simplicity in formulas.
            c.value = 0
        End If
    Next c

Cleanup:
    '--- 3. Clean-up (ALWAYS re-enable application features) ---
    Application.ScreenUpdating = True
    ' Restore the original calculation state
    Application.Calculation = originalCalcState
    
    ' Optional: Provide user feedback if an error occurred
    If Err.Number <> 0 Then
        MsgBox "An error occurred: " & Err.Description, vbCritical
    End If

End Sub
