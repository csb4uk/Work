Attribute VB_Name = "Unhighlight_Row"
Sub Unhighlight()
Attribute Unhighlight.VB_ProcData.VB_Invoke_Func = "X\n14"
    '=======================================================================================================================================
    'Unhighlight all cells in the active row
    '=======================================================================================================================================
    ActiveCell.Rows("1:1").EntireRow.Select
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.Font.ColorIndex = xlAutomatic
    ActiveCell.Offset(1, 0).Range("A1").Select
End Sub

