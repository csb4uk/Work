Attribute VB_Name = "ECMarkup1Row"
Sub ec_markup_one_row()
Attribute ec_markup_one_row.VB_ProcData.VB_Invoke_Func = "S\n14"
    '==============================================================================================================================================
    'Copy the entire row of the active cell, and insert 1 row below.  Make the active row blue text, and the inserted row red with a Strikethrough
    '==============================================================================================================================================
    ActiveCell.Rows("1:1").EntireRow.Select
    Selection.Copy
    ActiveCell.Offset(1, 0).Rows("1:1").EntireRow.Select
    Selection.Insert Shift:=xlDown
    Application.CutCopyMode = False
    With Selection.Font
        .FontStyle = "Regular"
        .Size = 8
        .Strikethrough = True
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .Color = 255
        .TintAndShade = 0
    End With
    ActiveCell.Offset(-1, 0).Rows("1:1").EntireRow.Select
    With Selection.Font
        .Color = -1003520
        .TintAndShade = 0
    End With
    ActiveCell.Select
End Sub
