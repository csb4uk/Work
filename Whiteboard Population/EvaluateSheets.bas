Attribute VB_Name = "EvaluateSheets"
Option Explicit

Public Sub evaluate_sheet()

Dim ref_sheet As String
Dim source_sheet As String
Dim wb As String

Dim ref_sheet_last_row As Integer
Dim source_sheet_last_row As Integer
Dim row_count As Integer
Dim column_count As Integer

Dim source_cell As Variant
Dim ref_cell As Variant

    wb = ActiveWorkbook.Name
    source_sheet = ActiveSheet.Name
    source_sheet_last_row = Workbooks(wb).Sheets(source_sheet).Cells(Workbooks(wb).Sheets(source_sheet).Rows.Count, "A").End(xlUp).Row

    ref_sheet = "Evaluate Sheet"
    ref_sheet_last_row = Workbooks(wb).Sheets(ref_sheet).Cells(Workbooks(wb).Sheets(ref_sheet).Rows.Count, "A").End(xlUp).Row

    For row_count = 1 To source_sheet_last_row
        For column_count = 1 To 17
            source_cell = Workbooks(wb).Sheets(source_sheet).Cells(row_count, column_count)
            ref_cell = Workbooks(wb).Sheets(ref_sheet).Cells(row_count, column_count)
            If source_cell <> ref_cell Then
                With Workbooks(wb).Sheets(source_sheet).Cells(row_count, column_count).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 255
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            Else
                With Workbooks(wb).Sheets(source_sheet).Cells(row_count, column_count).Interior
                    .Pattern = xlSolid
                    .PatternColorIndex = xlAutomatic
                    .Color = 4506738
                    .TintAndShade = 0
                    .PatternTintAndShade = 0
                End With
            End If
        Next column_count
    Next row_count

End Sub
