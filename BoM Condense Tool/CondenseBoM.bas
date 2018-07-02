Attribute VB_Name = "CondenseBoM"
Option Explicit

Public Sub condense_sheet()

Dim start_row As Integer
Dim last_row As Integer
Dim act_row As Integer
Dim qty As Integer
Dim comp_row As Integer

Dim item_number As String
Dim id_number As String
Dim a As String

start_row = 1
last_row = ActiveSheet.Range("A1").SpecialCells(xlCellTypeLastCell).Row
act_row = start_row

Do While act_row <= last_row
    item_number = ActiveSheet.Cells(act_row, 1).Value
    id_number = ActiveSheet.Cells(act_row, 2).Value
    qty = ActiveSheet.Cells(act_row, 3).Value
        For comp_row = act_row + 1 To last_row
            If ActiveSheet.Cells(comp_row, 2).Value = id_number Then
                If ActiveSheet.Cells(comp_row, 1).Value <> "" And ActiveSheet.Cells(comp_row, 1).Value <> "-" Then
                    item_number = item_number & ", " & ActiveSheet.Cells(comp_row, 1).Value
                End If
                qty = qty + ActiveSheet.Cells(comp_row, 3).Value
                ActiveSheet.Rows(comp_row).EntireRow.Delete

                comp_row = comp_row - 1
            End If
        Next comp_row
    ActiveSheet.Cells(act_row, 1).Value = item_number
    ActiveSheet.Cells(act_row, 2).Value = id_number
    ActiveSheet.Cells(act_row, 3).Value = qty
    last_row = ActiveSheet.Cells(ActiveSheet.Rows.Count, "B").End(xlUp).Row
    act_row = act_row + 1
Loop
ActiveSheet.Range("A:F").EntireColumn.AutoFit
End Sub



