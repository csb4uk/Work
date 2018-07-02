Attribute VB_Name = "CrossCheckBoM"
Option Explicit

Public Sub cross_check_BoM()
Dim source_wb_name As String
Dim source_ws_name As String
Dim ref_book_BoM As String
Dim ref_wb_name As String
Dim item_number As String
Dim id_number As String
Dim qty As String

Dim start_row_sb As Integer
Dim last_row_sb As Integer
Dim start_row_rb As Integer
Dim last_row_rb As Integer
Dim total_rows_sb As Integer
Dim total_rows_rb As Integer
Dim row_count As Integer

Dim i As Integer
Dim j As Integer
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim w As Integer
Dim x As Integer
Dim y As Integer
Dim z As Integer
Dim perc As Double
Dim total_items As Integer
Dim ans As Variant
Dim arr1 As Variant
Dim arr2 As Variant
Dim arr3() As Variant
Dim arr4() As Variant
Dim arr5() As Variant
Dim arr6() As Variant
Dim arr7() As Variant
ReDim Preserve arr3(0 To 2, 0)
ReDim Preserve arr4(0 To 2, 0)
ReDim Preserve arr5(0 To 2, 0)
ReDim Preserve arr6(0 To 2, 0)
ReDim Preserve arr7(0 To 2, 0)

Dim wkb As Workbook

Dim wks As Worksheet

Application.DisplayAlerts = False

'Store everything in the excel BOM into an array
start_row_sb = Selection.Rows(1).Row
last_row_sb = Selection.Rows.Count + start_row_sb - 1
total_rows_sb = last_row_sb - start_row_sb + 1
source_wb_name = ActiveWorkbook.Name
source_ws_name = ActiveSheet.Name
Call convert2text(start_row_sb, last_row_sb)
arr1 = Workbooks(source_wb_name).Sheets(source_ws_name).Range("A" & start_row_sb & ":F" & last_row_sb).Value

'If there is already a comparison sheet delete ir
For Each wks In Worksheets
    If wks.Name Like "Comparison" Then
        Sheets("Comparison").Delete
        Exit For
    End If
Next wks

'Find the BoM you wish to reference and extract the
TryAgain:
ref_book_BoM = InputBox("Please type in the BoM you wish to reference")
For Each wkb In Workbooks
    If wkb.Name Like ref_book_BoM & "*" Then
        ref_wb_name = wkb.Name
        wkb.Activate
        Exit For
    End If
Next wkb

If ActiveWorkbook.Name = source_wb_name Then
    ans = MsgBox("BoM of that name was not found.  Would you like to try again?", vbYesNo)
    If ans = vbYes Then
        GoTo TryAgain
    Else
        End
    End If
End If

start_row_rb = Selection.Rows(1).Row
last_row_rb = Selection.Rows.Count + start_row_sb - 1
Call convert2text(start_row_rb, last_row_rb)
total_rows_rb = last_row_rb - start_row_rb + 1
arr2 = Workbooks(ref_wb_name).Sheets(ActiveSheet.Name).Range("A" & start_row_rb & ":F" & last_row_rb).Value

total_items = 0
w = 0
x = 0
y = 0
z = 0

Workbooks(source_wb_name).Activate

For i = 1 To total_rows_sb
    If Not IsEmpty(arr1(i, 2)) Then
        If Not IsEmpty(arr1(i, 1)) Then
            item_number = arr1(i, 1)
            id_number = arr1(i, 2)
            qty = arr1(i, 3)
            For j = 1 To total_rows_rb
                If Not IsEmpty(arr2(j, 2)) Then
                    If item_number = CStr(arr2(j, 1)) Then
                        If id_number = CStr(arr2(j, 2)) Then
                            If qty = CStr(arr2(j, 3)) Then
                                Call eval_arr(arr5, y, item_number, id_number, qty)     'Items that match
                            Else
                                Call eval_arr(arr6, z, item_number, id_number, qty)     'Quantity of item varies from reference drawing
                            End If
                        Else
                            Call eval_arr(arr4, x, item_number, id_number, qty)         'Item ID does not match reference drawing
                        End If
                        Exit For
                    End If
                End If
            Next j
            If j > total_rows_rb Then
                Call eval_arr(arr3, w, item_number, id_number, qty)                     'Items that are not on the reference drawing
            End If
        Else
            item_number = arr1(i, 1)
            id_number = arr1(i, 2)
            qty = arr1(i, 3)
            For j = 1 To total_rows_rb
                If Not IsEmpty(arr2(j, 2)) Then
                    If id_number = CStr(arr2(j, 2)) Then
                        If qty = CStr(arr2(j, 3)) Then
                            Call eval_arr(arr5, y, item_number, id_number, qty)         'Items that match
                        Else
                            Call eval_arr(arr6, z, item_number, id_number, qty)         'Quantity of item varies from reference drawing
                        End If
                    Exit For
                    End If
                End If
            Next j
            If j > total_rows_rb Then
                Call eval_arr(arr3, w, item_number, id_number, qty)                     'Items that are not on the reference drawing
            End If
        End If
    total_items = total_items + 1
    End If
Next i

c = 0
For a = 1 To total_rows_rb
    If Not IsEmpty(arr2(a, 2)) Then
        item_number = arr2(a, 1)
        id_number = arr2(a, 2)
        qty = arr2(a, 3)
        For b = 1 To total_rows_sb
            If item_number = CStr(arr1(b, 1)) Then
                Exit For
            End If
        Next b
        If b > total_rows_sb Then
            Call eval_arr(arr7, c, item_number, id_number, qty)                         'Items that are not on the new BoM
            total_items = total_items + 1
        End If
    End If
Next a

Sheets.Add
ActiveSheet.Name = "Comparison"
Sheets("Comparison").Select
Range("A:E").NumberFormat = "@"
row_count = Sheets("Comparison").Cells(Rows.Count, 2).End(xlUp).Row 'Finds the number of rows used in the doc.

If Not IsEmpty(arr3) Then
    Sheets("Comparison").Range("A" & row_count) = "Items that are not on the reference drawing"
    Sheets("Comparison").Range("A" & row_count).HorizontalAlignment = xlLeft
    Call paste_arr(row_count, arr3)
    row_count = row_count + 2
End If
If Not IsEmpty(arr4) Then
    Sheets("Comparison").Range("A" & row_count) = "Item ID does not match reference drawing"
    Sheets("Comparison").Range("A" & row_count).HorizontalAlignment = xlLeft
    Call paste_arr(row_count, arr4)
    row_count = row_count + 2
End If
If Not IsEmpty(arr6) Then
    Sheets("Comparison").Range("A" & row_count) = "Quantity of item varies from reference drawing"
    Sheets("Comparison").Range("A" & row_count).HorizontalAlignment = xlLeft
    Call paste_arr(row_count, arr6)
    row_count = row_count + 2
End If
If Not IsEmpty(arr7) Then
    Sheets("Comparison").Range("A" & row_count) = "Items that are not on the new BoM"
    Sheets("Comparison").Range("A" & row_count).HorizontalAlignment = xlLeft
    Call paste_arr(row_count, arr7)
    row_count = row_count + 2
End If
If Not IsEmpty(arr5) Then
    Sheets("Comparison").Range("A" & row_count) = "Items that match"
    Sheets("Comparison").Range("A" & row_count).HorizontalAlignment = xlLeft
    Call paste_arr(row_count, arr5)
    row_count = row_count + 2
End If

perc = Format((UBound(arr5, 2) + 1) / (total_items) * 100, "0.00")
MsgBox ((UBound(arr5, 2) + 1) & "\" & total_items & " are a match, or " & perc & "%")

Application.DisplayAlerts = True

End Sub


Private Sub eval_arr(ByRef arr As Variant, ByRef counter As Integer, ByVal item_number As String, ByVal id_number As String, ByVal qty As Integer)
    ReDim Preserve arr(2, counter)
    arr(0, counter) = item_number       'Stores the item number in the first index of arr
    arr(1, counter) = id_number     'Stores the id number in the second index of arr
    arr(2, counter) = qty               'Stores the quantity in the third index of arr
    counter = counter + 1
End Sub
Private Sub convert2text(ByVal start_row As Integer, ByVal last_row As Integer)
Dim row_count As Integer
Dim row_str As String
    Range("B" & start_row & ":B" & last_row).NumberFormat = "@"
    row_count = start_row
        Do While row_count <= last_row
            If IsNumeric(Left(Range("B" & row_count), 1)) = True Then
                Do While Len(Range("B" & row_count)) < 5
                    Range("B" & row_count).Value = "0" & Range("B" & row_count).Value
                Loop
            End If
            row_count = row_count + 1
        Loop
End Sub
Private Sub paste_arr(ByRef row_count As Integer, ByRef arr As Variant)
Dim count_1 As Integer

For count_1 = 0 To UBound(arr, 2)
    row_count = row_count + 1
    Sheets("Comparison").Range("A" & row_count) = arr(0, count_1)
    Sheets("Comparison").Range("A" & row_count).HorizontalAlignment = xlCenter
    Sheets("Comparison").Range("B" & row_count).Value = arr(1, count_1)
    Sheets("Comparison").Range("B" & row_count).HorizontalAlignment = xlCenter
    Sheets("Comparison").Range("C" & row_count) = arr(2, count_1)
    Sheets("Comparison").Range("C" & row_count).HorizontalAlignment = xlCenter
Next count_1

End Sub

