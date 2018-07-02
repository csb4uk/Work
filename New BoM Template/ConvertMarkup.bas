Attribute VB_Name = "ConvertMarkup"
Option Explicit

Public Sub ConvertECMarkup()

    Dim start_row As Integer
    Dim last_row As Integer
    Dim row_counter As Integer
    Dim ans As Integer

    Application.ScreenUpdating = False
    ActiveSheet.Copy Before:=ActiveSheet
    ActiveSheet.Name = "Refrigeration BOM(Final)"
    
    '===============================================================================================================================================
    'Convert the markup BoM to a 'Final' version.  If the text is blue unhighlight it and make the text black.  If the text has a strikthrough
    'delete the row
    '===============================================================================================================================================
    If ActiveSheet.Range("K5").Value <> "" Then
        ans = MsgBox("Would you like to Rev the drawing?", vbYesNo)
        If ans = vbYes Then
            ActiveSheet.Range("K5").Value = Chr(Asc(Range("K5").Value) + 1)
        End If
    End If
    With ActiveSheet
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    For row_counter = 1 To last_row
        If IsNumeric(Range("B" & row_counter).Value) = True And Range("B" & row_counter).Value <> "" Then
            start_row = row_counter
            Exit For
        End If
    Next
    For row_counter = start_row To last_row
DeleteBreak:
        If Range("A" & row_counter).Rows("1:1").EntireRow.Font.Strikethrough = True Then
            Range("A" & row_counter).Rows("1:1").EntireRow.Delete Shift:=xlUp
            GoTo DeleteBreak
        End If
        If Range("A" & row_counter).Rows("1:1").EntireRow.Font.Color = 15773696 Then
           Range("A" & row_counter).Rows("1:1").EntireRow.Font.ColorIndex = xlAutomatic
        End If
    Next
    Application.ScreenUpdating = True
End Sub
