Attribute VB_Name = "AshraeTAble"
Option Explicit
Public Sub create_table()
    
    Dim start_row As Integer
    Dim last_row As Integer
    Dim transpose_row As Integer
    Dim current_row As Integer

    Dim first_char As Variant
    Dim second_char As Variant
    Dim msg As String

    'Initialize Variables
    start_row = Selection.Rows(1).Row
    current_row = start_row
    With ActiveSheet.UsedRange
        last_row = .Rows(.Rows.Count).Row
    End With
    transpose_row = id_transpose_row(current_row, last_row)
    Do While current_row < last_row
NextStation:
        transpose_row = id_transpose_row(current_row, last_row)
        first_char = Mid(Range("A" & current_row).Value, 1, 1)
        If IsNumeric(first_char) = False And first_char <> "" Then
            second_char = Mid(Range("A" & current_row).Value, 2, 1)
            If second_char = UCase(second_char) And second_char <> "" Then
                sub_transpose_rows current_row, transpose_row
                hyperlink_station transpose_row
                GoTo NextStation
            ElseIf second_char = "" Then
                Range("A" & current_row).Delete Shift:=xlUp
                GoTo NextStation
            Else
                If current_row <> start_row Then
                    next_city current_row, transpose_row
                End If
            End If
        End If
        current_row = current_row + 1
        With ActiveSheet.UsedRange
            last_row = .Rows(.Rows.Count).Row
        End With
    Loop
End Sub

Function id_transpose_row(ByVal current_row As Integer, ByVal last_row As Integer)
    Dim counter As Integer
    counter = current_row - 1
    Do While IsEmpty(Range("B" & counter).Value) = False
        counter = counter + 1
    Loop
    id_transpose_row = counter
End Function

Private Sub sub_transpose_rows(ByVal current_row As Integer, ByVal transpose_row As Integer)
    Range("A" & current_row & ":A" & current_row + 4).Copy
    Range("B" & transpose_row).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
        False, Transpose:=True
    Range("A" & current_row & ":A" & current_row + 4).Delete Shift:=xlUp
End Sub

Private Sub hyperlink_station(ByVal transpose_row As Integer)

    Dim folder_path As String
    Dim file_name As String
    Dim full_file As String
    Dim weather_station As String

    weather_station = Range("C" & transpose_row).Value
    folder_path = "I:\engineering\Refrigeration\NEW R DRIVE\ASHRAE Weather Data\STATIONS\"
    file_name = weather_station & "_p.pdf"
    full_file = Dir(folder_path & file_name)

    If Len(full_file) > 0 Then
        With Worksheets("Weather Station (US)")
            .Hyperlinks.Add Anchor:=.Range("C" & transpose_row), _
            Address:=folder_path & file_name, _
            ScreenTip:="Weather Station Data", _
            TextToDisplay:=weather_station
        End With
    End If
End Sub

Private Sub next_city(ByRef current_row As Integer, ByVal transpose_row As Integer)
    Range("A" & current_row & ":A" & transpose_row - 1).Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Row = transpose_row
End Sub

