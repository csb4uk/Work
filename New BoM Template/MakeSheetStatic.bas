Attribute VB_Name = "MakeSheetStatic"
Option Explicit

Public Sub MakeStaticSheet()

    Dim start_row As Integer
    Dim last_row As Integer
    Dim row_counter As Integer

    Dim swb As String
    Dim sws As String
    
    Dim wb_connection As Object
    
    Application.ScreenUpdating = False
    
    '===============================================================================================================================================
    'Copy the current sheet into a new sheet, rename it to (Static), copy all the cells in the sheet and paste them as values
    '===============================================================================================================================================
    ActiveSheet.Copy Before:=Sheets(1)
    Sheets(1).Name = "Refrigeration BOM(Static)"
    swb = ActiveWorkbook.Name
    sws = "Refrigeration BOM(Static)"
    With ActiveSheet
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    For row_counter = 1 To last_row
        If IsNumeric(Range("B" & row_counter).Value) = True And Range("B" & row_counter).Value <> "" Then
            start_row = row_counter
            Exit For
        End If
    Next

    '===============================================================================================================================================
    'Identify parts that have subcomponents.  Store the primary component and its subcomponents in the BoM list excel sheet
    'This is really only a possible future idea where we can store subcomponents that are typically used.  Such as clamps for filter-driers.
    '===============================================================================================================================================
    id_dependencies start_row, last_row, swb, sws

    '===============================================================================================================================================
    'Copy and paste the values
    '===============================================================================================================================================
    Workbooks(swb).Sheets(sws).Range("A" & start_row & ":N" & last_row).Copy
    Workbooks(swb).Sheets(sws).Range("A" & start_row & ":N" & last_row).PasteSpecial _
        Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    '===============================================================================================================================================
    'Remove any highlighting that is leftt in the document
    '===============================================================================================================================================
    With Rows(start_row & ":" & last_row).EntireRow.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    '===============================================================================================================================================
    'Disable any Refrigeration Data Access Database Connections
    '===============================================================================================================================================
    For Each wb_connection In Workbooks(swb).Connections
        If "Refrigeration Data" = wb_connection.Name Or "Refrigeration Data1" = wb_connection.Name Then
            ActiveWorkbook.Connections(wb_connection.Name).Delete
        End If
    Next
    Application.ScreenUpdating = True
End Sub

Private Sub id_dependencies(start_row, last_row, swb, sws)

    Dim current_row As Integer
    Dim id_last_no_1 As Integer
    Dim id_last_no_2 As Integer
    Dim dependent_row As Integer
    

    Dim file_path As String
    Dim file_name As String
    Dim BoM_wb As String
    Dim BoM_ws_dependent As String
    Dim BoM_ws_independent As String
    Dim primary_component As String
    Dim dependent_component As String
    Dim item_number As String
    Dim next_item As String
    Dim no_1 As String
    Dim no_2 As String

    current_row = start_row
    file_path = "R:\NEW R DRIVE\Source Codes\New BoM Template\"
    file_name = "BoM list.xlsm"
    Workbooks.Open _
        Filename:=file_path & file_name, _
        ReadOnly:=False
            BoM_wb = ActiveWorkbook.Name
            BoM_ws_dependent = "Dependent"
            BoM_ws_independent = "Independent"
    Workbooks(swb).Activate
    Do While current_row <= last_row
        dependent_row = current_row + 1
        item_number = Workbooks(swb).Sheets(sws).Range("A" & current_row).Value
            id_last_no_1 = id_last_number(item_number)
            no_1 = Mid(item_number, 1, id_last_no_1)
            primary_component = Workbooks(swb).Sheets(sws).Range("B" & current_row).Value
NextDependent:
        next_item = Workbooks(swb).Sheets(sws).Range("A" & dependent_row).Value
            id_last_no_2 = id_last_number(next_item)
            no_2 = Mid(next_item, 1, id_last_no_2)
            dependent_component = Workbooks(swb).Sheets(sws).Range("B" & dependent_row).Value
        If item_number <> "" And next_item <> "" Then
            If no_1 = no_2 Then                                     'Item is dependent
                write_id_relation primary_component, dependent_component, BoM_wb, BoM_ws_dependent
                dependent_row = dependent_row + 1
                GoTo NextDependent
            End If
        ElseIf primary_component <> "" And dependent_component <> "" Then
            write_id_relation primary_component, dependent_component, BoM_wb, BoM_ws_independent
        End If
        current_row = dependent_row
    Loop
    Workbooks(BoM_wb).Close SaveChanges:=True
End Sub

Private Function id_last_number(this_item)

    Dim start_char As String
    Dim str_counter As Integer

    Select Case Mid(this_item, 1, 1)
        Case "H"
            start_char = 2
        Case Else
            start_char = 1
    End Select

    For str_counter = start_char To Len(this_item)
        Select Case Asc(Mid(this_item, str_counter, 1))
            Case 48 To 57
                id_last_number = str_counter
            Case Else
                Exit For
        End Select
    Next str_counter
End Function

Private Sub write_id_relation(primary_component, component, BoM_wb, BoM_ws)

    Dim last_row As Integer
    Dim row_count As Integer
    Dim last_col As Integer
    Dim col_count As Integer

    Dim id_match As Boolean
    Dim col_match As Boolean


    id_match = False
    With Workbooks(BoM_wb).Sheets(BoM_ws)
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    For row_count = 1 To last_row
        If Workbooks(BoM_wb).Sheets(BoM_ws).Range("A" & row_count).Value = primary_component Then
            With Workbooks(BoM_wb).Sheets(BoM_ws)
                last_col = .Cells(row_count, .Columns.Count).End(xlToLeft).Column
            End With
            For col_count = 1 To last_col
                If Workbooks(BoM_wb).Sheets(BoM_ws).Cells(row_count, col_count).Value = component Then
                    col_match = True
                    Exit For
                Else
                    col_match = False
                End If
            Next
            If col_match = False Then
                Workbooks(BoM_wb).Sheets(BoM_ws).Cells(row_count, last_col + 1).NumberFormat = "@"
                Workbooks(BoM_wb).Sheets(BoM_ws).Cells(row_count, last_col + 1).Value = component
            End If
            id_match = True
        End If
    Next
    If id_match = False Then
        Workbooks(BoM_wb).Sheets(BoM_ws).Range("A" & last_row + 1).NumberFormat = "@"
        Workbooks(BoM_wb).Sheets(BoM_ws).Range("A" & last_row + 1).Value = primary_component
        Workbooks(BoM_wb).Sheets(BoM_ws).Range("B" & last_row + 1).NumberFormat = "@"
        Workbooks(BoM_wb).Sheets(BoM_ws).Range("B" & last_row + 1).Value = component
    End If
End Sub


