Attribute VB_Name = "DataNorm"
Option Explicit

Public Sub normalize_data()
    
    Dim wks As Worksheet
    Dim msg As String
    Dim ans As Integer
    Dim chamber_count As Integer

    '=====================================================================================================================
    'Loop through all worksheets and evaluate if you want to add the data to the transition sheet
    '=====================================================================================================================
    chamber_count = 0
    For Each wks In Worksheets
        msg = "Would you like to analyze transition data for " & wks.Name & "?"
        ans = MsgBox(msg, vbYesNo)
        If ans = vbYes Then
            create_initial_plot (wks.Name)
            chamber_count = chamber_count + 1
            evaluate_transitions wks.Name, chamber_count
        End If
    Next
    autosize_columns
    create_plots
End Sub
Private Sub create_initial_plot(ByVal ws_name)
    Dim date_col As Integer
    Dim time_col As Integer
    Dim ct_col As Integer
    Dim sp_col As Integer
    Dim last_col As Integer
    Dim last_row As Integer
    Dim series_counter As Integer

    id_columns last_col, date_col, time_col, ct_col, sp_col, last_row, ws_name

    series_counter = 0
    create_ct_plot series_counter, time_col, last_row, ct_col, date_col, ws_name

    create_dict_plots ws_name, sp_col, time_col, series_counter, last_row

    

End Sub

Private Sub id_columns(ByRef last_col, ByRef date_col, ByRef time_col, ByRef ct_col, ByRef sp_col, ByRef last_row, ByVal ws_name)
    Dim col_count As Integer
    With Sheets(ws_name)
        last_col = .Cells(1, .Columns.Count).End(xlToLeft).Column
        For col_count = 1 To last_col
            If UCase(.Cells(1, col_count).Value) = "DATE" Then
                date_col = col_count
            ElseIf UCase(.Cells(1, col_count).Value) = "TIME" Then
                time_col = col_count
            ElseIf InStr(1, UCase(.Cells(1, col_count).Value), "TEMPERATURE PV") > 0 Then
                ct_col = col_count
            ElseIf InStr(1, UCase(.Cells(1, col_count).Value), "TEMPERATURE SP") > 0 Then
                sp_col = col_count
            End If
        Next
        last_row = .Cells(.Rows.Count, date_col).End(xlUp).Row
    End With
End Sub

Private Sub create_ct_plot(ByRef series_counter, ByVal time_col, ByVal last_row, ByVal ct_col, ByVal date_col, ByVal ws_name)
    Dim series_coll As Series

    With Sheets(ws_name)
        .Activate
        .Shapes.AddChart.Select
        ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
        For Each series_coll In ActiveChart.SeriesCollection
            series_coll.Delete
        Next
        
        ActiveChart.Axes(xlCategory, xlPrimary).MinimumScale = .Cells(2, time_col)
        ActiveChart.Axes(xlCategory, xlPrimary).MaximumScale = .Cells(last_row, time_col)
        
        ActiveChart.SeriesCollection.NewSeries
        series_counter = series_counter + 1
        ActiveChart.SeriesCollection(series_counter).Name = .Cells(1, ct_col)
        ActiveChart.SeriesCollection(series_counter).XValues = .Range(.Cells(2, time_col), .Cells(last_row, time_col))
        ActiveChart.SeriesCollection(series_counter).Values = .Range(.Cells(2, ct_col), .Cells(last_row, ct_col))
    End With
End Sub
Private Sub create_dict_plots(ByVal ws_name, ByVal sp_col, ByVal time_col, ByVal series_counter, ByVal last_row)

    Dim start_row As Integer
    Dim row_counter As Integer
    Dim name_series As String
    Dim input_box_msg As String
    Dim key_counter As Integer
    Dim input_box_remove As Variant
    Dim dict_key As Variant

    Dim sp_dict As Object
    Set sp_dict = CreateObject("Scripting.Dictionary")

    row_counter = 3
    start_row = row_counter

    With Sheets(ws_name)
        Do While row_counter < last_row
            If .Cells(row_counter - 1, sp_col) <> .Cells(row_counter, sp_col) Then
                series_counter = series_counter + 1
                ActiveChart.SeriesCollection.NewSeries
                name_series = "Series " & series_counter - 1 & ", SP: " & .Cells(row_counter - 1, sp_col)
                ActiveChart.SeriesCollection(series_counter).Name = name_series
                sp_dict.Add name_series, start_row
                If series_counter = 1 Then
                    ActiveChart.SeriesCollection(series_counter).XValues = .Range(.Cells(start_row, time_col), .Cells(row_counter - 1, time_col))
                    ActiveChart.SeriesCollection(series_counter).Values = .Range(.Cells(start_row, sp_col), .Cells(row_counter - 1, sp_col))
                Else
                    ActiveChart.SeriesCollection(series_counter).XValues = .Range(.Cells(start_row - 1, time_col), .Cells(row_counter - 1, time_col))
                    ActiveChart.SeriesCollection(series_counter).Values = .Range(.Cells(start_row - 1, sp_col), .Cells(row_counter - 1, sp_col))
                End If
                start_row = row_counter
            End If
            row_counter = row_counter + 1
        Loop
        series_counter = series_counter + 1
        ActiveChart.SeriesCollection.NewSeries
        name_series = "Series " & series_counter - 1 & ", SP: " & .Cells(row_counter - 1, sp_col)
        ActiveChart.SeriesCollection(series_counter).Name = name_series
        sp_dict.Add name_series, start_row
        ActiveChart.SeriesCollection(series_counter).XValues = .Range(.Cells(start_row - 1, time_col), .Cells(row_counter - 1, time_col))
        ActiveChart.SeriesCollection(series_counter).Values = .Range(.Cells(start_row - 1, sp_col), .Cells(row_counter - 1, sp_col))
    End With
    With ActiveChart
        .SetElement (msoElementChartTitleAboveChart)
        .ChartTitle.Text = ws_name
        .Location Where:=xlLocationAsNewSheet, Name:=ws_name & "_Chart"
    End With


    Do While input_box_remove <> "N" Or input_box_remove <> ""
        
        input_box_msg = "Would you like to remove any of the following Setpoints?" & vbCrLf
        key_counter = 0
        For Each dict_key In sp_dict.Keys
            key_counter = key_counter + 1
            input_box_msg = input_box_msg & vbTab & key_counter & ". " & dict_key & vbCrLf
        Next
        input_box_remove = InputBox(input_box_msg)
        If input_box_remove <> "N" And input_box_remove <> "" Then
            remove_sp sp_dict, last_row, sp_col, ws_name, input_box_remove, input_box_msg
        Else
            Exit Do
        End If
    Loop
    Application.DisplayAlerts = False
    ActiveSheet.Delete
    Application.DisplayAlerts = True
End Sub
Private Sub remove_sp(ByRef sp_dict, ByVal last_row, ByVal sp_col, ByVal ws_name, ByVal input_box_remove, ByVal input_box_msg)

    
    Dim dict_key As Variant
    Dim input_box_msg_1 As String
    Dim dict_row As Integer
    Dim key_counter_1 As Integer
    Dim input_box_replace As Variant
    Dim setpoint_val As Double
    Dim replace_val_row As Integer
    Dim replace_val As Double
    Dim current_row As Integer
    Dim remove_key As Variant
    Dim new_key As Variant
    Dim series_num As Integer
    Dim new_series_name As String

    
    input_box_msg_1 = "You selected to remove Series " & input_box_remove & "." & vbCrLf
    input_box_msg_1 = input_box_msg_1 & "Which Setpoint value would you like to overwrite it with?" & vbCrLf
    input_box_replace = InputBox(input_box_msg_1 & input_box_msg)

    key_counter_1 = 0
    For Each dict_key In sp_dict.Keys
        key_counter_1 = key_counter_1 + 1
        If key_counter_1 = input_box_remove Then
            series_num = key_counter_1
            dict_row = sp_dict(dict_key)
            remove_key = dict_key
        ElseIf key_counter_1 = input_box_replace Then
            replace_val_row = sp_dict(dict_key)
        End If
    Next

    With Sheets(ws_name)
        replace_val = .Cells(replace_val_row, sp_col)
        current_row = dict_row
        setpoint_val = .Cells(current_row, sp_col)
        Do While .Cells(current_row, sp_col) = setpoint_val
            .Cells(current_row, sp_col) = replace_val
            current_row = current_row + 1
        Loop
    End With
    new_series_name = "Series " & series_num & ", SP: " & replace_val
    
    sp_dict.Remove (remove_key)
    Set sp_dict = DictAdd(sp_dict, new_series_name, dict_row, series_num - 1)
    
    ActiveChart.SeriesCollection(series_num + 1).Name = new_series_name
    Application.Wait Now + TimeSerial(0, 0, 1)    'Wait 2 seconds
    ActiveChart.Refresh
End Sub
Private Function DictAdd(start_dict, key_add, item_add, after_key) As Object

    Dim key_dict As Variant
    Dim key_count As Integer

    key_count = 0
    Set DictAdd = CreateObject("Scripting.Dictionary")
    For Each key_dict In start_dict
        key_count = key_count + 1
        DictAdd.Add key_dict, start_dict(key_dict)
        If CInt(key_count) = CInt(after_key) Then
            DictAdd.Add key_add, item_add
        End If
    Next
End Function
Private Sub evaluate_transitions(ByVal chamber_sn, ByVal chamber_count)

    Dim data_set As Variant
    Dim transition_dict As Object
    Dim row_counter As Integer
    Dim transition_counter As Integer
    Dim input_col As Integer
    Dim time_col As Integer
    Dim sp_col As Integer
    Dim ct_col As Integer
    Dim sp_start As Integer
    Dim row_start As Integer
    Dim sp_end As Integer
    Dim row_end As Integer
    Dim last_eval_row As Integer
    Dim ws_name As String
    '=====================================================================================================================
    'Gather the data in the current sheet into an array
    '=====================================================================================================================
    data_set = collect_data_sets(chamber_sn)
    '=====================================================================================================================
    'Identify the transtions made
    '=====================================================================================================================
    Set transition_dict = collect_transitions(data_set)

    For transition_counter = 1 To transition_dict.Count - 1
        row_counter = 2
        sp_start = transition_dict("Transition " & transition_counter)("SP")
        sp_end = transition_dict("Transition " & transition_counter + 1)("SP")
        ws_name = "Transition " & sp_start & " To " & sp_end
        row_start = transition_dict("Transition " & transition_counter + 1)("Row")

        If transition_dict.Exists("Transition " & transition_counter + 2) Then
            last_eval_row = transition_dict("Transition " & transition_counter + 2)("Row") - 1
        Else
            last_eval_row = UBound(data_set, 1)
        End If

        row_end = id_sp_reached(row_start, sp_end, data_set, last_eval_row, sp_start)
        If row_end > last_eval_row Then
            GoTo NextTransEval
        Else
            If wks_exists(ws_name, sp_start, sp_end) = False Then
                Sheets.Add
                ActiveSheet.Name = ws_name
            End If
            With Sheets(ws_name)
                .Activate
                time_col = 1
                .Cells(row_counter, time_col).Value = "Time"
                sp_col = 2
                .Cells(row_counter, sp_col).Value = "Setpoint"
                input_col = .Cells(2, .Columns.Count).End(xlToLeft).Column + 1
                ct_col = input_col
                .Cells(row_counter - 1, input_col).Value = chamber_sn
                .Cells(row_counter, ct_col).Value = "Chamber Temperature"
                unload_transitions row_start, row_end, row_counter, time_col, sp_col, ct_col, data_set, sp_end, ws_name
            End With
        End If
NextTransEval:
    Next
End Sub
Private Function collect_data_sets(ByVal chamber_sn)
    Dim last_col As Integer
    Dim last_row As Integer
    Dim arr As Variant
    '==================================================================================================================
    'For each worksheet the user identified to analyze the data, add the data to a Dictionary where the key is the
    'Worksheet name and the values is the Range from A1 to the last used row and column
    '==================================================================================================================
    With Sheets(chamber_sn)
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
        last_col = .Cells(2, .Columns.Count).End(xlToLeft).Column
        collect_data_sets = .Range(.Cells(1, 1), .Cells(last_row, last_col))
    End With
End Function
Private Function collect_transitions(ByRef data_set) ' As Dictionary

    Dim setpoint_column As Integer
    Dim column_count As Integer
    Dim row_count As Integer
    Dim dictionary_collection As Object
    Dim dictionary_collection_combined As Object
    Set dictionary_collection_combined = CreateObject("Scripting.Dictionary")
    '=====================================================================================================================
    'Find the Setpoint column
    '=====================================================================================================================
    For column_count = LBound(data_set, 2) To UBound(data_set, 2)
        If InStr(1, UCase(data_set(1, column_count)), "TEMPERATURE SP") > 0 Then
            setpoint_column = column_count
            Exit For
        End If
    Next
    '=====================================================================================================================
    'Add the Row and the Setpoint Temp to a Dictionary Each time the Setpoint changes
    '=====================================================================================================================
    For row_count = LBound(data_set, 1) To UBound(data_set, 1)
        If IsNumeric(data_set(row_count, setpoint_column)) = True And row_count > 1 Then
            If data_set(row_count, setpoint_column) <> data_set(row_count - 1, setpoint_column) Then
                Set dictionary_collection = CreateObject("Scripting.Dictionary")
                dictionary_collection.Add "Row", row_count
                dictionary_collection.Add "SP", data_set(row_count, setpoint_column)
                dictionary_collection_combined.Add "Transition " & dictionary_collection_combined.Count + 1, dictionary_collection
            End If
        End If
    Next
    Set collect_transitions = dictionary_collection_combined
End Function
Private Function wks_exists(ByRef ws_name, ByVal sp_start, ByVal sp_end)
    Dim wks As Worksheet
    Dim msg_box_str As String
    Dim msg_box_ans As Integer
    Dim wks_start_temp As String
    Dim wks_end_temp As String
    
    For Each wks In Worksheets
        If wks.Name = ws_name Then
            wks_exists = True
            Exit For
        End If
    Next
    If wks_exists = False Then
        For Each wks In Worksheets
            If InStr(1, wks.Name, "Transition") > 0 Then
                wks_start_temp = Mid(wks.Name, InStr(1, wks.Name, " ") + 1, InStr(1, wks.Name, "To") - InStr(1, wks.Name, " ") - 2)
                wks_end_temp = Mid(wks.Name, InStr(1, wks.Name, "To") + 3)
                If Abs(wks_start_temp - sp_start) <= 15 And Abs(wks_end_temp - sp_end) <= 15 Then
                    msg_box_str = "Would you like to add the transition from " & sp_start & " to " & sp_end & " to " & wks.Name & "?"
                    msg_box_ans = MsgBox(msg_box_str, vbYesNo)
                    If msg_box_ans = vbYes Then
                        ws_name = wks.Name
                        wks_exists = True
                        Exit For
                    End If
                End If
            End If
        Next
    End If
End Function
Private Function id_sp_reached(ByVal row_start, ByVal sp_end, ByRef data_set, ByVal last_eval_row, ByVal sp_start)
    Dim row_counter As Integer
    Dim column_counter As Integer
    Dim ct_column As Integer
    Dim setpoint_column As Integer
    Dim start_temp As Double
    Dim closest_temp As Double
    Dim closest_temp_row As Integer
    Dim msg_box_ans As Integer
    Dim msg_not_found As String

    For column_counter = LBound(data_set, 2) To UBound(data_set, 2)
        If InStr(1, UCase(data_set(1, column_counter)), "TEMPERATURE PV") > 0 Then
            ct_column = column_counter
            Exit For
        End If
    Next

    closest_temp = sp_end - sp_start
    closest_temp_row = row_start
    start_temp = data_set(row_start, ct_column)

    If sp_start > sp_end Then       'Chamber pull down
        For row_counter = row_start To last_eval_row
            If data_set(row_counter, ct_column) <= sp_end Then
                Exit For
            ElseIf Abs(closest_temp) > Abs(sp_end - data_set(row_counter, ct_column)) Then
                closest_temp = sp_end - data_set(row_counter, ct_column)
                closest_temp_row = row_counter
            End If
        Next
    Else
        For row_counter = row_start To last_eval_row    'Chamber heat up
            If data_set(row_counter, ct_column) >= sp_end Then
                Exit For
            ElseIf Abs(closest_temp) > Abs(sp_end - data_set(row_counter, ct_column)) Then
                closest_temp = sp_end - data_set(row_counter, ct_column)
                closest_temp_row = row_counter
            End If
        Next
    End If

    If row_counter <= last_eval_row Then
        id_sp_reached = row_counter
    ElseIf Abs(data_set(closest_temp_row, ct_column) - sp_end) <= 15 Then
        msg_not_found = "The transition from " & sp_start & " to " & sp_end & " did not reach setpoint." & vbCrLf
        msg_not_found = msg_not_found & "The following is a summary of the transition: " & vbCrLf
        msg_not_found = msg_not_found & vbTab & "Setpoint Start Temp" & vbTab & "|" & "  " & sp_start & vbCrLf
        msg_not_found = msg_not_found & vbTab & "Actual Start Temp" & vbTab & "|" & "  " & start_temp & vbCrLf
        msg_not_found = msg_not_found & vbTab & "Setpoint End Temp" & vbTab & "|" & "  " & sp_end & vbCrLf
        msg_not_found = msg_not_found & vbTab & "Closest End Temp" & vbTab & "|" & "  " & data_set(closest_temp_row, ct_column) & vbCrLf
        msg_not_found = msg_not_found & "Would you like to evaluate the transition from " & sp_start & " to " & data_set(closest_temp_row, ct_column) & "?"

        msg_box_ans = MsgBox(msg_not_found, vbYesNo)
        If msg_box_ans = vbYes Then
            id_sp_reached = closest_temp_row
        Else
            id_sp_reached = row_counter
        End If
    Else
        id_sp_reached = row_counter
    End If
End Function
Private Sub unload_transitions(ByVal row_start, ByVal row_end, ByVal row_counter, ByVal time_col, ByVal sp_col, ByVal ct_col, ByRef data_set, ByVal sp_end, ByVal ws_name)
    Dim current_row As Integer
    Dim time_increment As Variant
    Dim column_counter As Integer
    Dim chamber_temp_column As Integer
    Dim time_counter As Variant


    current_row = row_start
    time_increment = data_set(current_row + 1, 2) - data_set(current_row, 2)
    time_counter = 0
    For column_counter = LBound(data_set, 2) To UBound(data_set, 2)
        If InStr(1, UCase(data_set(1, column_counter)), "TEMPERATURE PV") > 0 Then
            chamber_temp_column = column_counter
            Exit For
        End If
    Next

    Do While current_row <= row_end
        row_counter = row_counter + 1
        If current_row = row_start Then
            Sheets(ws_name).Cells(row_counter, time_col).NumberFormat = "h:mm:ss"
            Sheets(ws_name).Cells(row_counter, time_col) = TimeSerial(Hour(0), Minute(0), Second(0))
        Else
            Sheets(ws_name).Cells(row_counter, time_col).NumberFormat = "h:mm:ss"
            Sheets(ws_name).Cells(row_counter, time_col) = TimeSerial(Hour(time_counter), Minute(time_counter), Second(time_counter))
        End If
        Sheets(ws_name).Cells(row_counter, sp_col) = sp_end
        Sheets(ws_name).Cells(row_counter, ct_col) = data_set(current_row, chamber_temp_column)
        current_row = current_row + 1
        time_counter = time_counter + time_increment
    Loop
End Sub

Private Sub autosize_columns()
    Dim wks As Worksheet
    Dim last_col As Integer

    For Each wks In Worksheets
        With Sheets(wks.Name)
            .Activate
            last_col = .Cells(2, .Columns.Count).End(xlToLeft).Column
            With .Range(Columns(1), Columns(last_col))
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .EntireColumn.AutoFit
            End With
        End With
    Next
End Sub
Private Sub create_plots()

    Dim wks As Worksheet
    Dim last_col As Integer
    Dim col_count As Integer
    Dim time_col As Integer
    Dim last_row As Integer
    Dim eval_trans As Integer
    Dim series_counter As Integer
    Dim collection_counter As Variant
    Dim eval_trans_msg As String
    
    Dim series_coll As Series

    series_counter = 0
    For Each wks In Worksheets
        If wks.Name Like "Transition*" Then
            eval_trans_msg = "Would you like to plot the " & wks.Name & "?"
            eval_trans = MsgBox(eval_trans_msg, vbYesNo)
            If eval_trans = vbYes Then
                wks.Activate
                With Sheets(wks.Name)
                    last_col = .Cells(2, .Columns.Count).End(xlToLeft).Column
                    For col_count = 1 To last_col
                        last_row = .Cells(.Rows.Count, col_count).End(xlUp).Row
                        Select Case .Cells(2, col_count).Value
                            Case "Setpoint"
                                time_col = 1
                                series_counter = series_counter + 1
                                .Shapes.AddChart.Select
                                ActiveChart.ChartType = xlXYScatterSmoothNoMarkers
                                For Each series_coll In ActiveChart.SeriesCollection
                                    series_coll.Delete
                                Next
                                ActiveChart.SeriesCollection.NewSeries
                                ActiveChart.SeriesCollection(series_counter).Name = .Cells(2, col_count)
                                ActiveChart.SetElement (msoElementChartTitleAboveChart)
                                ActiveChart.ChartTitle.Text = wks.Name
                            Case "Chamber Temperature"
                                time_col = 1
                                series_counter = series_counter + 1
                                ActiveChart.SeriesCollection.NewSeries
                                ActiveChart.SeriesCollection(series_counter).Name = .Cells(1, col_count)
                            Case Else
                                GoTo NextColumn
                        End Select
                        ActiveChart.SeriesCollection(series_counter).XValues = .Range(.Cells(3, time_col), .Cells(last_row, time_col))
                        ActiveChart.SeriesCollection(series_counter).Values = .Range(.Cells(3, col_count), .Cells(last_row, col_count))
NextColumn:
                    Next
                End With
                ActiveChart.Location Where:=xlLocationAsNewSheet, Name:=wks.Name & "_Chart"
                series_counter = 0
            End If
        End If
    Next

End Sub




