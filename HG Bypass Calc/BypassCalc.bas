Attribute VB_Name = "BypassCalc"
Option Explicit

Public Sub hg_bypass_calc()


    Dim refrigerant As String
    Dim wks_name As String
    Dim data_sheet As String

    Dim last_col As Integer
    Dim last_row As Integer
    Dim x1 As Integer
    Dim x3 As Integer
    Dim x1_col As Integer
    Dim x3_col As Integer
    Dim z1 As Integer
    Dim z3 As Integer
    Dim x1_z1_col As Integer
    Dim x1_z3_col As Integer
    Dim x3_z1_col As Integer
    Dim x3_z3_col As Integer

    Dim SST As Double
    Dim SCT As Double
    Dim cap As Double

    Dim lin_int As Boolean

    Dim sst_coll As Collection
    Set sst_coll = New Collection

    Dim sct_coll As Collection
    Set sct_coll = New Collection

    Dim valve_arr() As Variant
    ReDim Preserve valve_arr(0 To 1, 0)

    wks_name = ActiveSheet.Name
    data_sheet = "Data"
    refrigerant = Sheets(wks_name).Cells(2, 2)
    SST = return_temp(Sheets(wks_name).Cells(3, 2), Sheets(wks_name).Cells(3, 3))
    SCT = return_temp(Sheets(wks_name).Cells(4, 2), Sheets(wks_name).Cells(4, 3))
    cap = return_cap(Sheets(wks_name).Cells(5, 2), Sheets(wks_name).Cells(5, 3))

    With Sheets(data_sheet)
        last_col = .Cells(1, .Columns.Count).End(xlToLeft).Column
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    clear_table (wks_name)
    
    add_coll sst_coll, sct_coll, last_col, 1, data_sheet


    '=============Complete==================

    If Contains(sst_coll, SST) And Contains(sct_coll, SCT) Then
        return_col_match last_col, data_sheet, SST, SCT, cap, valve_arr, last_row, refrigerant
    ElseIf Contains(sst_coll, SST) Then
        return_cols x1, SCT, x3, sct_coll, "SCT"
        id_cols x1, x3, x1_col, x3_col, data_sheet, last_col, SST, "SST"
        int_temp x1, x3, x1_col, x3_col, SCT, last_row, data_sheet, refrigerant, cap, valve_arr
    ElseIf Contains(sct_coll, SCT) Then
        return_cols x1, SST, x3, sst_coll, "SST"
        id_cols x1, x3, x1_col, x3_col, data_sheet, last_col, SCT, "SCT"
        int_temp x1, x3, x1_col, x3_col, SST, last_row, data_sheet, refrigerant, cap, valve_arr
    Else
        return_cols x1, SCT, x3, sct_coll, "SCT"
        return_cols z1, SST, z3, sst_coll, "SST"
        id_cols_bilinear x1, SCT, x3, z1, SST, z3, x1_z1_col, x1_z3_col, x3_z1_col, x3_z3_col, last_col, data_sheet
        bilinear_interpolation x1, SCT, x3, z1, SST, z3, x1_z1_col, x1_z3_col, x3_z1_col, x3_z3_col, refrigerant, valve_arr, data_sheet, last_row, cap
    '=======================================
    End If
    Unload_Arr wks_name, valve_arr
End Sub


Private Sub clear_table(ByVal wks_name)
    Dim last_row As Integer
    Dim row_counter As Integer
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With Sheets(wks_name)
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    
    For row_counter = last_row To 1 Step -1
        If Sheets(wks_name).Cells(row_counter, 2).Value <> "Capacity" Then
            Sheets(wks_name).Cells(row_counter, 2).Value = ""
        Else
            Exit For
        End If
    Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function return_temp(ByVal temp, ByVal units)
    Select Case units
        Case "°F"
            return_temp = temp
        Case "°C"
            return_temp = (temp * 1.8) + 32
    End Select
End Function
Private Function return_cap(ByVal cap, ByVal units)
    Select Case units
        Case "Tons"
            return_cap = cap
        Case "BTUs"
            return_cap = cap / 12000
    End Select
End Function
Private Sub add_coll(ByRef sst_coll, ByRef sct_coll, ByVal last_col, ByVal current_row, ByVal sht_name)
    Dim col_counter As Integer
    Dim key As Variant
    Dim key_start As Integer
    Dim key_end As Integer
    Dim key_sst As Integer
    Dim key_sct As Integer
    For col_counter = 1 To last_col
        key = Sheets(sht_name).Cells(current_row, col_counter).Value
        If InStr(1, key, "SST") > 0 Then
        
            key_start = return_start_key(key, "SST")
            key_end = return_end_key(1, key, "°F", key_start)
            key_sst = Mid(key, key_start, key_end)
            
            key_start = return_start_key(key, "SCT")
            key_end = return_end_key(InStr(1, key, "SCT"), key, "°F", key_start)
            key_sct = Mid(key, key_start, key_end)
            
            If IsNumeric(key_sst) Then
                If Contains(sst_coll, key_sst) = False Then
                    sst_coll.Add key_sst
                End If
                If Contains(sct_coll, key_sct) = False Then
                    sct_coll.Add key_sct
                End If
            End If
        End If
    Next
End Sub
Private Function return_start_key(key, search_str)
    return_start_key = InStr(1, key, search_str) + 4
End Function
Private Function return_end_key(start_loc, key, search_str, key_start)
    return_end_key = InStr(start_loc, key, search_str) - key_start
End Function
Private Function Contains(ByVal col, ByVal key) As Boolean
    Dim count_col As Integer
    Contains = False
    For count_col = 1 To col.Count
        If col(count_col) = key Then
            Contains = True
            Exit Function
        End If
    Next
End Function
Private Sub return_col_match(ByVal last_col, ByVal wks_name, ByVal SST, ByVal SCT, ByVal cap, ByRef valve_arr, ByVal last_row, ByVal refrigerant)

    Dim column_counter As Integer
    Dim key_sst As Integer
    Dim key_sct As Integer
    Dim key_start As Integer
    Dim key_end As Integer
    Dim key As Variant
    Dim row_counter As Integer
    Dim cur_cap As Double
    Dim arr_counter As Integer

    column_counter = 1
    Do While column_counter <= last_col
        key = Sheets(wks_name).Cells(1, column_counter).Value
        If InStr(1, key, "SST") > 0 Then
            key_start = return_start_key(key, "SST")
            key_end = return_end_key(1, key, "°F", key_start)
            key_sst = Mid(key, key_start, key_end)
            
            key_start = return_start_key(key, "SCT")
            key_end = return_end_key(InStr(1, key, "SCT"), key, "°F", key_start)
            key_sct = Mid(key, key_start, key_end)
            If key_sst = SST And key_sct = SCT Then
                Exit Do
            End If
        End If
        column_counter = column_counter + 1
    Loop
    arr_counter = 0
    For row_counter = 2 To last_row
        If Sheets(wks_name).Cells(row_counter, 1) = refrigerant Then
            cur_cap = Sheets(wks_name).Cells(row_counter, column_counter).Value
            If cur_cap >= cap Then
                ReDim Preserve valve_arr(0 To 1, 0 To arr_counter)
                valve_arr(0, arr_counter) = Sheets(wks_name).Cells(row_counter, 2).Value & "-" & Sheets(wks_name).Cells(row_counter, 3).Value
                valve_arr(1, arr_counter) = cur_cap
                arr_counter = arr_counter + 1
            End If
        End If
    Next
End Sub
Private Sub return_cols(ByRef x1, ByVal x2, ByRef x3, ByVal this_coll, ByVal driver_val)
    Dim collection_counter As Integer
    If driver_val = "SCT" Then
        For collection_counter = 2 To this_coll.Count
            x1 = this_coll(collection_counter - 1)
            x3 = this_coll(collection_counter)
            If x1 < x2 And x2 < x3 Then
                Exit For
            End If
        Next
    ElseIf driver_val = "SST" Then
        For collection_counter = 2 To this_coll.Count
            x1 = this_coll(collection_counter - 1)
            x3 = this_coll(collection_counter)
            If x1 > x2 And x2 > x3 Then
                Exit For
            End If
        Next
    Else
        MsgBox ("Error")
    End If
End Sub
Private Sub id_cols(ByVal x1, ByVal x3, ByRef x1_col, ByRef x3_col, ByVal sht_name, ByVal last_col, ByVal match_val, ByVal driver_val)
    Dim column_counter As Integer
    Dim key As Variant
    Dim key_start As Integer
    Dim key_end As Integer
    Dim key_sst As Integer
    Dim key_sct As Integer
    column_counter = 1
    Do While column_counter <= last_col
        key = Sheets(sht_name).Cells(1, column_counter).Value
        If InStr(1, key, "SST") > 0 Then
            key_start = return_start_key(key, "SST")
            key_end = return_end_key(1, key, "°F", key_start)
            key_sst = Mid(key, key_start, key_end)
            
            key_start = return_start_key(key, "SCT")
            key_end = return_end_key(InStr(1, key, "SCT"), key, "°F", key_start)
            key_sct = Mid(key, key_start, key_end)
            If driver_val = "SST" Then
                If key_sst = match_val Then
                    Select Case key_sct
                        Case x1
                            x1_col = column_counter
                        Case x3
                            x3_col = column_counter
                        Case Else
                    End Select
                End If
            ElseIf driver_val = "SCT" Then
                If key_sct = match_val Then
                    Select Case key_sst
                        Case x1
                            x1_col = column_counter
                        Case x3
                            x3_col = column_counter
                        Case Else
                    End Select
                End If
            End If
        End If
        column_counter = column_counter + 1
    Loop
End Sub

Private Sub id_cols_bilinear(ByVal x1, ByVal x2, ByVal x3, ByVal z1, ByVal z2, ByVal z3, _
                                    ByRef x1_z1_col, ByRef x1_z3_col, ByRef x3_z1_col, ByRef x3_z3_col, ByVal last_col, ByVal sht_name)

    Dim column_counter As Integer
    Dim key As Variant
    Dim key_start As Integer
    Dim key_end As Integer
    Dim key_sst As Integer
    Dim key_sct As Integer
    column_counter = 1

    Do While column_counter <= last_col
        key = Sheets(sht_name).Cells(1, column_counter).Value
        If InStr(1, key, "SST") > 0 Then
            key_start = return_start_key(key, "SST")
            key_end = return_end_key(1, key, "°F", key_start)
            key_sst = Mid(key, key_start, key_end)
            
            key_start = return_start_key(key, "SCT")
            key_end = return_end_key(InStr(1, key, "SCT"), key, "°F", key_start)
            key_sct = Mid(key, key_start, key_end)

            Select Case key_sct
                Case x1
                    If key_sst = z1 Then
                        x1_z1_col = column_counter
                    ElseIf key_sst = z3 Then
                        x1_z3_col = column_counter
                    End If
                Case x3
                    If key_sst = z1 Then
                        x3_z1_col = column_counter
                    ElseIf key_sst = z3 Then
                        x3_z3_col = column_counter
                    End If
            End Select
        End If
        column_counter = column_counter + 1
    Loop
End Sub
Private Sub int_temp(ByVal x1, ByVal x3, ByVal x1_col, ByVal x3_col, ByVal x2, ByVal last_row, ByVal wks_name, ByVal refrigerant, ByVal cap, ByRef valve_arr)
    Dim row_counter As Integer
    Dim cur_cap As Double
    Dim arr_counter As Integer
    Dim y1 As Double
    Dim y3 As Double

    arr_counter = 0
    For row_counter = 2 To last_row
        If Sheets(wks_name).Cells(row_counter, 1) = refrigerant Then
            y1 = return_val(wks_name, row_counter, x1_col)
            y3 = return_val(wks_name, row_counter, x3_col)
            If y1 <> 0 And y3 <> 0 Then
                cur_cap = linear_interpolation(x1, x2, x3, y1, y3)
                ReDim Preserve valve_arr(0 To 1, 0 To arr_counter)
                valve_arr(0, arr_counter) = Sheets(wks_name).Cells(row_counter, 2).Value & "-" & Sheets(wks_name).Cells(row_counter, 3).Value
                valve_arr(1, arr_counter) = cur_cap
                arr_counter = arr_counter + 1
            Else
                ReDim Preserve valve_arr(0 To 1, 0 To arr_counter)
                valve_arr(0, arr_counter) = Sheets(wks_name).Cells(row_counter, 2).Value & "-" & Sheets(wks_name).Cells(row_counter, 3).Value
                valve_arr(1, arr_counter) = "Outside Envelope"
                arr_counter = arr_counter + 1
            End If
        End If
    Next
End Sub

Private Sub bilinear_interpolation(ByVal x1, ByVal x2, ByVal x3, ByVal z1, ByVal z2, ByVal z3, _
                                        ByVal x1_z1_col, ByVal x1_z3_col, ByVal x3_z1_col, ByVal x3_z3_col, _
                                        ByVal refrigerant, ByRef valve_arr, ByVal wks_name, ByVal last_row, ByVal cap)

    Dim row_counter As Integer
    Dim cur_cap As Double
    Dim arr_counter As Integer
    Dim x1_z1 As Double
    Dim x1_z3 As Double
    Dim x3_z1 As Double
    Dim x3_z3 As Double

    arr_counter = 0
    For row_counter = 2 To last_row
        If Sheets(wks_name).Cells(row_counter, 1) = refrigerant Then
            x1_z1 = return_val(wks_name, row_counter, x1_z1_col)
            x1_z3 = return_val(wks_name, row_counter, x1_z3_col)
            x3_z1 = return_val(wks_name, row_counter, x3_z1_col)
            x3_z3 = return_val(wks_name, row_counter, x3_z3_col)
            
            If x1_z1 <> 0 And x1_z3 <> 0 And x3_z1 <> 0 And x3_z3 <> 0 Then
                cur_cap = bilin_interp(x1, x2, x3, z1, z2, z3, x1_z1, x1_z3, x3_z1, x3_z3)
                ReDim Preserve valve_arr(0 To 1, 0 To arr_counter)
                valve_arr(0, arr_counter) = Sheets(wks_name).Cells(row_counter, 2).Value & "-" & Sheets(wks_name).Cells(row_counter, 3).Value
                valve_arr(1, arr_counter) = cur_cap
                arr_counter = arr_counter + 1
            Else
                ReDim Preserve valve_arr(0 To 1, 0 To arr_counter)
                valve_arr(0, arr_counter) = Sheets(wks_name).Cells(row_counter, 2).Value & "-" & Sheets(wks_name).Cells(row_counter, 3).Value
                valve_arr(1, arr_counter) = "Outside Envelope"
                arr_counter = arr_counter + 1
            End If
        End If
    Next

End Sub
Private Function return_val(wks_name, row_counter, col_counter)
    If IsNumeric(Sheets(wks_name).Cells(row_counter, col_counter).Value) Then
        return_val = Sheets(wks_name).Cells(row_counter, col_counter).Value
    Else
        return_val = 0
    End If
End Function
Private Function linear_interpolation(ByVal x1, ByVal x2, ByVal x3, ByVal y1, ByVal y3)
    linear_interpolation = (((x2 - x1) * (y3 - y1)) / (x3 - x1)) + y1
End Function
Private Function bilin_interp(ByVal x1, ByVal x, ByVal x2, ByVal z1, ByVal z, ByVal z2, ByVal x1_z1, ByVal x1_z2, ByVal x2_z1, ByVal x2_z2)
    
    Dim val_1 As Double
    Dim val_2 As Double
    Dim val_3 As Double
    Dim val_4 As Double

    val_1 = (((x2 - x) * (z2 - z)) / ((x2 - x1) * (z2 - z1))) * x1_z1
    val_2 = (((x - x1) * (z2 - z)) / ((x2 - x1) * (z2 - z1))) * x1_z2
    val_3 = (((x2 - x) * (z - z1)) / ((x2 - x1) * (z2 - z1))) * x2_z1
    val_4 = (((x - x1) * (z - z1)) / ((x2 - x1) * (z2 - z1))) * x2_z2
    bilin_interp = val_1 + val_2 + val_3 + val_4
End Function
Private Sub Unload_Arr(ByVal wks_name, ByRef valve_arr)

    Dim last_row As Integer
    Dim import_row As Integer
    Dim arr_counter As Integer
    Dim row_counter As Integer

    With Sheets(wks_name)
        last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With
    For arr_counter = LBound(valve_arr, 2) To UBound(valve_arr, 2)
        For row_counter = 1 To last_row
            If Sheets(wks_name).Cells(row_counter, 1) = valve_arr(0, arr_counter) Then
                Sheets(wks_name).Cells(row_counter, 2) = valve_arr(1, arr_counter)
            End If
        Next
    Next
End Sub



