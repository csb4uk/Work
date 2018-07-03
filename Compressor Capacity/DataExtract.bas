Attribute VB_Name = "DataExtract"
Public Sub data_extract_1(ByRef count_1 As Integer, ByVal count_2 As Integer, ByVal num_rows_ref As Integer, ByVal comp As String, _
                          ByRef comp_log_arr As Variant, ByRef comp_options_arr As Variant)
    'While searching through all the rows check to see if column B of the row is a compressor match.  If it is extract the data
    Do While count_2 <= num_rows_ref - 1
        'If compressor is a match then populate data to the userform, otherwise go to the next row.  The data will only need to be populated by one match because
        'it will be the same for each frequency so once it finds a match it will not go through all the loops.  Time saving measure
        If comp = comp_log_arr(count_2, 2) Then
            If UserForm1.comp_label_2 = "Compressor" Then
                UserForm1.comp_label_2 = comp_log_arr(count_2, 1)     'Assigns compressor type variable to UserForm
                UserForm1.hp_label = comp_log_arr(count_2, 4) & " HP"  'Assigns HP variable to UserForm
            End If
            'Determine if the compressor is 50Hz or 60Hz and assign the values to the proper boxes in the UserForm
            Select Case comp_log_arr(count_2, 7)
                Case 50
                    If UserForm1.txt_disp_50 = "" Then
                        With UserForm1
                            .txt_hz_50 = comp_log_arr(count_2, 7)          'Assigns hz to UserForm
                            .txt_rpm_50 = comp_log_arr(count_2, 8)         'Assigns rpm to UserForm
                            .txt_disp_50 = comp_log_arr(count_2, 9)        'Assigns Displacement to UserForm
                        End With
                    End If
                Case 60
                    If UserForm1.txt_disp_60 = "" Then
                        With UserForm1
                            .txt_hz_60 = comp_log_arr(count_2, 7)          'Assigns hz to UserForm
                            .txt_rpm_60 = comp_log_arr(count_2, 8)         'Assigns rpm to UserForm
                            .txt_disp_60 = comp_log_arr(count_2, 9)        'Assigns Displacement to UserForm
                        End With
                    End If
            End Select
            'Store Voltage, Hz, Phase and Primary Refrigerant to be used as possible option button combinations
            ReDim Preserve comp_options_arr(0 To 12, 0 To count_1)
            comp_options_arr(0, count_1) = CStr(comp_log_arr(count_2, 5))
            comp_options_arr(1, count_1) = CStr(comp_log_arr(count_2, 7))
            comp_options_arr(2, count_1) = CStr(comp_log_arr(count_2, 6))
            comp_options_arr(3, count_1) = CStr(comp_log_arr(count_2, 15))
            comp_options_arr(4, count_1) = CStr(comp_log_arr(count_2, 16))
            comp_options_arr(5, count_1) = CStr(comp_log_arr(count_2, 17))
            comp_options_arr(6, count_1) = CStr(comp_log_arr(count_2, 18))
            comp_options_arr(7, count_1) = CStr(comp_log_arr(count_2, 19))
            comp_options_arr(8, count_1) = CStr(comp_log_arr(count_2, 20))
            comp_options_arr(9, count_1) = CStr(comp_log_arr(count_2, 21))
            comp_options_arr(10, count_1) = CStr(comp_log_arr(count_2, 22))
            comp_options_arr(11, count_1) = CStr(comp_log_arr(count_2, 23))
            comp_options_arr(12, count_1) = CStr(comp_log_arr(count_2, 24))
            count_1 = count_1 + 1
        End If
    count_2 = count_2 + 1
    Loop
End Sub
Public Sub ob_eval(ByRef cond_arr As Variant)
    Dim cond_ctrl As Control
    For Each cond_ctrl In UserForm1.v_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = True Then
                cond_arr(0) = cond_ctrl.Caption
            End If
        End If
    Next cond_ctrl
    For Each cond_ctrl In UserForm1.hz_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = True Then
                cond_arr(1) = cond_ctrl.Caption
            End If
        End If
    Next cond_ctrl
    For Each cond_ctrl In UserForm1.ph_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = True Then
                cond_arr(2) = cond_ctrl.Caption
            End If
        End If
    Next cond_ctrl
    For Each cond_ctrl In UserForm1.prim_ref_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = True Then
                cond_arr(3) = cond_ctrl.Caption
            End If
        End If
    Next cond_ctrl
        If cond_arr(3) = "" Then
            MsgBox "Please Select Primary Refrigerant"
            Exit Sub
        End If
    For Each cond_ctrl In UserForm1.casc_ref_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = True Then
                cond_arr(4) = cond_ctrl.Caption
            End If
        End If
    Next cond_ctrl
End Sub
Public Sub data_extract_2(ByVal count_5 As Integer, ByVal count_6 As Integer, ByVal num_rows_ref As Integer, ByVal comp As String, _
                          ByRef comp_log_arr As Variant, ByRef cond_arr_1 As Variant, ByRef new_comp_options_arr As Variant)
Dim opt_ctrl_2 As Control

        'While searching through all the rows check to see if column B of the row is a compressor match.  If it is extract the data into new_comp_options_arr as a String
        Do While count_6 <= num_rows_ref - 1
            If comp = comp_log_arr(count_6, 2) _
                And (cond_arr_1(0) = comp_log_arr(count_6, 5) Or cond_arr_1(0) = "") _
                And (cond_arr_1(1) = comp_log_arr(count_6, 7) Or cond_arr_1(1) = "") _
                And (cond_arr_1(2) = comp_log_arr(count_6, 6) Or cond_arr_1(2) = "") _
                And (cond_arr_1(3) = comp_log_arr(count_6, 15) Or cond_arr_1(3) = comp_log_arr(count_6, 16) Or cond_arr_1(3) = comp_log_arr(count_6, 17) _
                    Or cond_arr_1(3) = comp_log_arr(count_6, 18) Or cond_arr_1(3) = comp_log_arr(count_6, 19) Or cond_arr_1(3) = comp_log_arr(count_6, 20) _
                    Or cond_arr_1(3) = comp_log_arr(count_6, 21) Or cond_arr_1(3) = comp_log_arr(count_6, 22) Or cond_arr_1(3) = comp_log_arr(count_6, 23) _
                    Or cond_arr_1(3) = comp_log_arr(count_6, 24) Or cond_arr_1(3) = "") Then
                        ReDim Preserve new_comp_options_arr(0 To 3, 0 To count_5)
                            new_comp_options_arr(0, count_5) = CStr(comp_log_arr(count_6, 5))
                            new_comp_options_arr(1, count_5) = CStr(comp_log_arr(count_6, 7))
                            new_comp_options_arr(2, count_5) = CStr(comp_log_arr(count_6, 6))
                            new_comp_options_arr(3, count_5) = CStr(cond_arr_1(3))
                            count_5 = count_5 + 1
            End If
        count_6 = count_6 + 1
        Loop
        
        'Run through all the controls again and change the color and lock status if it is an available option
        For Each opt_ctrl_2 In UserForm1.comp_control_frame.Controls
            If TypeOf opt_ctrl_2 Is MSForms.OptionButton Then
                If Left(opt_ctrl_2.Caption, 2) = "R-" And cond_arr_1(3) = "" Then
                Else
                    For count_10 = 0 To count_5 - 1
                        For count_9 = 0 To 3
                            If new_comp_options_arr(count_9, count_10) = opt_ctrl_2.Caption Then
                                opt_ctrl_2.ForeColor = &H80000012
                                opt_ctrl_2.Locked = False
                                GoTo NextBreak
                            End If
                        Next count_9
                    Next count_10
                End If
            End If
NextBreak:
        Next opt_ctrl_2
End Sub
Public Sub data_extract_3(ByRef count_11 As Variant, ByVal count_12 As Integer, ByVal num_rows_ref As Integer, ByVal comp As String, _
                          ByRef comp_log_arr As Variant, ByRef cond_arr_2 As Variant, ByRef new_comp_options_arr_1 As Variant, ByRef comp_type As Variant)
    Do While count_12 <= num_rows_ref - 1
        If comp = comp_log_arr(count_12, 2) _
            And (cond_arr_2(0) = CStr(comp_log_arr(count_12, 5)) Or cond_arr_2(0) = "") _
            And (cond_arr_2(1) = CStr(comp_log_arr(count_12, 7)) Or cond_arr_2(1) = "") _
            And (cond_arr_2(2) = CStr(comp_log_arr(count_12, 6)) Or cond_arr_2(2) = "") _
            And (cond_arr_2(3) = CStr(comp_log_arr(count_12, 15)) Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 16)) Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 17)) _
                Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 18)) Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 19)) Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 20)) _
                Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 21)) Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 22)) Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 23)) _
                Or cond_arr_2(3) = CStr(comp_log_arr(count_12, 24)) Or cond_arr_2(0) = "") Then
                    ReDim Preserve new_comp_options_arr_1(0 To 1, count_11)
                        'Store the compressor code and Hz to be evaluated later
                        If comp_type = "" Then
                            comp_type = CStr(comp_log_arr(count_12, 1))
                        End If
                        new_comp_options_arr_1(0, count_11) = CStr(comp_log_arr(count_12, 3))
                        new_comp_options_arr_1(1, count_11) = CStr(comp_log_arr(count_12, 7))
                        count_11 = count_11 + 1
        End If
    count_12 = count_12 + 1
    Loop
End Sub
Public Sub find_files(ByRef file_name As Variant, ByRef file_path As Variant, ByVal comp As String, ByVal FOLDER_PATH As String)
    file_path = FOLDER_PATH & "\Hermetic\" & comp & "\Master Compressor Capacity Information\"
    file_name = Dir(file_path & "\*.csv")
    If Len(file_name) = 0 Then
        file_path = FOLDER_PATH & "\Scroll\Low Temperature\" & comp & "\Master Compressor Capacity Information\"
        file_name = Dir(file_path & "\*.csv")
        If Len(file_name) = 0 Then
            file_path = FOLDER_PATH & "\Semi-Hermetic\Low Temperature\" & comp & "\Master Compressor Capacity Information\"
            file_name = Dir(file_path & "\*.csv")
            If Len(file_name) = 0 Then
                MsgBox ("File path not found")
            End If
        End If
    End If
End Sub
Public Sub extract_file_name(ByVal file_name As String, ByRef comp_code As Variant, ByRef comp_hz As Variant)
    Dim left_hash, right_hash, length_comp_code, length_hz As Integer
    left_hash = InStr(1, file_name, "-")    'Reads the file of the coefficient sheet to find the first "-".  This allows us to extract the voltage code
    right_hash = InStr(left_hash + 1, file_name, "-")   'Reads the file of the coefficient sheet to find the second "-".  This allows us to extract the voltage code & Hz
    length_comp_code = right_hash - left_hash - 1   'gives the number of characters of the compressor code
    length_hz = 2   'number of characters of the Hz
    comp_code = Mid(file_name, left_hash + 1, length_comp_code) 'Extracts the compressor code from the file name
    comp_hz = Mid(file_name, right_hash + 1, length_hz) 'Extracts the Hz from the file name
End Sub
Public Sub extract_coefficients(ByVal file_path As String, ByVal file_name As String, ByRef cap_coefficients As Variant, ByRef watts_coefficients As Variant, _
                                ByRef mass_flow_coefficients As Variant)
    Dim ref_co_book, ref_co_sheet As String
    Workbooks.Open _
    Filename:=(file_path & file_name), _
    ReadOnly:=True
    ref_co_book = ActiveWorkbook.name   'Assign active book as a variable to write to later on
    ref_co_sheet = ActiveSheet.name     'Assign active sheet as a variable to write to later on
    cap_coefficients = Workbooks(ref_co_book).Worksheets(ref_co_sheet).Range("S2:AB2").Value        'Range of capacity coefficients
    watts_coefficients = Workbooks(ref_co_book).Worksheets(ref_co_sheet).Range("AC2:AL2").Value     'Range of watts coefficients
    mass_flow_coefficients = Workbooks(ref_co_book).Worksheets(ref_co_sheet).Range("AW2:BF2").Value 'Range of mass flow coefficients
    Workbooks(ref_co_book).Close SaveChanges:=False 'Close coefficient workbook
End Sub

