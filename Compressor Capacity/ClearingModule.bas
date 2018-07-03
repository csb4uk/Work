Attribute VB_Name = "ClearingModule"
Public Sub clear_vars_1()
Dim opt_ctrl_reset As Control
    With UserForm1
        .comp_label_2 = "Compressor"
        .hp_label = "HP"
        .txt_disp_50 = ""
        .txt_disp_60 = ""
        .txt_hz_50 = ""
        .txt_hz_60 = ""
        .txt_rpm_50 = ""
        .txt_rpm_60 = ""
    End With
    For Each opt_ctrl_reset In UserForm1.Controls
        If TypeName(opt_ctrl_reset) = "OptionButton" Then
            opt_ctrl_reset.Value = False
            opt_ctrl_reset.ForeColor = &H80000011
            opt_ctrl_reset.Locked = True
        End If
    Next opt_ctrl_reset
End Sub

Public Sub clear_opt_buttons()
Dim cond_ctrl, cond_ctrl_alt As Control
    'For each control in the comp_control_frame, if it is an option button, and is not selected, lock and change color to grey them out
    For Each cond_ctrl In UserForm1.v_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = False Then
                cond_ctrl.ForeColor = &H80000011
                cond_ctrl.Locked = True
            End If
        End If
    Next cond_ctrl
    For Each cond_ctrl In UserForm1.hz_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = False Then
                cond_ctrl.ForeColor = &H80000011
                cond_ctrl.Locked = True
            End If
        End If
    Next cond_ctrl
    For Each cond_ctrl In UserForm1.ph_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = False Then
                cond_ctrl.ForeColor = &H80000011
                cond_ctrl.Locked = True
            End If
        End If
    Next cond_ctrl
    For Each cond_ctrl In UserForm1.prim_ref_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = True Then
                For Each cond_ctrl_alt In UserForm1.prim_ref_frame.Controls
                    If TypeOf cond_ctrl_alt Is MSForms.OptionButton Then
                        If cond_ctrl_alt.Value = False Then
                            cond_ctrl_alt.ForeColor = &H80000011
                            cond_ctrl_alt.Locked = True
                        End If
                    End If
                Next cond_ctrl_alt
                GoTo NextBreakCondCtrlCasc:
             End If
        End If
    Next cond_ctrl
NextBreakCondCtrlCasc:
    For Each cond_ctrl In UserForm1.casc_ref_frame.Controls
        If TypeOf cond_ctrl Is MSForms.OptionButton Then
            If cond_ctrl.Value = True Then
                For Each cond_ctrl_alt In UserForm1.casc_ref_frame.Controls
                    If TypeOf cond_ctrl_alt Is MSForms.OptionButton Then
                        If cond_ctrl_alt.Value = False Then
                            cond_ctrl_alt.ForeColor = &H80000011
                            cond_ctrl_alt.Locked = True
                        End If
                    End If
                Next cond_ctrl_alt
                GoTo NextBreakCondCtrlEnd:
             End If
        End If
    Next cond_ctrl
NextBreakCondCtrlEnd:
End Sub

Public Sub clear_vars_2(ByRef comp_options_arr As Variant, ByVal count_1 As Integer)
'This sub clears the selected options and returns to all the original available options
Dim opt_ctrl_3, opt_ctrl_4 As Control
Dim count_3, count_4 As Integer
    For Each opt_ctrl_3 In UserForm1.Controls
        If TypeName(opt_ctrl_3) = "OptionButton" Then
            opt_ctrl_3.Value = False
        End If
    Next opt_ctrl_3
    For Each opt_ctrl_4 In UserForm1.comp_control_frame.Controls
    If TypeOf opt_ctrl_4 Is MSForms.OptionButton Then
        For count_3 = 0 To count_1 - 1
            For count_4 = 0 To 12
                If comp_options_arr(count_4, count_3) = opt_ctrl_4.Caption Then
                    opt_ctrl_4.ForeColor = &H80000012
                    opt_ctrl_4.Locked = False
                    GoTo NextBreak
                End If
            Next count_4
        Next count_3
    End If
NextBreak:
    Next opt_ctrl_4
End Sub

Public Sub clear_sheets(ByVal source_book As String, ByVal source_sheet As String)
    Workbooks(source_book).Sheets(source_sheet).Range("B11:M100").Delete Shift:=xlUp
    Application.DisplayAlerts = False
    Do While Sheets.Count > Sheets("Smart Sheet Template_Casc").Index
        Sheets(Sheets("Smart Sheet Template_Casc").Index + 1).Delete
    Loop
    Application.DisplayAlerts = True
End Sub
