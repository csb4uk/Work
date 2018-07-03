Attribute VB_Name = "PopulateData"
Public Sub comp_pop_data(ByRef num_rows_ref, ByRef comp_log_arr As Variant)
    Workbooks.Open _
        Filename:="R:\NEW R DRIVE\Refrigeration Compressors\Compressor log.xlsm", _
        ReadOnly:=True
    num_rows_ref = Application.WorksheetFunction.CountA(Workbooks("Compressor log").Sheets("Sheet1").Range("B:B")) 'number of rows in reference sheet1
    'Store everything in comp log into an array to cut down processing time
    comp_log_arr = Workbooks("Compressor log").Sheets("Sheet1").Range("A2:X" & num_rows_ref)
    Workbooks("Compressor log").Close SaveChanges:=False
End Sub

Public Sub pop_avail_opt(ByVal count_1 As Integer, ByRef col_opt_event As Collection, ByRef comp_options_arr As Variant)
Dim opt_ctrl_1 As Control   'Used to loop through controls on the UserForm
Dim count_3, count_4 As Integer
    'For every control in the compressor control frame see if the type is an option button.  If it is assign the button to a new collection.
    'This new collection writes to the Class module clsOptEvent.  Now that it is in that module if any option buttons are selected then it will run through the code in the
    'class module.
    For Each opt_ctrl_1 In UserForm1.comp_control_frame.Controls
        If TypeOf opt_ctrl_1 Is MSForms.OptionButton Then
            Set opt_event = New clsOptEvent
            Set opt_event.OptionButtonEvents = opt_ctrl_1
            col_opt_event.Add opt_event
            'Runs through each option button stored in comp_options_arr, makes the text black and unlocks the option button to enable selection if the caption
            'of the option button matches the value in the comp_options_arr
            For count_3 = 0 To count_1 - 1
                For count_4 = 0 To 12
                    If comp_options_arr(count_4, count_3) = opt_ctrl_1.Caption _
                        Or opt_ctrl_1.Caption = "R-410A" _
                        Or opt_ctrl_1.Caption = "R-508B" _
                        Or opt_ctrl_1.Caption = "R-23" Then
                            opt_ctrl_1.ForeColor = &H80000012
                            opt_ctrl_1.Locked = False
                            GoTo NextBreak
                    End If
                Next count_4
            Next count_3
        End If
NextBreak:
    Next opt_ctrl_1
End Sub
    
Public Sub pop_top_table(ByVal source_book As String, ByVal source_sheet As String)
    With Workbooks(source_book).Sheets(source_sheet)
        .Range("B3").Value = UserForm1.comp_selection
        .Range("E3").Value = CInt(UserForm1.txt_hz_60)
        .Range("F3").Value = CInt(UserForm1.txt_rpm_60)
        .Range("G3").Value = Format(CSng(UserForm1.txt_disp_60), "#.00")
        .Range("E4").Value = CInt(UserForm1.txt_hz_50)
        .Range("F4").Value = CInt(UserForm1.txt_rpm_50)
        .Range("G4").Value = Format(CSng(UserForm1.txt_disp_50), "#.00")
    End With
End Sub

Public Sub pop_table_1(ByVal source_book As String, ByVal source_sheet As String, ByVal import_row As Integer, ByVal comp_hz As String, ByVal comp_code As String, _
                       ByRef cap_coefficients As Variant, ByRef watts_coefficients As Variant, ByRef mass_flow_coefficients As Variant)
    With Workbooks(source_book).Sheets(source_sheet)
        .Range("B" & (import_row)).Value = comp_hz & " HZ"
        .Range("C" & (import_row)).Value = comp_code
        .Range("C:C").EntireColumn.AutoFit
        .Range("D" & (import_row + 1) & ":M" & (import_row + 1)).Value = cap_coefficients
        .Range("D" & (import_row + 2) & ":M" & (import_row + 2)).Value = watts_coefficients
        .Range("D" & (import_row + 3) & ":M" & (import_row + 3)).Value = mass_flow_coefficients
    End With
End Sub

Public Sub formula_update(ByVal source_book As String, ByVal comp_hz As String, ByVal comp_code As String, _
                          ByVal import_row As Integer, ByVal act_sheet As String, ByVal refrig As String, ByVal casc_refrig As String, ByVal comp_type As String)
Dim count_16, count_17, end_count_16 As Integer
    Workbooks(source_book).Sheets(act_sheet).Range("C2").Value = "R" & refrig
    If Left(act_sheet, 3) = Left(refrig, 3) Then
        Workbooks(source_book).Sheets(act_sheet).Range("D2").Value = "R" & refrig
    Else
        Workbooks(source_book).Sheets(act_sheet).Range("D2").Value = "R" & casc_refrig
    End If
    If comp_type = "Hermetic" Then
        Workbooks(source_book).Sheets(act_sheet).Range("R6").Value = 40
        Workbooks(source_book).Sheets(act_sheet).Range("W6").Value = 40
    Else
        Workbooks(source_book).Sheets(act_sheet).Range("R6").Value = 65
        Workbooks(source_book).Sheets(act_sheet).Range("W6").Value = 65
    End If
    With Workbooks(source_book).Sheets(act_sheet)
        .Range("K6").Formula = "='Compressor Summary'!$D$" & (import_row + 1) & "+'Compressor Summary'!$E$" & (import_row + 1) & "*C6" & _
            "+'Compressor Summary'!$F$" & (import_row + 1) & "*F6+'Compressor Summary'!$G$" & (import_row + 1) & "*C6^2" & _
            "+'Compressor Summary'!$H$" & (import_row + 1) & "*C6*F6+'Compressor Summary'!$I$" & (import_row + 1) & "*F6^2" & _
            "+'Compressor Summary'!$J$" & (import_row + 1) & "*C6^3+'Compressor Summary'!$K$" & (import_row + 1) & "*F6*C6^2" & _
            "+'Compressor Summary'!$L$" & (import_row + 1) & "*C6*F6^2+'Compressor Summary'!$M$" & (import_row + 1) & "*F6^3"
        .Range("L6").Formula = "='Compressor Summary'!$D$" & (import_row + 2) & "+'Compressor Summary'!$E$" & (import_row + 2) & "*C6" & _
            "+'Compressor Summary'!$F$" & (import_row + 2) & "*F6+'Compressor Summary'!$G$" & (import_row + 2) & "*C6^2" & _
            "+'Compressor Summary'!$H$" & (import_row + 2) & "*C6*F6+'Compressor Summary'!$I$" & (import_row + 2) & "*F6^2" & _
            "+'Compressor Summary'!$J$" & (import_row + 2) & "*C6^3+'Compressor Summary'!$K$" & (import_row + 2) & "*F6*C6^2" & _
            "+'Compressor Summary'!$L$" & (import_row + 2) & "*C6*F6^2+'Compressor Summary'!$M$" & (import_row + 2) & "*F6^3"
        .Range("N6").Formula = "='Compressor Summary'!$D$" & (import_row + 3) & "+'Compressor Summary'!$E$" & (import_row + 3) & "*C6" & _
            "+'Compressor Summary'!$F$" & (import_row + 3) & "*F6+'Compressor Summary'!$G$" & (import_row + 3) & "*C6^2" & _
            "+'Compressor Summary'!$H$" & (import_row + 3) & "*C6*F6+'Compressor Summary'!$I$" & (import_row + 3) & "*F6^2" & _
            "+'Compressor Summary'!$J$" & (import_row + 3) & "*C6^3+'Compressor Summary'!$K$" & (import_row + 3) & "*F6*C6^2" & _
            "+'Compressor Summary'!$L$" & (import_row + 3) & "*C6*F6^2+'Compressor Summary'!$M$" & (import_row + 3) & "*F6^3"
    End With

    'Find the number of rows in the new Smart Sheet
    end_count_16 = Workbooks(source_book).Sheets(act_sheet).Cells(Rows.Count, 3).End(xlUp).Row

    'Determine if the frequency is 60Hz or 50 Hz so you know which displacement to use (aka G3 or G4 in compressor summary).  This impacts cell AJ6
    'as well as all of the mass flow calculations which is the reason for count_16 and end_count_16; so I can loop through those rows
    If comp_hz = 60 Then
        With Workbooks(source_book).Sheets(act_sheet)
                .Range("AJ6").Formula = "=(N6*X6)/'Compressor Summary'!$G$3"
        End With
        For count_16 = 7 To end_count_16
            With Workbooks(source_book).Sheets(act_sheet)
                .Range("N" & count_16).Formula = "=(AL" & count_16 & "*'Compressor Summary'!$G$3)/X" & count_16
            End With
        Next count_16
    Else
        With Workbooks(source_book).Sheets(act_sheet)
                .Range("AJ6").Formula = "=(N6*X6)/'Compressor Summary'!$G$4"
        End With
        For count_16 = 7 To end_count_16
            With Workbooks(source_book).Sheets(act_sheet)
                .Range("N" & count_16).Formula = "=(AL" & count_16 & "*'Compressor Summary'!$G$4)/X" & count_16
            End With
        Next count_16
    End If
    With ActiveSheet
    .Range("AL6").GoalSeek _
        Goal:=.Range("AJ6").Value, _
        ChangingCell:=.Range("AK6")
    End With
    For count_17 = 19 To 80
        With ActiveSheet
            .Range("J" & count_17).GoalSeek _
            Goal:=.Range("K" & count_17).Value, _
            ChangingCell:=.Range("U" & count_17)
        End With
    Next count_17
End Sub

