Attribute VB_Name = "ExtractBudgetHours"

Option Private Module
Option Explicit
Option Base 1

Public Sub extract_budget_hours(ByVal budget_file_path As String, ByVal budget_file_name As String, ByVal job_number As String, _
                                ByRef customer_name As String, ByRef model_number As String, ByRef cab_hours As Double, ByRef electrical_hours As Double, ByRef refrigeration_hours As Double)

    Dim wks As Worksheet

    Dim current_ws_array As Variant
    Dim customer_name_array() As Variant
    Dim model_number_array() As Variant
    Dim cab_hours_array() As Variant
    Dim electrical_hours_array() As Variant
    Dim refrigeration_hours_array() As Variant

    Dim ref_book As String
    Dim ref_sheet As String
    Dim budget_hours_str As String
    Dim msg As String
    Dim ans_continue As String

    Dim current_ws_last_row As Integer
    Dim ref_row_counter As Integer
    Dim number_of_budget_sht As Integer
    Dim budget_sheet_counter As Integer
    Dim array_counter As Integer
    Dim ans_input_box As Integer

    Workbooks.Open _
        Filename:=budget_file_path & budget_file_name, _
        ReadOnly:=True


    ref_book = ActiveWorkbook.Name
    For Each wks In Worksheets
        If wks.Name Like "BUDGET*" And wks.Name Like "BUDGET HOOD*" = False Then
            number_of_budget_sht = number_of_budget_sht + 1
        End If
    Next wks

    If number_of_budget_sht > 1 Then
        For Each wks In Worksheets
            If wks.Name Like "BUDGET*" And wks.Name Like "BUDGET HOOD*" = False Then
                If InStr(1, Workbooks(ref_book).Sheets(wks.Name).Range("B3").Value, "SOLD") > 0 Then
                    Worksheets(wks.Name).Activate
                    ref_sheet = ActiveSheet.Name
                    current_ws_last_row = Workbooks(ref_book).Sheets(wks.Name).Cells(Workbooks(ref_book).Sheets(wks.Name).Rows.Count, "A").End(xlUp).Row
                    current_ws_array = Workbooks(ref_book).Sheets(wks.Name).Range("A1" & ":V" & current_ws_last_row).Value
                    budget_sheet_counter = budget_sheet_counter + 1
                    ReDim Preserve customer_name_array(budget_sheet_counter)
                    ReDim Preserve model_number_array(budget_sheet_counter)
                    ReDim Preserve cab_hours_array(budget_sheet_counter)
                    ReDim Preserve electrical_hours_array(budget_sheet_counter)
                    ReDim Preserve refrigeration_hours_array(budget_sheet_counter)
                    Call EvaluateBudgetHours.evaluate_budget_array(current_ws_array, customer_name_array(budget_sheet_counter), _
                                            model_number_array(budget_sheet_counter), cab_hours_array(budget_sheet_counter), _
                                            electrical_hours_array(budget_sheet_counter), refrigeration_hours_array(budget_sheet_counter))
                    customer_name = customer_name_array(budget_sheet_counter)
                    model_number = model_number_array(budget_sheet_counter)
                    cab_hours = cab_hours_array(budget_sheet_counter)
                    electrical_hours = electrical_hours_array(budget_sheet_counter)
                    refrigeration_hours = refrigeration_hours_array(budget_sheet_counter)
                    GoTo CloseBreak
                End If
            End If
        Next wks
        If ref_sheet = "" Then
            For Each wks In Worksheets
                If wks.Name Like "BUDGET*" And wks.Name Like "BUDGET HOOD*" = False Then
                    Worksheets(wks.Name).Activate
                    current_ws_last_row = Workbooks(ref_book).Sheets(wks.Name).Cells(Workbooks(ref_book).Sheets(wks.Name).Rows.Count, "A").End(xlUp).Row
                    current_ws_array = Workbooks(ref_book).Sheets(wks.Name).Range("A1" & ":V" & current_ws_last_row).Value
                    budget_sheet_counter = budget_sheet_counter + 1
                    ReDim Preserve customer_name_array(budget_sheet_counter)
                    ReDim Preserve model_number_array(budget_sheet_counter)
                    ReDim Preserve cab_hours_array(budget_sheet_counter)
                    ReDim Preserve electrical_hours_array(budget_sheet_counter)
                    ReDim Preserve refrigeration_hours_array(budget_sheet_counter)
                    Call EvaluateBudgetHours.evaluate_budget_array(current_ws_array, customer_name_array(budget_sheet_counter), _
                                            model_number_array(budget_sheet_counter), cab_hours_array(budget_sheet_counter), _
                                            electrical_hours_array(budget_sheet_counter), refrigeration_hours_array(budget_sheet_counter))
                End If
            Next wks
            msg = "Please select the budget hours you would like to import for job number " & job_number & ":" & vbCrLf & vbCrLf
            For array_counter = 1 To budget_sheet_counter
                budget_hours_str = budget_hours_str & array_counter & ". " & vbTab & "Customer Name: " & customer_name_array(array_counter) & _
                                                                                 vbCrLf & vbTab & "Model Number: " & model_number_array(array_counter) & _
                                                                                 vbCrLf & vbTab & "Cab Hours: " & cab_hours_array(array_counter) & _
                                                                                 vbCrLf & vbTab & "Electrical Hours: " & electrical_hours_array(array_counter) & _
                                                                                 vbCrLf & vbTab & "Refrigeration Hours: " & refrigeration_hours_array(array_counter) & vbCrLf & vbCrLf
            Next array_counter

            On Error GoTo ErrIB2
SelectSheet:
            ans_input_box = InputBox(msg & budget_hours_str)
            customer_name = customer_name_array(ans_input_box)
            model_number = model_number_array(ans_input_box)
            cab_hours = cab_hours_array(ans_input_box)
            electrical_hours = electrical_hours_array(ans_input_box)
            refrigeration_hours = refrigeration_hours_array(ans_input_box)
            GoTo CloseBreak
ErrIB2:
            ans_continue = MsgBox("Would you like to skip this entry?", vbYesNo)
            If ans_continue = vbNo Then
                Resume SelectSheet
            Else
                GoTo CloseBreak
            End If
        End If
    Else
        For Each wks In Worksheets
            If wks.Name Like "BUDGET*" And wks.Name Like "BUDGET HOOD*" = False Then
                Worksheets(wks.Name).Activate
                current_ws_last_row = Workbooks(ref_book).Sheets(wks.Name).Cells(Workbooks(ref_book).Sheets(wks.Name).Rows.Count, "A").End(xlUp).Row
                current_ws_array = Workbooks(ref_book).Sheets(wks.Name).Range("A1" & ":V" & current_ws_last_row).Value
                budget_sheet_counter = budget_sheet_counter + 1
                ReDim Preserve customer_name_array(budget_sheet_counter)
                ReDim Preserve model_number_array(budget_sheet_counter)
                ReDim Preserve cab_hours_array(budget_sheet_counter)
                ReDim Preserve electrical_hours_array(budget_sheet_counter)
                ReDim Preserve refrigeration_hours_array(budget_sheet_counter)
                Call EvaluateBudgetHours.evaluate_budget_array(current_ws_array, customer_name_array(budget_sheet_counter), _
                                            model_number_array(budget_sheet_counter), cab_hours_array(budget_sheet_counter), _
                                            electrical_hours_array(budget_sheet_counter), refrigeration_hours_array(budget_sheet_counter))
                customer_name = customer_name_array(budget_sheet_counter)
                model_number = model_number_array(budget_sheet_counter)
                cab_hours = cab_hours_array(budget_sheet_counter)
                electrical_hours = electrical_hours_array(budget_sheet_counter)
                refrigeration_hours = refrigeration_hours_array(budget_sheet_counter)
                Exit For
            End If
        Next wks
    End If

CloseBreak:
    Workbooks(budget_file_name).Close SaveChanges:=False
End Sub

