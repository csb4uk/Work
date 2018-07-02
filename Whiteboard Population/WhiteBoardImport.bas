Attribute VB_Name = "WhiteBoardImport"
Option Explicit


Public Sub import_whiteboard()

    Dim wkb As Workbook
    Dim wks As Worksheet

    Dim master_wb As String
    Dim master_ws As String
    Dim slave_wb As String
    Dim slave_ws As String

    Dim master_ws_last_row As Integer
    Dim master_ws_start_row_custom_jobs As Integer
    Dim master_ws_start_row_standard_jobs As Integer
    Dim master_ws_start_row_service_jobs As Integer
    Dim master_ws_arr_custom_jobs_counter As Integer
    Dim master_ws_arr_standard_jobs_counter As Integer
    Dim master_ws_arr_service_jobs_counter As Integer
    Dim master_ws_row_counter As Integer
    Dim input_month_array_counter As Integer

    Dim master_ws_arr_custom_jobs() As Variant
    Dim master_ws_arr_standard_jobs() As Variant
    Dim master_ws_arr_service_jobs() As Variant
    Dim input_month_array() As Variant
    Dim ans As Variant


NextMonth:
    slave_wb = "Whiteboard Scripts.xlsm"
    slave_ws = "Sheet1"


    input_month_array_counter = input_month_array_counter
    ReDim Preserve input_month_array(input_month_array_counter)
    Call sub_import_month(master_ws, input_month_array, input_month_array_counter)

    'Find the Whiteboard Workbook
    For Each wkb In Workbooks
        If wkb.Name Like "White board schedule, 2017.xlsx" Then
            master_wb = wkb.Name
            wkb.Activate
            Exit For
        End If
    Next wkb

    'Verify that the Whiteboard sheet is active
    If ActiveWorkbook.Name = slave_wb Then
        MsgBox ("The whiteboard file does not appear to be open.  Please open the file before running the program")
        End
    End If

    Workbooks(master_wb).Worksheets(master_ws).Activate
    
    'Find the last used row in the desired excel sheet for the following loop
    master_ws_last_row = Workbooks(master_wb).Sheets(master_ws).Cells(Workbooks(master_wb).Sheets(master_ws).Rows.Count, "A").End(xlUp).Row
    master_ws_row_counter = 1

    'Extract Custom Jobs
    Call extract_jobs(master_ws_row_counter, master_ws_last_row, master_wb, master_ws, master_ws_start_row_custom_jobs)
    master_ws_arr_custom_jobs_counter = master_ws_arr_custom_jobs_counter
    ReDim Preserve master_ws_arr_custom_jobs(master_ws_arr_custom_jobs_counter)
    master_ws_arr_custom_jobs(master_ws_arr_custom_jobs_counter) = Workbooks(master_wb).Sheets(master_ws).Range("A" & master_ws_start_row_custom_jobs & ":Q" & master_ws_row_counter).Value

    'Extract Standard Jobs
    Call extract_jobs(master_ws_row_counter, master_ws_last_row, master_wb, master_ws, master_ws_start_row_standard_jobs)
    master_ws_arr_standard_jobs_counter = master_ws_arr_custom_jobs_counter
    ReDim Preserve master_ws_arr_standard_jobs(master_ws_arr_standard_jobs_counter)
    master_ws_arr_standard_jobs(master_ws_arr_standard_jobs_counter) = Workbooks(master_wb).Sheets(master_ws).Range("A" & master_ws_start_row_standard_jobs & ":Q" & master_ws_row_counter).Value

    'Extract Service Jobs
    Call extract_jobs(master_ws_row_counter, master_ws_last_row, master_wb, master_ws, master_ws_start_row_service_jobs)
    master_ws_arr_service_jobs_counter = master_ws_arr_custom_jobs_counter
    ReDim Preserve master_ws_arr_service_jobs(master_ws_arr_service_jobs_counter)
    master_ws_arr_service_jobs(master_ws_arr_service_jobs_counter) = Workbooks(master_wb).Sheets(master_ws).Range("A" & master_ws_start_row_service_jobs & ":Q" & master_ws_row_counter).Value

    ans = MsgBox("Would you like to import another month?", vbYesNo)
    
    Select Case ans
        Case vbYes
            input_month_array_counter = input_month_array_counter + 1
            master_ws_arr_custom_jobs_counter = master_ws_arr_custom_jobs_counter + 1
            master_ws_arr_standard_jobs_counter = master_ws_arr_custom_jobs_counter + 1
            master_ws_arr_service_jobs_counter = master_ws_arr_custom_jobs_counter + 1

            GoTo NextMonth
        Case vbNo
    End Select

End Sub

Public Sub sub_import_month(ByRef master_ws As String, ByRef input_month_array As Variant, ByVal input_month_array_counter As Integer)

    Dim import_month As Integer

InputMonth:
    import_month = InputBox("Please input a number between 1 and 12" & vbLf & vbLf & "1. January" & vbLf & "2. February" & vbLf & "3. March" & vbLf _
             & "4. April" & vbLf & "5. May" & vbLf & "6. June" & vbLf & "7. July" & vbLf & "8. August" & vbLf & "9. September" & vbLf _
             & "10. October" & vbLf & "11. November" & vbLf & "12. December" & vbLf)

    If import_month = False Then
        GoTo InputMonth
    Else
        Select Case import_month
            Case 1
                master_ws = "JAN"
                input_month_array(input_month_array_counter) = "January"
            Case 2
                master_ws = "FEB"
                input_month_array(input_month_array_counter) = "February"
            Case 3
                master_ws = "MAR"
                input_month_array(input_month_array_counter) = "March"
            Case 4
                master_ws = "APR"
                input_month_array(input_month_array_counter) = "April"
            Case 5
                master_ws = "MAY"
                input_month_array(input_month_array_counter) = "May"
            Case 6
                master_ws = "JUNE"
                input_month_array(input_month_array_counter) = "June"
            Case 7
                master_ws = "JULY"
                input_month_array(input_month_array_counter) = "July"
            Case 8
                master_ws = "AUG"
                input_month_array(input_month_array_counter) = "August"
            Case 9
                master_ws = "SEP"
                input_month_array(input_month_array_counter) = "September"
            Case 10
                master_ws = "OCT"
                input_month_array(input_month_array_counter) = "October"
            Case 11
                master_ws = "NOV"
                input_month_array(input_month_array_counter) = "November"
            Case 12
                master_ws = "DEC"
                input_month_array(input_month_array_counter) = "December"
        End Select
    End If
End Sub

Public Sub extract_jobs(ByRef master_ws_row_counter As Integer, ByVal master_ws_last_row As Integer, ByVal master_wb As String, ByVal master_ws As String, ByRef master_ws_start_row_jobs As Integer)

    Dim master_ws_row_counter_1 As Integer
    Dim master_ws_row_counter_2 As Integer

    'Loop through column A to see where 'Job # is at for custom units
    For master_ws_row_counter_1 = master_ws_row_counter To master_ws_last_row
        If Workbooks(master_wb).Sheets(master_ws).Range("A" & master_ws_row_counter_1).Value = "Job #" Then
            Exit For
        End If
    Next master_ws_row_counter_1

    'Identify the number of custom jobs
    master_ws_row_counter_2 = master_ws_row_counter_1 + 1
    master_ws_start_row_jobs = master_ws_row_counter_1 + 1

    Do While Workbooks(master_wb).Sheets(master_ws).Range("A" & master_ws_row_counter_2).Value <> ""
        master_ws_row_counter = master_ws_row_counter_2
        master_ws_row_counter_2 = master_ws_row_counter_2 + 1
    Loop

End Sub

