Attribute VB_Name = "PopulateWhiteboard"
Option Explicit

Public Sub populate_whiteboard()

    Dim start_row As Integer
    Dim last_row As Integer
    Dim row_count As Integer
    Dim length_serial_number As Integer
    Dim number_of_jobs As Integer
    Dim string_counter As Integer
    Dim number_of_letters As Integer

    Dim cab_hours As Double
    Dim electrical_hours As Double
    Dim refrigeration_hours As Double

    Dim job_number_array() As Variant

    Dim source_book As String
    Dim source_sheet As String
    Dim serial_number As String
    Dim year_job_number As String
    Dim job_number As String
    Dim file_path As String
    Dim budget_file_name As String
    Dim budget_file_path As String
    Dim customer_name As String
    Dim model_number As String
    Dim job_type As String
    Dim stability_file_path As String
    Dim tc_file_path As String
    Dim ans_msg_hours As String
    Dim vts_file_path As String


    Application.ScreenUpdating = False

    'Assign the Workbook and Sheet we want to write to as variables
    source_book = ActiveWorkbook.Name
    source_sheet = ActiveSheet.Name

    'Identify the start row and the last row of the selection box
    start_row = Selection.Rows(1).Row
    last_row = Selection.Rows.Count + start_row - 1

    'Loop through all the rows of the selected cells
    For row_count = start_row To last_row
        Application.ScreenUpdating = False
        serial_number = ActiveSheet.Range("A" & row_count).Value    'Identify the serial_number
        'Identify the type of job
        If IsNumeric(Left(serial_number, 1)) = True Then    'If this is a custom job
            year_job_number = "20" & Left(serial_number, InStr(1, serial_number, "-") - 1)  'Extract the year from the left 2 digits of the serial number
            
            For string_counter = InStr(1, serial_number, "-") To Len(serial_number)
                If IsNumeric(Mid(serial_number, string_counter,1)) = False Then
                    number_of_letters = number_of_letters + 1
                Else
                    Exit For
            Next string_counter
            Call NumberOfJobs.return_number_of_jobs_custom(serial_number, job_number, number_of_jobs, job_number_array, number_of_letters)
            file_path = "P:\Jobs\CUSTOM" & "\" & year_job_number & "\" & job_number & "\"        'Identify the file path based on all the information extracted from the serial_number.  It is in the form "P:\Jobs\CUSTOM\[year]\[job number]\"
        Else
            'If the job is standard or service it will go here
            For string_counter = 1 To Len(serial_number)
                If IsNumeric(Mid(serial_number, string_counter, 1)) = False Then
                    job_type = job_type & Mid(serial_number, string_counter, 1)
                Else
                    Exit For
                End If
            Next string_counter

            If InStr(1, serial_number, "-") > 0 Then  'There are multiple jobs
                Call NumberOfJobs.return_number_of_jobs_std(serial_number, job_number, job_type, number_of_jobs, job_number_array)
            Else
                number_of_jobs = 1
                job_number = serial_number
            End If

            year_job_number = "20" & Left(Mid(job_number, Len(job_type) + 1), 2)

            Select Case job_type
                Case "IND"
                    file_path = "P:\Jobs\IND JOBS\" & job_number & "\"
                Case "TC"
                    tc_file_path = "P:\Jobs\TC JOBS\" & year_job_number & "\"
                    Call TraversePath.TraversePath_STD_file_path(tc_file_path, file_path, job_number)
                Case "ZP"
                    file_path = "P:\Jobs\Z-PLUS JOBS" & "\" & year_job_number & "\" & job_number & "\"
                Case "ST"
                    stability_file_path = "P:\Jobs\ST JOBS\" & year_job_number & "\"
                    Call TraversePath.TraversePath_STD_file_path(stability_file_path, file_path, job_number)
                Case "MC"
                    file_path = "P:\Jobs\MC JOBS\MC-" & Right(year_job_number, 2) & "\" & job_number & "\"
                Case "SRV"
                    file_path = "P:\Jobs\IND JOBS\" & job_number & "\"
                Case "VT"
                    vts_file_path = "P:\Jobs\VT JOBS\" & year_job_number & "\"
                    Call TraversePath.TraversePath_STD_file_path(vts_file_path, file_path, job_number)
            End Select
        End If
        If file_path = "" Then
            MsgBox ("File not found, skip to next entry")
            GoTo NextBreak
        End If
        Call TraversePath.TraversePath(file_path, budget_file_name, budget_file_path, job_number)   'Call a program that will Identify all the files located in the file path folder
        If budget_file_name <> "" Then
            Call ExtractBudgetHours.extract_budget_hours(budget_file_path, budget_file_name, job_number, customer_name, model_number, cab_hours, electrical_hours, refrigeration_hours)      'Call a program that will extract all of the information in the budget
            'Write all of the data extracted into the Whiteboard Sheet
            If number_of_jobs > 1 Then
                ans_msg_hours = MsgBox("Would you like to multiply the budget by " & number_of_jobs & " since there are " & _
                                        number_of_jobs & " jobs attached to " & serial_number & "?" & vbCrLf & vbCrLf & _
                                        "Currently the hours are as followed:" & vbCrLf & vbCrLf & _
                                        "Cab Hours: " & cab_hours & vbCrLf & _
                                        "Electrical Hours: " & electrical_hours & vbCrLf & _
                                        "Refrigeration Hours: " & refrigeration_hours & vbCrLf, vbYesNo)
                If ans_msg_hours = vbYes Then
                    Workbooks(source_book).Sheets(source_sheet).Range("B" & row_count).Value = customer_name
                    Workbooks(source_book).Sheets(source_sheet).Range("C" & row_count).Value = model_number
                    If cab_hours = 0 Then
                        Workbooks(source_book).Sheets(source_sheet).Range("H" & row_count).Value = ""
                    Else
                        Workbooks(source_book).Sheets(source_sheet).Range("H" & row_count).Value = cab_hours * number_of_jobs
                    End If
                    If electrical_hours = 0 Then
                        Workbooks(source_book).Sheets(source_sheet).Range("K" & row_count).Value = ""
                    Else
                        Workbooks(source_book).Sheets(source_sheet).Range("K" & row_count).Value = electrical_hours * number_of_jobs
                    End If
                    If refrigeration_hours = 0 Then
                        Workbooks(source_book).Sheets(source_sheet).Range("N" & row_count).Value = ""
                    Else
                        Workbooks(source_book).Sheets(source_sheet).Range("N" & row_count).Value = refrigeration_hours * number_of_jobs
                    End If
                Else
                    Workbooks(source_book).Sheets(source_sheet).Range("B" & row_count).Value = customer_name
                    Workbooks(source_book).Sheets(source_sheet).Range("C" & row_count).Value = model_number
                    If cab_hours = 0 Then
                        Workbooks(source_book).Sheets(source_sheet).Range("H" & row_count).Value = ""
                    Else
                        Workbooks(source_book).Sheets(source_sheet).Range("H" & row_count).Value = cab_hours
                    End If
                    If electrical_hours = 0 Then
                        Workbooks(source_book).Sheets(source_sheet).Range("K" & row_count).Value = ""
                    Else
                        Workbooks(source_book).Sheets(source_sheet).Range("K" & row_count).Value = electrical_hours
                    End If
                    If refrigeration_hours = 0 Then
                        Workbooks(source_book).Sheets(source_sheet).Range("N" & row_count).Value = ""
                    Else
                        Workbooks(source_book).Sheets(source_sheet).Range("N" & row_count).Value = refrigeration_hours
                    End If
                End If
            Else
                Workbooks(source_book).Sheets(source_sheet).Range("B" & row_count).Value = customer_name
                Workbooks(source_book).Sheets(source_sheet).Range("C" & row_count).Value = model_number
                If cab_hours = 0 Then
                    Workbooks(source_book).Sheets(source_sheet).Range("H" & row_count).Value = ""
                Else
                    Workbooks(source_book).Sheets(source_sheet).Range("H" & row_count).Value = cab_hours
                End If
                
                If electrical_hours = 0 Then
                    Workbooks(source_book).Sheets(source_sheet).Range("K" & row_count).Value = ""
                Else
                    Workbooks(source_book).Sheets(source_sheet).Range("K" & row_count).Value = electrical_hours
                End If
                
                If refrigeration_hours = 0 Then
                    Workbooks(source_book).Sheets(source_sheet).Range("N" & row_count).Value = ""
                Else
                    Workbooks(source_book).Sheets(source_sheet).Range("N" & row_count).Value = refrigeration_hours
                End If
            End If
        Else
            MsgBox ("No budget found for " & serial_number)
        End If
NextBreak:
        Call ClearVariables.clear_vars(job_number_array, serial_number, year_job_number, job_number, customer_name, _
                                        model_number, budget_file_path, budget_file_name, file_path, number_of_jobs, _
                                        cab_hours, electrical_hours, refrigeration_hours, job_type)
        Application.ScreenUpdating = True
    Next row_count
Application.ScreenUpdating = True
End Sub
