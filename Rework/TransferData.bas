Attribute VB_Name = "TransferData"
Option Explicit
Option Base 1
Public Sub transfer_data()

    Dim start_row As Integer
    Dim last_row As Integer
    Dim current_row As Integer
    Dim counter_total_jobs As Integer

    Dim hours As Double

    Dim rework_date As Date

    Dim this_workbook As String
    Dim this_worksheet As String
    Dim job As String
    Dim discipline As String
    Dim cause As String

    Dim existing_job As Boolean

    Dim array_info As Variant
    Dim array_rework() As Variant

    'Initialize Variables
    start_row = Selection.Rows(1).Row
    last_row = Selection.Rows.Count + start_row - 1
    this_workbook = ActiveWorkbook.Name
    this_worksheet = ActiveSheet.Name
    array_info = Workbooks(this_workbook).Worksheets(this_worksheet).Range("A" & start_row & ":I" & last_row).Value
    counter_total_jobs = 0

    'Loop through all custom jobs
    For current_row = LBound(array_info, 1) To UBound(array_info, 1)
        'If Column B (job number) Has a Value then evaluate the job, if it does not assume it is a blank row and move to the next selection
        If array_info(current_row, 2) <> "" Then
            'Assign the job and discipline in Column B and F to collect all data attributed to that classification
            job = array_info(current_row, 2)
            discipline = array_info(current_row, 6)
            cause = array_info(current_row, 7)

            'Figure out if the job has already been evaluated if it has, skip to the next job
            loop_previous_jobs start_row, current_row, job, discipline, array_info, existing_job, cause
            
            'If the job has yet to be evaluated then continue evaluating it
            If existing_job = False Then
                'Increase the size of the rework array for the new job
                counter_total_jobs = counter_total_jobs + 1
                ReDim Preserve array_rework(7, 1 To counter_total_jobs)

                'Assign initial job values from the first match found
                rework_date = array_info(current_row, 1)
                hours = array_info(current_row, 8)

                'Combine the rest of the matches found in the selection
                combine_data current_row, last_row, array_info, job, discipline, rework_date, hours, cause

                'Format the data to export to the engineering rework document
                format_data_rework array_rework, counter_total_jobs, rework_date, job, discipline, hours, cause
            End If
        End If
        existing_job = False
    Next current_row
    
    'Sort rework array by date
    sort_rework_array array_rework

    'Send the data to the rework document
    send_data array_rework

End Sub

Private Sub loop_previous_jobs(ByVal start_row As Integer, ByVal current_row As Integer, ByVal job As String, ByVal discipline As String, ByRef array_info As Variant, ByRef existing_job As Boolean, ByVal cause)
    Dim counter As Integer

    'Find out if the job has already been evaluated
    For counter = LBound(array_info, 1) To current_row - 1
        If array_info(counter, 2) = job And array_info(counter, 6) = discipline And array_info(counter, 7) = cause Then
            existing_job = True
            Exit Sub
        End If
    Next counter
    existing_job = False
End Sub

Private Sub combine_data(ByVal current_row As Integer, ByVal last_row As Integer, ByRef array_info As Variant, ByVal job As String, ByVal discipline As String, ByRef rework_date As Date, _
    ByRef hours As Double, ByVal cause)

    Dim counter As Integer
    'If the job has not been evaluated loop through the rest of the rows to combine the hours for the job and the discipline
    For counter = current_row + 1 To UBound(array_info, 1)
        If array_info(counter, 2) = job And array_info(counter, 6) = discipline And array_info(counter, 7) = cause Then
            hours = hours + array_info(counter, 8)
                If DateValue(array_info(counter, 1)) < rework_date Then
                    rework_date = array_info(counter, 1)
                End If
        End If
    Next counter
End Sub

Private Sub format_data_rework(ByRef array_rework As Variant, ByVal counter_total_jobs As Integer, ByVal rework_date As Date, ByVal job As String, _
    ByVal discipline As String, ByVal hours As Double, ByVal cause)

        Dim counter As Integer
        Dim num_letters As Integer
    
        array_rework(1, counter_total_jobs) = rework_date
        array_rework(2, counter_total_jobs) = UCase(MonthName(Month(rework_date), True))
        array_rework(3, counter_total_jobs) = job

        For counter = InStr(1, job, "-") + 1 To Len(job)
            If IsNumeric(Mid(job, counter, 1)) = True Then
                array_rework(4, counter_total_jobs) = Mid(job, InStr(1, job, "-") + 1, num_letters)
                Exit For
            End If
            num_letters = num_letters + 1
        Next counter

        array_rework(5, counter_total_jobs) = hours

        Select Case discipline
            Case "R"
                array_rework(6, counter_total_jobs) = "REFRIGERATION"
            Case "E"
                array_rework(6, counter_total_jobs) = "ELECTRICAL"
            Case "C"
                array_rework(6, counter_total_jobs) = "CABINETRY"
            Case "T"
                array_rework(6, counter_total_jobs) = "RETEST"
            Case Else
                array_rework(6, counter_total_jobs) = "UNKNOWN"
        End Select
        Select Case cause
            Case "D"
                array_rework(7, counter_total_jobs) = "DESIGN ERROR"
            Case "U"
                array_rework(7, counter_total_jobs) = "UNKNOWN"
            Case "W"
                array_rework(7, counter_total_jobs) = "WHITESHEET"
            Case "T"
                array_rework(7, counter_total_jobs) = "N/A"
        End Select
End Sub

Private Sub sort_rework_array(ByRef array_rework As Variant)
    Dim current_date As Date

    Dim current_month As String
    Dim current_job As String
    Dim current_job_type As String
    Dim current_discipline As String
    Dim current_cause As String

    Dim last_day_of_month As Integer
    Dim previous_array_counter As Integer
    Dim current_array_counter As Integer

    Dim current_hours As Double

    Dim array_counter As Variant

    For current_array_counter = 2 To UBound(array_rework, 2)
        current_date = array_rework(1, current_array_counter)
        current_month = array_rework(2, current_array_counter)
        current_job = array_rework(3, current_array_counter)
        current_job_type = array_rework(4, current_array_counter)
        current_hours = array_rework(5, current_array_counter)
        current_discipline = array_rework(6, current_array_counter)
        current_cause = array_rework(7, current_array_counter)
        For previous_array_counter = current_array_counter - 1 To 1 Step -1
            If (array_rework(1, previous_array_counter) <= current_date) Then
                GoTo NextBreak
            Else
                For array_counter = LBound(array_rework, 1) To UBound(array_rework, 1)
                    array_rework(array_counter, previous_array_counter + 1) = array_rework(array_counter, previous_array_counter)
                Next array_counter
            End If
        Next previous_array_counter
        previous_array_counter = 0
NextBreak:
        array_rework(1, previous_array_counter + 1) = current_date
        array_rework(2, previous_array_counter + 1) = current_month
        array_rework(3, previous_array_counter + 1) = current_job
        array_rework(4, previous_array_counter + 1) = current_job_type
        array_rework(5, previous_array_counter + 1) = current_hours
        array_rework(6, previous_array_counter + 1) = current_discipline
        array_rework(7, previous_array_counter + 1) = current_cause
    Next current_array_counter
End Sub

Private Sub send_data(ByRef array_rework As Variant)
    Dim previous_date As Date

    Dim file_path_rework As String
    Dim file_name_rework As String
    Dim rework_wb As String
    Dim rework_ws As String
    Dim file_path_whiteboard As String
    Dim file_name_whiteboard As String
    Dim whiteboard_wb As String
    Dim whiteboard_ws As String


    Dim rework_year As Integer
    Dim last_row_rework_wb As Integer
    Dim import_row_rework_wb As Integer
    Dim rework_row_counter As Integer
    Dim whiteboard_last_row As Integer
    Dim whiteboard_row_counter As Integer
    Dim total_counter As Integer
    Dim total_jobs As Integer

    Dim job_match As Boolean

    Dim wkb As Workbook

    Dim monthly_hours As Double

    Dim whiteboard_array As Variant

    Application.ScreenUpdating = False
        'Extract data from Whiteboard
        file_path_whiteboard = "P:\WHITE BOARD\"
        file_name_whiteboard = "White board schedule, 2018.xlsx"
        Workbooks.Open _
            Filename:=file_path_whiteboard & file_name_whiteboard, _
            ReadOnly:=True
                whiteboard_wb = ActiveWorkbook.Name
                whiteboard_ws = "JOBS RELEASED"
                whiteboard_last_row = Workbooks(whiteboard_wb).Worksheets(whiteboard_ws).Range("A" & Rows.Count()).End(xlUp).Row
                whiteboard_array = Workbooks(whiteboard_wb).Worksheets(whiteboard_ws).Range("A2" & ":O" & whiteboard_last_row)
        Workbooks(whiteboard_wb).Close SaveChanges:=False

        rework_year = Year(array_rework(1, 1))

        'The following code has been omitted so Kathleen can import to a rework sheet that is already opened
            'Open the Engineering Rework File
            'file_path_rework = "P:\IND ENG INFO\REWORK-RFT\REWORK\" & rework_year & "\"
            'file_name_rework = "ENGREWORK " & rework_year & ".xlsx"
            'Workbooks.Open _
            '    Filename:=file_path_rework & file_name_rework
            'rework_wb = ActiveWorkbook.Name
            'rework_ws = "DATA SORT"
            'Workbooks(rework_wb).Worksheets(rework_ws).Activate

        'Find the rework workbook that is already opened
        For Each wkb In Workbooks
            If wkb.Name Like "ENGREWORK*" Then
                wkb.Activate
                rework_wb = ActiveWorkbook.Name
                rework_ws = "DATA SORT"
                Worksheets(rework_ws).Activate
                Exit For
            End If
        Next wkb


        'Find where to import the data
        last_row_rework_wb = Workbooks(rework_wb).Worksheets(rework_ws).Range("A" & Rows.Count()).End(xlUp).Row
        import_row_rework_wb = last_row_rework_wb

        'Import the data
        For rework_row_counter = LBound(array_rework, 2) To UBound(array_rework, 2)
            import_row_rework_wb = import_row_rework_wb + 1
            With Workbooks(rework_wb).Worksheets(rework_ws)
                If array_rework(2, rework_row_counter) <> .Range("B" & import_row_rework_wb - 1).Value And _
                    .Range("B" & import_row_rework_wb - 1).Value <> "MTH" Then
                    previous_date = .Range("A" & import_row_rework_wb - 1).Value
                    .Range("A" & import_row_rework_wb).Value = DateSerial(Year(previous_date), Month(previous_date) + 1, 0)
                    With .Range("E" & import_row_rework_wb)
                        .Value = "TOTAL:"
                        .HorizontalAlignment = xlRight
                    End With
                    For total_counter = 1 To import_row_rework_wb - 1
                        If .Range("B" & import_row_rework_wb - 1).Value = .Range("B" & total_counter).Value Then
                            monthly_hours = monthly_hours + .Range("F" & total_counter).Value
                        End If
                    Next total_counter
                    .Range("F" & import_row_rework_wb).Value = monthly_hours
                    With .Range("A" & import_row_rework_wb & ":J" & import_row_rework_wb).Interior
                        .Pattern = xlSolid
                        .PatternColorIndex = xlAutomatic
                        .ThemeColor = xlThemeColorLight1
                        .TintAndShade = 0
                        .PatternTintAndShade = 0
                    End With
                    With .Range("A" & import_row_rework_wb & ":J" & import_row_rework_wb).Font
                        .ThemeColor = xlThemeColorDark1
                        .TintAndShade = 0
                        .Bold = True
                    End With
                    import_row_rework_wb = import_row_rework_wb + 1
                End If
                .Range("A" & import_row_rework_wb).Value = array_rework(1, rework_row_counter)
                .Range("B" & import_row_rework_wb).Value = array_rework(2, rework_row_counter)
                .Range("C" & import_row_rework_wb).Value = array_rework(3, rework_row_counter)
                .Range("D" & import_row_rework_wb).Value = array_rework(4, rework_row_counter)
                .Range("F" & import_row_rework_wb).Value = array_rework(5, rework_row_counter)
                .Range("G" & import_row_rework_wb).Value = array_rework(6, rework_row_counter)
                .Range("I" & import_row_rework_wb).Value = array_rework(7, rework_row_counter)

                For whiteboard_row_counter = LBound(whiteboard_array, 1) To UBound(whiteboard_array, 1)
                    If whiteboard_array(whiteboard_row_counter, 1) <> "" Then
                        If whiteboard_array(whiteboard_row_counter, 1) = array_rework(3, rework_row_counter) Then
                            match_whiteboard_data rework_wb, rework_ws, import_row_rework_wb, whiteboard_array, whiteboard_row_counter, rework_row_counter, array_rework, job_match
                        Else
                            number_of_jobs total_jobs, whiteboard_array(whiteboard_row_counter, 1)
                            If total_jobs > 1 Then
                                evaluate_all_jobs total_jobs, whiteboard_array(whiteboard_row_counter, 1), array_rework(3, rework_row_counter), rework_wb, rework_ws, import_row_rework_wb, _
                                                    whiteboard_array, whiteboard_row_counter, rework_row_counter, array_rework, job_match
                            End If
                        End If
                        If job_match = True Then
                            Exit For
                        End If
                    End If
                Next whiteboard_row_counter
            End With
        job_match = False
        Next rework_row_counter
        Application.ScreenUpdating = True
End Sub

Private Sub match_whiteboard_data(ByVal rework_wb, ByVal rework_ws, ByVal import_row_rework_wb, ByRef whiteboard_array, _
    ByVal whiteboard_row_counter, ByVal rework_row_counter, ByRef array_rework, ByRef job_match)
    Workbooks(rework_wb).Worksheets(rework_ws).Range("E" & import_row_rework_wb).Value = whiteboard_array(whiteboard_row_counter, 2)
    With Workbooks(rework_wb).Worksheets(rework_ws)
        Select Case array_rework(6, rework_row_counter)
            Case "REFRIGERATION"
                .Range("H" & import_row_rework_wb).Value = UCase(whiteboard_array(whiteboard_row_counter, 15))
            Case "ELECTRICAL"
                .Range("H" & import_row_rework_wb).Value = UCase(whiteboard_array(whiteboard_row_counter, 12))
            Case "CABINETRY"
                .Range("H" & import_row_rework_wb).Value = UCase(whiteboard_array(whiteboard_row_counter, 9))
            Case "RETEST"
                .Range("H" & import_row_rework_wb).Value = "RETEST"
            Case Else
                .Range("H" & import_row_rework_wb).Value = "N/A"
        End Select
    End With
    job_match = True
End Sub
Private Sub number_of_jobs(ByRef total_jobs, ByVal serial_number)

    Dim character_counter As Integer
    Dim first_number As Integer
    Dim last_number As Integer
    Dim pos_1 As Variant
    Dim pos_2 As Variant
    Dim multi_jobs As Boolean
    Dim job_counter_coll As New Collection

    If Left(serial_number, 3) = "INT" Then
        serial_number = Mid(serial_number, 4)
    End If

    'Determine if there are multiple jobs
    pos_1 = Mid(serial_number, Len(serial_number) - 2, 1)
    pos_2 = Mid(serial_number, Len(serial_number) - 3, 1)
    If (IsNumeric(Left(serial_number, 1)) = True) And ((pos_1 = "-") Or (pos_2 = "-")) Then 'multiple jobs custom
        multi_jobs = True
    ElseIf (IsNumeric(Left(serial_number, 1)) = False) And (InStr(1, serial_number, "-") > 0) Then  'multiple jobs standard
        multi_jobs = True
    Else
        multi_jobs = False
        total_jobs = 1
    End If
    If multi_jobs = True Then
        For character_counter = Len(serial_number) To 1 Step -1
            If IsNumeric(Mid(serial_number, character_counter - 1, 1)) = True And IsNumeric(Mid(serial_number, character_counter, 1)) = True Then
                job_counter_coll.Add Mid(serial_number, character_counter - 1, 2)
                If job_counter_coll.Count = 2 Then
                    first_number = job_counter_coll(2)
                    last_number = job_counter_coll(1)
                    If first_number < last_number Then
                        total_jobs = (last_number - first_number) + 1
                    Else
                        total_jobs = ((last_number + 100) - first_number) + 1
                    End If
                    Exit For
                End If
            End If
        Next
    End If
End Sub
Private Sub evaluate_all_jobs(ByVal total_jobs, ByVal whiteboard_serial_number, ByVal rework_serial_number, ByVal rework_wb, ByVal rework_ws, ByVal import_row_rework_wb, ByRef whiteboard_array, _
    ByVal whiteboard_row_counter, ByVal rework_row_counter, ByRef array_rework, ByRef job_match)

    Dim counter As Integer
    Dim character_count As Integer
    Dim last_letter As String
    Dim first_job As String
    Dim current_job As String


    For character_count = Len(whiteboard_serial_number) To 1 Step -1
        If Mid(whiteboard_serial_number, character_count, 1) = "-" Then
            Select Case IsNumeric(Right(whiteboard_serial_number, 1))
                Case True   'Non-warranty job
                    last_letter = ""
                    first_job = Mid(whiteboard_serial_number, 1, character_count - 1)
                    Exit For
                Case False  'Warranty Job
                    last_letter = Right(whiteboard_serial_number, 1)
                    first_job = Mid(whiteboard_serial_number, 1, character_count - 2)
                    Exit For
            End Select
        End If
    Next

    For counter = 0 To total_jobs - 1
        current_job = Mid(first_job, 1, Len(first_job) - 3) & Mid(first_job, Len(first_job) - 2) + counter & last_letter
        If current_job = rework_serial_number Then
            match_whiteboard_data rework_wb, rework_ws, import_row_rework_wb, whiteboard_array, whiteboard_row_counter, rework_row_counter, array_rework, job_match
            Exit For
        End If
    Next


End Sub

