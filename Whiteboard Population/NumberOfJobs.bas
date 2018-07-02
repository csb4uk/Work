Attribute VB_Name = "NumberOfJobs"
Option Private Module
Public Sub return_number_of_jobs_custom(serial_number, job_number, number_of_jobs, job_number_array, ByVal number_of_letters As Integer)

    Dim job_counter As Integer

    If Len(Mid(serial_number, InStr(1, serial_number, "-") + number_of_letters)) > 5 Then           'If the length of the job number is greater than 5 it indicates there are multiple jobs under the one entry
        job_number = Mid(serial_number, InStr(1, serial_number, "-") + number_of_letters, 5)        'Identify the job number by using the 5 numbers located 2 spaces to the right of the "-" in the serial number
        If Right(serial_number, 2) > Right(job_number, 2) Then                      'If the last two digits of the serial number are greater than the last two digits of the job number
            number_of_jobs = Right(serial_number, 2) - Right(job_number, 2) + 1         'Number of jobs is equal to the last two numbers of the serial number minus the last two numbers of the job number
        ElseIf Right(serial_number, 2) < Right(job_number, 2) Then                  'If the last two digits of the serial number are greater than the last two digits of the job number
            number_of_jobs = Right(serial_number, 2) + 100 - Right(job_number, 2)   'Number of jobs is equal to the last two numbers of the serial number plus 100 minus the last two numbers of the job number'
        End If

        'Create an array of the job numbers to be used later
        For job_counter = 0 To number_of_jobs - 1
            ReDim Preserve job_number_array(job_counter)
            job_number_array(job_counter) = CInt(job_number) + job_counter
        Next job_counter

    Else
        job_number = Mid(serial_number, InStr(1, serial_number, "-") + number_of_letters)        'If there are not multiple job numbers, just assign the value here
        number_of_jobs = 1
    End If
End Sub

Public Sub return_number_of_jobs_std(serial_number, job_number, job_type, number_of_jobs, job_number_array)

    Dim job_counter As Integer
    Dim dummy1 As String

        If InStr(1, serial_number, "-") > 0 Then   'If there are multiple jobs
                job_number = Mid(serial_number, 1, Len(job_type) + 7)        'Identify the job number by using the 5 numbers located 2 spaces to the right of the "-" in the serial number
                If Right(serial_number, 2) > Right(job_number, 2) Then                      'If the last two digits of the serial number are greater than the last two digits of the job number
            number_of_jobs = Right(serial_number, 2) - Right(job_number, 2) + 1         'Number of jobs is equal to the last two numbers of the serial number minus the last two numbers of the job number
        ElseIf Right(serial_number, 2) < Right(job_number, 2) Then                  'If the last two digits of the serial number are greater than the last two digits of the job number
            number_of_jobs = Right(serial_number, 2) + 100 - Right(job_number, 2)   'Number of jobs is equal to the last two numbers of the serial number plus 100 minus the last two numbers of the job number'
        End If
        'Create an array of the job numbers to be used later
        For job_counter = 0 To number_of_jobs - 1
            ReDim Preserve job_number_array(job_counter)
            job_number_array(job_counter) = Right(job_number, 7) + job_counter
        Next job_counter

    Else
        job_number = Mid(serial_number, InStr(1, serial_number, "-") + 3)        'If there are not multiple job numbers, just assign the value here
        number_of_jobs = 1
    End If


End Sub
