Attribute VB_Name = "TraversePathSubfolder"
Option Private Module
Option Explicit
Option Base 1

Public Sub TraversePath_subfolder(ByRef file_path As String, ByRef name_of_budget As Variant, ByRef date_of_budget As Variant, _
                                        ByRef combined_info As Variant, ByRef number_of_budgets As Integer, ByRef budget_file_path_arr As Variant)
    Dim file_name As String
     
    file_name = Dir(file_path)   'file_name is the file name
     
    'Explore current directory
    Do Until file_name = vbNullString     'Search the folder for all files and subfolders
        If InStr(1, file_name, "Budget") > 0 And InStr(1, file_name, ".xl") Then
            number_of_budgets = number_of_budgets + 1
            ReDim Preserve name_of_budget(number_of_budgets)
            ReDim Preserve date_of_budget(number_of_budgets)
            ReDim Preserve combined_info(number_of_budgets)
            ReDim Preserve budget_file_path_arr(number_of_budgets)
            name_of_budget(number_of_budgets) = file_name
            date_of_budget(number_of_budgets) = FileDateTime(file_path & file_name)
            combined_info(number_of_budgets) = name_of_budget(number_of_budgets) & vbLf & date_of_budget(number_of_budgets)
            budget_file_path_arr(number_of_budgets) = file_path
        End If
        file_name = Dir()   'Go to the next file
    Loop
End Sub

Public Sub TraversePath_subfolder_STD(ByRef std_file_path As String, ByRef file_path As String, ByRef job_number As String)
    Dim file_name As String
    Dim directory As Variant
     
    file_name = Dir(std_file_path, vbDirectory)   'file_name is the file name
     
    'Explore current directory
    Do Until file_name = vbNullString     'Search the folder for all files and subfolders
        If file_name = job_number Then
            file_path = std_file_path & file_name & "\"
            Exit Sub
        End If
        file_name = Dir()   'Go to the next file
    Loop
End Sub

