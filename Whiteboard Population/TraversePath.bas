Attribute VB_Name = "TraversePath"
Option Private Module
Option Explicit
Option Base 1


Public Sub TraversePath(ByRef file_path As String, ByRef budget_file_name As String, ByRef budget_file_path As String, _
                            ByVal job_number As String)
    Dim file_name As String
    Dim msg As String
    Dim budget_names As String
    Dim ans_continue As String

    Dim directory As Variant
    Dim name_of_budget() As Variant
    Dim date_of_budget() As Variant
    Dim combined_info() As Variant
    Dim budget_file_path_arr() As Variant
    
    Dim number_of_budgets As Integer
    Dim ans_input_box As Integer
    Dim budget_counter As Integer

    Dim dirCollection As Collection     'Store the collection of directories as a Collection
    Set dirCollection = New Collection  'Make the directory Collection a new Collection
    
    file_name = Dir(file_path, vbDirectory)   'file_name is the file name
     
    'Explore current directory
    Do Until file_name = vbNullString     'Search the folder for all files and subfolders
        Debug.Print file_name             'Print out the name of the file in the immediate window
        If Left(file_name, 1) <> "." And (GetAttr(file_path & file_name) And vbDirectory) = vbDirectory Then     'If the file name is a subfolder then add it to the directory collection
            dirCollection.Add file_name
        End If
        If InStr(1, file_name, "Budget") > 0 And InStr(1, file_name, ".xlsm") Then
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
     
    'Explore subsequent directories
    For Each directory In dirCollection     'For each directory in the collection
        'Debug.Print "---SubDirectory: " & directory & "---"
        TraversePathSubfolder.TraversePath_subfolder file_path & directory & "\", name_of_budget, date_of_budget, _
                                                        combined_info, number_of_budgets, budget_file_path_arr
    Next directory
    If number_of_budgets > 1 Then
        msg = "Select which Budget file you would like to import for Job Number " & job_number & ":" & vbLf & vbLf
        For budget_counter = 1 To number_of_budgets
            budget_names = budget_names & budget_counter & ". " & combined_info(budget_counter) & vbLf & vbLf
        Next budget_counter
        'On Error GoTo ErrIB
SelectBudget:
        ans_input_box = InputBox(msg & budget_names)
        budget_file_name = name_of_budget(ans_input_box)
        budget_file_path = budget_file_path_arr(ans_input_box)
        Exit Sub
ErrIB:
        ans_continue = MsgBox("Would you like to skip this entry?", vbYesNo)
        If ans_continue = vbNo Then
            Resume SelectBudget
        Else
            Exit Sub
        End If
    ElseIf number_of_budgets = 1 Then
        budget_file_name = name_of_budget(1)
        budget_file_path = budget_file_path_arr(1)
        Exit Sub
    End If
End Sub

Public Sub TraversePath_STD_file_path(ByRef std_file_path As String, ByRef file_path As String, ByRef job_number As String)
    Dim file_name As String

    Dim directory As Variant

    Dim dirCollection As Collection     'Store the collection of directories as a Collection
    Set dirCollection = New Collection  'Make the directory Collection a new Collection
    
    file_name = Dir(std_file_path, vbDirectory)   'file_name is the file name
     
    'Explore current directory
    Do Until file_name = vbNullString     'Search the folder for all files and subfolders
        'Debug.Print file_name             'Print out the name of the file in the immediate window
        If Left(file_name, 1) <> "." And (GetAttr(std_file_path & file_name) And vbDirectory) = vbDirectory Then     'If the file name is a subfolder then add it to the directory collection
            dirCollection.Add file_name
        End If
        If file_name = job_number Then
            file_path = std_file_path & file_name & "\"
            Exit Sub
        End If
        file_name = Dir()   'Go to the next file
    Loop
     
    'Explore subsequent directories
    For Each directory In dirCollection     'For each directory in the collection
        'Debug.Print "---SubDirectory: " & directory & "---"
        TraversePathSubfolder.TraversePath_subfolder_STD std_file_path & directory & "\", file_path, job_number
    Next directory
End Sub
