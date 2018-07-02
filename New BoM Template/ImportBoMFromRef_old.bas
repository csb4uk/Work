Attribute VB_Name = "ImportBoMFromRef_old"
Option Explicit

Public Sub import_BoM_from_ref_job_old()

    Dim rb_name As String
    Dim rs_name As String
    Dim sb_name As String
    Dim ss_name As String

    Dim rb_start_row As Integer
    Dim rb_last_row As Integer
    Dim BoM_ID_counter As Integer
    Dim sb_row_counter As Integer
    Dim sb_start_row As Integer

    Dim ref_BoM_IDs As Variant

    Application.ScreenUpdating = False
    
    '===============================================================================================================================================
    'Store the current workbook and worksheet as variables.  The program will extract IDs from another excel sheet so this one will lose focus
    'when that happens.  We need these variables to call the program back to the current workbook
    '===============================================================================================================================================
    sb_name = ActiveWorkbook.Name
    ss_name = ActiveSheet.Name
    
    '===============================================================================================================================================
    'The start row, or import row, is identified by the active cell that the user is in.  This is where the IDs from the BoM to be imported will
    'come in.
    '===============================================================================================================================================
    sb_start_row = Selection.Rows(1).Row
    sb_row_counter = Selection.Rows(1).Row

    '===============================================================================================================================================
    'Find the workbook to import based on what the user types in
    '===============================================================================================================================================
    rb_name = find_rb_wb(rs_name)
    If rb_name = "" Then
        MsgBox ("BoM not found")
        End
    End If

    '===============================================================================================================================================
    'Once the workbook BoM is found get all of the Items, Quantities, and ID Numbers from the BoM to be imported
    '===============================================================================================================================================
    ref_BoM_IDs = get_BoM_IDs(rs_name, rb_start_row, rb_last_row)

    '===============================================================================================================================================
    'Activate the original workbook to import the item numbers, quantities, and id numbers
    '===============================================================================================================================================
    Workbooks(sb_name).Activate

    '===============================================================================================================================================
    'Insert rows in to the worksheet that matches the number of items to be imported.  This keeps the same format and the note about the spare
    'parts on the bottom of the worksheet
    '===============================================================================================================================================
    insert_rows sb_start_row, UBound(ref_BoM_IDs, 1)

    '===============================================================================================================================================
    'Populate the current BoM with all of the items, quantities, and ID numbers retrieved from the imported BoM
    '===============================================================================================================================================
    populate_BoM ref_BoM_IDs, sb_start_row, sb_row_counter
    
    Application.ScreenUpdating = True
    
End Sub

Function find_rb_wb(ByRef rs_name As String)
    Dim rb_BoM As String
    Dim wkb As Workbook
    Dim ans As Integer
    
    '===============================================================================================================================================
    'Prompt the user to enter the name of the BoM that you wish to import.  The for loop takes the user input and loops through all the Workbooks
    'that are currently opened in the same window as the BoM you are trying to write to.  It compares the input string to find the closest match.
    'Below is an example list of open workbooks, user input, and the BoM that is found
        '        Open Workbooks        |        User Input        |        BoM to Import
        '     MCB-RF00M, ZP-RF26M      |            ZP            |          ZP-RF26M
        '      ZP-RF25M, ZP-RF26M      |            ZP            |          ZP-RF25M   *Program defaults to the first match it finds
        '      ZP-RF25M, ZP-RF26M      |          ZP-RF26M        |          ZP-RF26M
    'If no match is found ask if the user would like to try again in the event of a typo
    '===============================================================================================================================================
TryAgain:
    find_rb_wb = ""
    rb_BoM = InputBox("Enter the Excel BoM you wish to import")
    For Each wkb In Workbooks
        If wkb.Name Like rb_BoM & "*" Then
            find_rb_wb = wkb.Name
            wkb.Activate
            rs_name = ActiveSheet.Name
            Exit For
        End If
    Next wkb
    If find_rb_wb = "" Then
        ans = MsgBox("BoM not found, would you like to try again?", vbYesNo)
        If ans = vbYes Then
            GoTo TryAgain
        End If
    End If
End Function

Function get_BoM_IDs(ByVal rs_name As String, ByRef rb_start_row As Integer, ByRef rb_last_row As Integer)

    Dim row_counter As Integer
    
    '===============================================================================================================================================
    'ID the last row of the reference book
    '===============================================================================================================================================
    
    With Worksheets(rs_name)
        rb_last_row = .Cells(.Rows.Count, 1).End(xlUp).Row
    End With

    '===============================================================================================================================================
    'Loop through each row in the BoM to be imported.
        'Looping through column B find a starting row by any of the following criteria:
            'If the value of the cell is an ID Number (IsNumeric)
            'If the value of the cell is "N/A" in the event the compressor is part of the condensing unit and shows up as "N/A"
            'If the value of the cell's first 2 characters are "P-" in the event the compressor was special ordered
            'And the value of the cell cannot be empty
    '===============================================================================================================================================
    For row_counter = 1 To rb_last_row
        If ((IsNumeric(Worksheets(rs_name).Range("B" & row_counter).Value) = True _
                Or UCase(Worksheets(rs_name).Range("B" & row_counter).Value) = "N/A" _
                Or UCase(Mid(Worksheets(rs_name).Range("B" & row_counter).Value, 1, 2)) = "P-") _
                And Worksheets(rs_name).Range("B" & row_counter).Value <> Empty) Then
            rb_start_row = row_counter
            Exit For
        End If
    Next

    '===============================================================================================================================================
    'Once the start row and last row is indentified store the arrange in to the get_BoM_IDs array
    '===============================================================================================================================================
    get_BoM_IDs = Worksheets(rs_name).Range("A" & rb_start_row & ":G" & rb_last_row).Value
End Function
Private Sub insert_rows(sb_start_row, last_insert_row)
    '===============================================================================================================================================
    'Take the amount of items in the array and add that many rows to the spreadsheet
    '===============================================================================================================================================
    Range("A" & sb_start_row).Rows("1:1").EntireRow.Copy
    last_insert_row = last_insert_row + sb_start_row
    Range("A" & sb_start_row & ":A" & last_insert_row).EntireRow.Insert Shift:=xlDown
    Application.CutCopyMode = False
End Sub
Private Sub populate_BoM(ref_BoM_IDs, sb_start_row, sb_row_counter)
    Dim BoM_ID_counter As Integer
    Dim str_counter As Integer
    Dim last_alpha_num As Integer

    Dim id_no As String
    Dim item_number As String

    Application.EnableEvents = False

    '===============================================================================================================================================
    'Loop through all of the array elements in ref_BoM_IDs to import the items, quantities, and ID numbers into the new BoM
    '===============================================================================================================================================
    For BoM_ID_counter = LBound(ref_BoM_IDs, 1) To UBound(ref_BoM_IDs, 1)
    
        '===============================================================================================================================================
        'The program will run through the ID numbers in the array and make suggestions for subcomponents of ID numbers. For example, if you have
        '00901 (Item 50), it has subcomponents such as 02529 and 00572.  It will suggest that they be listed as Items 50A and 50B.  If it is the
        'first ID number it will not check to see if it is a subcomponents of the previous part
        '===============================================================================================================================================
        If BoM_ID_counter = LBound(ref_BoM_IDs, 1) Then                                         'If it is the first iteration
            Range("A" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 1)                  'Copy the Range
        Else
            '===============================================================================================================================================
            'Logic is as followed:
                'Check to see if the ID number already has an item number.  This is done by looking at the first character of the value that is stored in
                    'the first element of the array.  If the first character is numeric then it is either identified as a subcomponent already, or it
                    'is its own item.  Either way is should be written to column A and not modified.
            '===============================================================================================================================================
            
            Select Case IsNumeric(Mid(ref_BoM_IDs(BoM_ID_counter, 1), 1, 1))                    'Check if the first character is a number
                Case True                                                                       'If it is a number
                    Range("A" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 1)          'Copy the value to the range
                Case False                                                                      'If it is not a number

                    '=======================================================================================================================================
                    'If the first item is not numeric, check to see if there is an ID number.  If the ID number is blank then there is not a value that needs
                    'to be assigned
                    '===============================================================================================================================================

                    Select Case ref_BoM_IDs(BoM_ID_counter, 2) = ""                             'Check if there is an ID number
                        Case True                                                               'If there is not an ID number
                            Range("A" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 1)  'Store the value
                        Case False

                            '=======================================================================================================================================
                            'If there is an ID number and the item number is not numberic, evaluate whether or not it is a humidity component.  If it is
                                'then it is a valid item number and can be stored in the A column
                            '=======================================================================================================================================
                            Select Case Mid(ref_BoM_IDs(BoM_ID_counter, 1), 1, 1)
                                Case "H"
                                    Range("A" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 1)  'Store the value
                                Case Else

                                    '===================================================================================================================
                                    'If the item is not a humidity component check to see if it is a 92XXX number.  92XXX cannot be subcomponents,
                                        'so we can assign an empty item number to column A
                                    '===================================================================================================================
                                    Select Case Mid(ref_BoM_IDs(BoM_ID_counter, 2), 1, 2)
                                        Case "92"
                                            Range("A" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 1)
                                        '===================================================================================================================
                                        'If there is any other scenario assume it is a subcomponent.  Step back to the previous cell and see if there
                                        'is already a letter listed as the subcomponent.  If there is not default to the letter A
                                        '===================================================================================================================
                                        Case Else
                                            item_number = id_last_item(sb_row_counter, sb_start_row)            'Step backward from the current row to the start row and suggest the last used item number plus a letter
                                            Select Case IsAlpha(item_number, str_counter, last_alpha_num)       'Find if there is already a letter used
                                                Case True
                                                    item_number = Mid(item_number, 1, str_counter - 1) & Chr(Asc(Mid(item_number, str_counter, 1)) + 1) & Mid(item_number, str_counter + 1)
                                                Case False
                                                    Select Case last_alpha_num
                                                        Case Len(item_number)
                                                            item_number = item_number & "A"
                                                        Case Else
                                                            item_number = Mid(item_number, 1, last_alpha_num) & "A" & Mid(item_number, last_alpha_num + 1)
                                                    End Select
                                            End Select
                                            Range("A" & sb_row_counter).Value = item_number
                                            highlight_cells sb_row_counter, Range("A" & sb_row_counter).Column
                                    End Select
                            End Select
                    End Select
            End Select
        End If

        '=======================================================================================================================================
        'If there is an ID number, assign the ID number from the reference BoM to a variable.  In the past the BoM IDs have been formatted as
        'Zip Codes for some reason.  As a result when they are directly extracted/copied the number will be shown without leading 0s.
        'For example 00038 only comes in as 38.  The following code will make sure the length of the ID number is 5, and will add 0s to the
        'front of the number until it gets there.
        '=======================================================================================================================================
        id_no = Trim(ref_BoM_IDs(BoM_ID_counter, 2))
        If id_no <> "" And IsNumeric(Mid(id_no, 1, 1)) = True Then
            Do While Len(id_no) < 5
                id_no = "0" & id_no
            Loop

        '=======================================================================================================================================
        'If there is not an ID number, it is either "N/A" or a "P-", or it is a section heading such as "Upper Refrigeration".  Below is the
        'code for something that has a value in the ID number column. Since these will not have ID numbers that have changed in our system
        'the program will bring in what has been typed on the previous/reference BoM and highlight the cells yellow.
        '=======================================================================================================================================

        ElseIf IsNumeric(Mid(id_no, 1, 1)) = False And Mid(id_no, 1, 1) <> "" Then
            id_no = UCase(id_no)
            Range("F" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 4)
            highlight_cells sb_row_counter, Range("F" & sb_row_counter).Column
            
            Range("I" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 5)
            highlight_cells sb_row_counter, Range("I" & sb_row_counter).Column
            
            Range("N" & sb_row_counter).NumberFormat = "@"
            Range("N" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 6)
            highlight_cells sb_row_counter, Range("N" & sb_row_counter).Column
            
        '=======================================================================================================================================
        'If it is a section heading (ID number column is blank on the reference BoM), bring in the section heading and bold, Underline, and
        'highlight the cells yellow
        '=======================================================================================================================================

        ElseIf id_no = "" Then
            If ref_BoM_IDs(BoM_ID_counter, 4) <> "" Then
                Range("F" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 4)
                highlight_cells sb_row_counter, Range("F" & sb_row_counter).Column
                    With Range("F" & sb_row_counter).Font
                        .Bold = True
                        .Underline = xlUnderlineStyleSingle
                    End With
            End If
        End If
        Range("B" & sb_row_counter).Value = id_no
        Range("D" & sb_row_counter).Value = ref_BoM_IDs(BoM_ID_counter, 3)
        sb_row_counter = sb_row_counter + 1
    Next BoM_ID_counter
    Application.EnableEvents = True
End Sub

Private Function IsAlpha(item_number, str_counter, last_alpha_num)

    Dim start_char As Integer

    '=======================================================================================================================================
    'This is used to identify subcomponents.  It first checks to see if the item number is a humidity component to identify a start character.
    'After the start character is identified find the last subcomponent number if one exists and add one.  Example: if 50A was the last used
    'subcomponent this will increment that by one to make the next one 50B.
    '=======================================================================================================================================

    Select Case Mid(item_number, 1, 1)
        Case "H"
            start_char = 2
        Case Else
            start_char = 1
    End Select
    For str_counter = start_char To Len(item_number)
        Select Case Asc(Mid(item_number, str_counter, 1))
            Case 65 To 90, 97 To 122
                IsAlpha = True
                Exit For
            Case 48 To 57
                last_alpha_num = str_counter
            Case Else
                IsAlpha = False
        End Select
    Next str_counter
End Function
Private Function id_last_item(sb_row_counter, sb_start_row)
    Dim step_counter As Integer

    '=======================================================================================================================================
    'This will step back to the previous row of the workbook that is being written to.  It will check to find the last item number so
    'the subcomponents will be able to be incremented. Example: If the code is writing to row 49, it will look back to row 48 to get the
    'item number that was last written.
    '=======================================================================================================================================
    step_counter = sb_row_counter - 1
    If IsNumeric(Mid(Range("A" & step_counter).Value, 1, 1)) = True Or IsNumeric(Mid(Range("A" & step_counter).Value, 2, 1)) = True Then
        id_last_item = Range("A" & step_counter).Value
    End If
End Function
Private Sub highlight_cells(sb_row_counter, column_number)

    '=======================================================================================================================================
    'highlights the cell of the current columnd and row.
    '=======================================================================================================================================

    With Cells(sb_row_counter, column_number).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
End Sub


