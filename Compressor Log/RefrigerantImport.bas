Attribute VB_Name = "RefrigerantImport"
Option Explicit


Sub Refrigerant_import()
Dim source_book, comp, comp_type, file_name, key_1, key_2, key_3, first_num, second_num As String
Dim start_cell, num_rows, adobe_rows, j, i, c_count, m As Integer
Dim arr1(), arr2, a_val As Variant
ReDim Preserve arr1(0 To 8, 0)

'Turn stuff off so the program can run faster
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False

source_book = ActiveWorkbook.Name    'So that if the workbook name is changed the code will still work
start_cell = ActiveCell.Row  'Starting location to add compatible refrigerants
Sheets("Sheet1").Select
num_rows = Workbooks("Compressor log").Sheets("Sheet1").Cells(Rows.Count, 2).End(xlUp).Row 'Finds the number of rows used in the doc.

'If you only want to do one iteration uncomment the following line
num_rows = start_cell

Do While start_cell < num_rows + 1  'While all the cells have not been evaluated.
    'Delete previously imported information from Adobe
    Sheets("Sheet2").Range("A1:Y300").EntireColumn.Delete
    
    'Check to see if column O has already had the refrigerants put in.  If it has not then execute the code.
    'This saves time if you are doing a full loop update rather than individually updating the rows
    Sheets("Sheet1").Select
    If Workbooks(source_book).Sheets("Sheet1").Range("O" & start_cell).Value = "" Then
    
        'Assign R-404A to column O since that is what all the sheets are rated at
        Workbooks(source_book).Sheets("Sheet1").Range("O" & start_cell).Value = "R-404A"
        comp = Workbooks(source_book).Sheets("Sheet1").Range("B" & start_cell) 'Identify the compressor
        comp_type = Workbooks(source_book).Sheets("Sheet1").Range("A" & start_cell) 'Identify the compressor
        
        'Use the file_name in hyperlink from column M.  Note that the file must not say "Summary".
        'It must be the actual file name with no shortened version or the code will fail here
        file_name = Evaluate("M" & start_cell).Value
        
        ThisWorkbook.FollowHyperlink (file_name)    'Open hyperlink of sheet. You will get a sheet error if you do not remove "Summary".  It needs to see the whole link.
        
        Call extract_adobe

        j = 0   'Counter
        adobe_rows = 300    'Arbitrary number of rows to look through

        For i = 1 To adobe_rows    'Go through all rows on Sheet2 where the adobe info is pasted
            a_val = Range("A" & i).Value    'All pasted info is in column A so assign that to a value
            If Left(a_val, 2) = "R-" Then   'Look through column A until it finds the string "R-" which signals a refrigerant
                arr2 = Split(a_val, " ")    'Split everything in that column into its own value in an array so you can find elements.  The space signals a new item
                If Application.WorksheetFunction.CountA(arr2) > 5 Then  'Make sure that an application is specified in the summary sheet
                    
                    'If the application is suited for low temperature write all the elements to a larger array to be evaluated later
                    If arr2(5) = "Low" Or arr2(6) = "Low" And (comp_type = "Scroll" Or comp_type = "Semi-Hermetic") Then
                        Select Case Application.WorksheetFunction.CountA(arr2)
                            Case Is = 9
                                Call eval_9(arr1, arr2, j)
                            Case Is = 8
                                Call eval_8(arr1, arr2, j)
                            Case Is = 7
                                Call eval_7(arr1, arr2, j)
                        End Select
                        j = j + 1
                        ReDim Preserve arr1(8, j)
                    Else: comp_type = "Hermetic"
                        Select Case Application.WorksheetFunction.CountA(arr2)
                            Case Is = 9
                                Call eval_9(arr1, arr2, j)
                            Case Is = 8
                                Call eval_8(arr1, arr2, j)
                            Case Is = 7
                                Call eval_7(arr1, arr2, j)
                        End Select
                        j = j + 1
                        ReDim Preserve arr1(8, j)
                    End If
                End If
            End If
        Next i
        
        'If there are no matching conditions for low temp, go to the next row in the compressor log, otherwise execute the following code
        If Application.WorksheetFunction.CountA(arr1) > 0 Then
            Sheets("Sheet1").Select
            c_count = 25    'Acts as a column counter and initializes to paste in column Y since the other refrigerants have designated locations
            key_1 = Range("E" & start_cell).Text    'Store the voltage of the active row a variable
            key_2 = Range("F" & start_cell).Text    'Store the phase of the active row as a variable
            key_3 = Range("G" & start_cell).Text    'Store the Hz of the active row as a variable
            If InStr(1, key_1, "-") > 0 Then        'If the voltage has a "-" we need to change the "-" to a "/"
                first_num = Mid(key_1, 1, 3)        'Store the first three numbers
                second_num = Mid(key_1, 5, 3)       'Store the second three numbers
                key_1 = first_num & "/" & second_num    'Combine first three numbers with a "/" and second three numbers
            End If
            
            'Run through each element of the array
            For m = 0 To j - 1
                'Make sure the voltage, phase, and hz in the array matches the values obtained from the three keys above
                If arr1(2, m) = key_1 And arr1(3, m) = key_2 And arr1(4, m) = key_3 Then
                    'If the above statement is true check to see if the application is low temperature
                    
                    'If both of the above statements are true write the refrigerant to the appropriate column
                    Select Case arr1(0, m)
                        Case "R-404A"
                        Case "R-507"
                            Cells(start_cell, 16) = arr1(0, m)
                        Case "R-134a"
                            Cells(start_cell, 17) = arr1(0, m)
                        Case "R-22"
                            Cells(start_cell, 18) = arr1(0, m)
                        Case "R-448A"
                            Cells(start_cell, 19) = arr1(0, m)
                        Case "R-449A"
                            Cells(start_cell, 20) = arr1(0, m)
                        Case "R-407C"
                            Cells(start_cell, 21) = arr1(0, m)
                        Case "R-407A"
                            Cells(start_cell, 22) = arr1(0, m)
                        Case "R-407F"
                            Cells(start_cell, 23) = arr1(0, m)
                        Case "R-502"
                            Cells(start_cell, 24) = arr1(0, m)
                        Case Else
                            Cells(start_cell, c_count) = arr1(0, m)
                            c_count = c_count + 1
                    End Select
                    
                End If
            Next m
        End If
    End If
    start_cell = start_cell + 1
Loop
'Turn the things on that were previously turned off to improve performance
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub

Private Sub extract_adobe()
    Application.Wait Now + TimeSerial(0, 0, 1)  'Wait one second
    Application.SendKeys ("^a")     'Highlight all text
    Application.SendKeys ("^c")     'Copy all text
    Application.Wait Now + TimeSerial(0, 0, 1)  'Wait one second
    Application.SendKeys ("^q")     'Close Adobe Reader
    Application.Wait Now + TimeSerial(0, 0, 0.5)    'Wait 0.5 seconds

    AppActivate "Microsoft Excel"   'Activate excel again
    Sheets("Sheet2").Select 'Select sheet 2
    Sheets("Sheet2").Range("A1").Select 'Select Range to paste
    'Paste Adobe info to excel
    ActiveSheet.PasteSpecial Format:="Unicode Text", Link:=False, _
    DisplayAsIcon:=False, NoHTMLFormatting:=True
    Application.Wait Now + TimeSerial(0, 0, 1)  'Wait one second
End Sub

Private Sub eval_9(ByRef arr1 As Variant, ByRef arr2 As Variant, ByVal j As Integer)
Dim n As Integer
Dim arr2_len As Integer
    For n = 0 To 8
        If n = 2 Or n = 3 Or n = 4 Then
            arr2_len = Len(arr2(n))
            Select Case arr2_len
                Case 1
                    arr1(3, j) = arr2(n)
                Case 2
                    arr1(4, j) = arr2(n)
                Case Else
                    arr1(2, j) = arr2(n)
            End Select
        Else
            arr1(n, j) = arr2(n)
        End If
    Next n
End Sub
Private Sub eval_8(ByRef arr1 As Variant, ByRef arr2 As Variant, ByVal j As Integer)
Dim n As Integer
Dim arr2_len As Integer
    For n = 0 To 7
        If n = 2 Or n = 3 Or n = 4 Then
            arr2_len = Len(arr2(n))
            Select Case arr2_len
                Case 1
                    arr1(3, j) = arr2(n)
                Case 2
                    arr1(4, j) = arr2(n)
                Case Else
                    arr1(2, j) = arr2(n)
            End Select
        Else
            arr1(n, j) = arr2(n)
        End If
    Next n
End Sub
Private Sub eval_7(ByRef arr1 As Variant, ByRef arr2 As Variant, ByVal j As Integer)
Dim n As Integer
Dim arr2_len As Integer
    For n = 0 To 6
        If n = 2 Or n = 3 Or n = 4 Then
            arr2_len = Len(arr2(n))
            Select Case arr2_len
                Case 1
                    arr1(3, j) = arr2(n)
                Case 2
                    arr1(4, j) = arr2(n)
                Case Else
                    arr1(2, j) = arr2(n)
            End Select
        Else
            arr1(n, j) = arr2(n)
        End If
    Next n
End Sub
