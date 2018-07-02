Attribute VB_Name = "ExportBoMVisual"
Option Explicit
Private Declare Function GetForegroundWindow Lib "user32" _
    () As Long
Private Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" (ByVal hwnd As Long, _
    ByVal lpString As String, ByVal cch As Long) As Long

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_NUMLOCK = &H90
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer


Public Sub ExportBoM()

Dim arr As Variant

Dim last_row As Integer
Dim start_row As Integer
Dim arr_counter As Integer

Dim Wintext As String
Dim source_book As String
Dim source_sheet As String
Dim apple As String

Dim excel_win As String

    'Set the active workbook and sheet to variables
    source_book = ActiveWorkbook.Name
    source_sheet = ActiveSheet.Name
    
    start_row = Selection.Rows(1).Row
    last_row = Selection.Rows.Count + start_row - 1
    
    'Generate the total number of each ID number into column L if the ID number has not already been accounted for
    Call generate_qty(start_row, last_row, source_book, source_sheet)
       
    Call convert2text(start_row, last_row)  'Convert the numbers in column b to text.

    
    arr = Range("A" & start_row & ":P" & last_row).Value    'put range of B:P into an array to speed up processing
    
    Call window_get(Wintext)    'Get the window name of the excel sheet to be able to use AppActivate later on
        excel_win = Wintext     'assign the text of the window to a variable
    
    AppActivate excel_win   'Activate the excel window
    
    arr = sort_id_array(arr)

    UserForm1.Height = 600
    UserForm1.Width = 750
    UserForm1.Left = 125
    UserForm1.Top = 25
    With UserForm1.ListBox1
        .Left = 12
        .Width = 726
        .Height = 500
        .Top = 18
        .ColumnCount = 7
        .ColumnWidths = "0.5 in;0.75 in;0.5 in;2.75 in;2.75 in;2 in;0.5 in"
    End With
    With UserForm1.lblDelete
        .Width = 72
        .Left = UserForm1.Width / 2 - (UserForm1.lblDelete.Width + 12)
        .Height = 24
        .Top = UserForm1.ListBox1.Top + UserForm1.ListBox1.Height + 18
    End With
    With UserForm1.lblContinue
        .Width = 72
        .Left = UserForm1.Width / 2 + 12
        .Height = 24
        .Top = UserForm1.lblDelete.Top
    End With
    With UserForm1.ListBox1
        On Error GoTo ErrCatcher
        For arr_counter = LBound(arr, 1) To UBound(arr, 1)
            .AddItem arr(arr_counter, 1)
            .List(.ListCount - 1, 1) = arr(arr_counter, 2)
            .List(.ListCount - 1, 2) = arr(arr_counter, 4)
            .List(.ListCount - 1, 3) = arr(arr_counter, 6)
            .List(.ListCount - 1, 4) = arr(arr_counter, 9)
            .List(.ListCount - 1, 5) = arr(arr_counter, 14)
            .List(.ListCount - 1, 6) = arr(arr_counter, 16)
        Next
    End With
    UserForm1.Show
    
ErrCatcher:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case -2147352571
                MsgBox ("Item " & arr(arr_counter, 2) & " is not found in the database.  Please add part or clear the error in the cell to continue")
                Exit Sub
            Case Else
                MsgBox (Err.Number & " refers to an error: " & vbCrLf & Err.Description & vbCrLf & "Please add to error handler")
        End Select
        Resume
    End If
End Sub

Private Function sort_id_array(ByRef start_arr)
    Dim array_row_counter As Integer
    Dim array_column_counter As Integer
    Dim temp_collection As Collection
    Dim previous_row_counter As Integer
    Dim arr_counter As Integer

    For array_row_counter = 2 To UBound(start_arr, 1)
        If IsNumeric(Mid(start_arr(array_row_counter, 2), 1, 1)) = False Then
            Set temp_collection = New Collection
            For array_column_counter = LBound(start_arr, 2) To UBound(start_arr, 2)
                temp_collection.Add start_arr(array_row_counter, array_column_counter)
            Next
            For previous_row_counter = array_row_counter - 1 To LBound(start_arr, 1) Step -1
                If IsNumeric(Mid(start_arr(previous_row_counter, 2), 1, 1)) = False Then
                    GoTo NextBreak
                Else
                    For arr_counter = LBound(start_arr, 2) To UBound(start_arr, 2)
                        start_arr(previous_row_counter + 1, arr_counter) = start_arr(previous_row_counter, arr_counter)
                    Next
                End If
            Next
            previous_row_counter = 0
NextBreak:
            For array_column_counter = LBound(start_arr, 2) To UBound(start_arr, 2)
                start_arr(previous_row_counter + 1, array_column_counter) = temp_collection(array_column_counter)
            Next
        End If
    Next
    sort_id_array = start_arr
End Function
Public Sub ex_vis()
    Dim arr As Variant
    Dim id_num As Variant

    Dim serial_num As String
    Dim Wintext As String
    Dim excel_win As String
    Dim vis_window As String
    Dim mat_window As String

    Dim j As Integer
    Dim i As Integer
    Dim k As Integer
    Dim a As Integer
    Dim qty As Integer
    Dim total_items As Integer
    Dim start_row As Integer
    Dim msg_ans_1 As Integer
    Dim msg_ans_2 As Integer

        total_items = 0
        arr = UserForm1.ListBox1.List
        ReDim Preserve arr(UBound(arr, 1), UserForm1.ListBox1.ColumnCount - 1)
EnterSerialNum:
        serial_num = InputBox("Please type in the serial number of the project")
        Call window_get(Wintext)    'Get the window name of the excel sheet to be able to use AppActivate later on
            excel_win = Wintext     'assign the text of the window to a variable

        vis_window = "Manufacturing Window - Infor VISUAL - CSZ - [" & serial_num & "/1]"   'Name of the visual sheet in order to use AppActivate

        AppActivate excel_win   'Activate the excel window

        MsgBox ("Please select the desired operation to populate the BoM in the visual window.") 'Double Check that it will import to the correct place before continuing
        j = 1
        'n = 1
            'go through each item in the array
            For k = 0 To UBound(arr, 1)
                id_num = arr(k, 1)  'Assign the ID number in column B to a variable
                qty = arr(k, 6)     'Assign the quantity that was generated in column G to a variable
                    If (id_num = "" Or qty = 0) Then
                        GoTo NextBreak
                    Else
                        If j = 1 Then   'On the first instance of the export, the visual window will need to be activated.  After that the focus will stay there
                            Call act_vis(serial_num, j, vis_window)
                            If j = 1 Then
                                msg_ans_1 = MsgBox("Would you like to try a different serial number?", vbYesNo)
                                If msg_ans_1 = vbYes Then
                                    GoTo EnterSerialNum
                                Else
                                    msg_ans_2 = MsgBox("Would you like to edit the items to be imported?", vbYesNo)
                                    If msg_ans_2 = vbYes Then
                                        UserForm1.Show
                                    Else
                                        End
                                    End If
                                End If
                            End If
                        End If
                        Call window_get(Wintext)
                        mat_window = Wintext
                        Call pop_vis(id_num, qty, mat_window, Wintext, excel_win) 'populate visual and double check that the Manufacturing number is a match
                        total_items = total_items + 1
                        'n = n + 1
                    End If
NextBreak:
            Next k
            AppActivate excel_win   'Activate the excel window
            start_row = Selection.Rows(1).Row
            Range("P" & start_row - 1).Value = total_items
            'Application.Interactive = True
            NUM_On
End Sub


Private Sub generate_qty(ByVal start_row As Integer, ByVal last_row As Integer, ByVal source_book As String, ByVal source_sheet As String)
Dim qty As Variant
Dim current_row, z, eval_num As Integer
Dim id_number As String
    current_row = start_row                                     'Acts as a counter later on in the code
    Do While current_row <= last_row                            'While current row is less than or equal to last row
        id_number = Range("B" & current_row).Value              'Assign ID number to variable
        If id_number <> "" And IsNumeric(Left(id_number, 1)) = True Then  'If ID number is not blank or a special order part, then run the following code
            If current_row = start_row Then                     'If it is the first entry execute this code.
                Call qty_count(qty, current_row, last_row, id_number)   'count the number of occurences
            Else                                                'If it is not the first entry
                For z = start_row To current_row - 1            'search the rows from the row the code started on to one above the current row
                    If id_number = Range("B" & z).Value Then    'If the ID number is found in column B
                        GoTo NextIDBreak                        'Skip the code and enter 0 in for the quantity since the ID number has already been accounted for
                    End If
                Next z
                Call qty_count(qty, current_row, last_row, id_number)   'count the number of occurences
            End If
NextIDBreak:
            Call pop_number(current_row, qty)   'Populate the number of rows
        End If
    current_row = current_row + 1
    Loop
End Sub
Private Sub convert2text(ByVal start_row As Integer, ByVal last_row As Integer)
Dim row_count As Integer
Dim row_str As String
    Range("B" & start_row & ":B" & last_row).NumberFormat = "@"
    row_count = start_row
        Do While row_count <= last_row
            If IsNumeric(Left(Range("B" & row_count), 1)) = True Then
                Do While Len(Range("B" & row_count)) < 5
                    Range("B" & row_count).Value = "0" & Range("B" & row_count).Value
                Loop
            End If
            row_count = row_count + 1
        Loop
End Sub
Private Sub qty_count(ByRef qty As Variant, ByVal current_row As Integer, ByVal last_row As Integer, ByVal id_number As String)
Dim y As Integer
    For y = current_row To last_row                 'use i as a counter to count through all the rows
        If id_number = Range("B" & y) Then          'if the ID number is the same as the value of column B and row i
            qty = qty + Range("D" & y).Value        'add to the total quantity of the ID number
        End If
    Next y
End Sub
Private Sub pop_number(ByVal current_row As Integer, ByRef qty As Variant)
    Range("P" & current_row).Value = qty
    qty = 0
End Sub
Private Sub act_vis(ByVal serial_num As String, ByRef j As Integer, ByVal vis_window As String)
    
    Dim msg As String
    Dim ans As Integer
    
    On Error GoTo ErrCatcher1
    AppActivate vis_window
    Application.Wait Now + TimeSerial(0, 0, 2)    'Wait 2 seconds
    SendKeys "^m", True
    Application.Wait Now + TimeSerial(0, 0, 0.5)
    j = j + 1
    
ErrCatcher1:
    If Err.Number <> 0 Then
        Select Case Err.Number
            Case 5
                msg = "The visual window could not be set."
                msg = msg & "  This is likely because there are currently 2 windows open, or there were previously two windows open." & vbCrLf & vbCrLf
                msg = msg & "The window for visual needs to show: " & vbCrLf & vbTab & vis_window & vbCrLf & vbCrLf
                msg = msg & "Would you like to try again?"
                ans = MsgBox(msg, vbYesNo)
                Select Case ans
                    Case vbYes
                        Resume
                    Case Else
                        Exit Sub
                End Select
            Case Else
                MsgBox (Err.Number & " refers to an error: " & vbCrLf & Err.Description & vbCrLf & "Please add to error handler")
        End Select
        Resume
    End If
End Sub
Private Sub pop_vis(ByVal id_num As String, ByVal qty As Integer, ByVal mat_window As String, ByRef Wintext As String, ByVal excel_win As String)
    Dim msg As String
    Dim ans As Integer
    
    SendKeys id_num, True
    SendKeys "{TAB}", True
    Call window_get(Wintext)
    If mat_window <> Wintext Then
        AppActivate excel_win
        msg = "Auto-browse in enabled in Visual.  To disable select the options menu of the dialog box, and make sure there is not a check mark next to the "
        msg = msg & "'Auto browse enabled' selection box"
        msg = msg & vbCrLf & vbCrLf & "Are you ready to continue import?"
        ans = MsgBox(msg, vbYesNo)
        Select Case ans
            Case vbYes
                AppActivate mat_window
                Application.Wait Now + TimeSerial(0, 0, 2)    'Wait 2 seconds
                SendKeys "{TAB}", True
            Case Else
                End
        End Select
    End If
    SendKeys "{TAB}", True
    SendKeys qty, True
    SendKeys "+{F12}", True
End Sub
Private Sub window_get(ByRef Wintext As String)
    Dim hwnd As Long
    Dim L As Long
    hwnd = GetForegroundWindow()
    Wintext = String(255, vbNullChar)
    L = GetWindowText(hwnd, Wintext, 255)
    Wintext = Left(Wintext, InStr(1, Wintext, vbNullChar) - 1)
End Sub
Private Sub NUM_On()  'Turn NUM-Lock on
  If Not (GetKeyState(vbKeyNumlock) = 1) Then
    keybd_event VK_NUMLOCK, 1, 0, 0
    keybd_event VK_NUMLOCK, 1, KEYEVENTF_KEYUP, 0
  End If
End Sub

