Attribute VB_Name = "VisualImport"
Option Explicit

Private Declare Function GetForegroundWindow Lib "user32" _
    () As Long
Private Declare Function GetWindowText Lib "user32" _
    Alias "GetWindowTextA" (ByVal hwnd As Long, _
    ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function CloseClipboard Lib "user32" () As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_NUMLOCK = &H90
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub VisImport()
Dim arr As Variant
Dim arr_vis_mfg() As Variant
Dim id_num_char As Variant
ReDim Preserve arr_vis_mfg(1 To 4, 1 To 2)

Dim last_row As Integer
Dim start_row As Integer
Dim current_row_1 As Integer
Dim current_row_2 As Integer
Dim i As Integer
Dim k As Integer
Dim a As Integer
Dim total_items As Integer
Dim start_item As Integer
Dim num_iterations As Integer
Dim current_iteration As Integer
Dim current_iteration_1 As Integer
Dim z As Integer
Dim sub_id As Integer
Dim timer_count_1 As Integer
Dim timer_count_2 As Integer
Dim sub_id_count As Integer

Dim Wintext As String
Dim source_book As String
Dim source_sheet As String
Dim vis_window As String
Dim excel_win As String
Dim serial_num As String
Dim id_num As String
Dim vendor_id As String
Dim vendor_part_id As String
Dim vendor_name As String
Dim num_eval As String
Dim S As String
Dim qty As String
Dim msg As String
Dim ans As String
Dim ch_vis_window As String

Dim DataObj As New MSForms.DataObject

source_book = ActiveWorkbook.Name   'Name of the current workbook
source_sheet = ActiveSheet.Name     'Name of the current worksheet
start_row = Selection.Rows(1).Row
last_row = Selection.Rows.Count + start_row - 1 'Get the last row from column D
Columns("S:S").NumberFormat = "@"


MsgBox ("Please select the desired location to extract the BoM in the visual window.") 'Double Check that it will import to the correct place before continuing

'Generate the total number of each ID number into column L if the ID number has not already been accounted for
Columns("P:S").ClearContents
Call generate_qty(start_row, last_row)

'Gets the name of the excel window so the application can recall it later with AppActivate
Call window_get(Wintext)
    excel_win = Wintext
    
'Make sure the serial number is in cell H2, if it is not, prompt the user for it
serial_num = InputBox("Please type in the serial number of the project")
'Range("P1").Value = serial_num  'Serial number needed to activate visual window

'Store the name of the visual window you are looking for here
vis_window = "Manufacturing Window - Infor VISUAL - CSZ - [" & serial_num & "/1]"

'Empty Clipboard
Call ClearClipboard

'===============================================================================================================================================================
AppActivate vis_window  'Activate visual window
Call extract_num_iterations_1
If DataObj Is Nothing Then
    Set DataObj = New MSForms.DataObject
End If
DataObj.GetFromClipboard        'Assign the clipboard information to DataObj
Do While DataObj.GetFormat(1) = False
    Sleep (100)
    If timer_count_1 = 40 Then
        MsgBox ("Total number of items could not be counted")
        End
    End If
    timer_count_1 = timer_count_1 + 1
Loop
total_items = DataObj.GetText

If total_items = 0 Then
    End
End If
total_items = total_items / 10
Call ClearClipboard

'===============================================================================================================================================================
Call extract_num_iterations_2
If DataObj Is Nothing Then
    Set DataObj = New MSForms.DataObject
End If
DataObj.GetFromClipboard        'Assign the clipboard information to DataObj
Do While DataObj.GetFormat(1) = False
    Sleep (100)
    If timer_count_1 = 20 Then
        MsgBox ("Total number of items could not be counted")
        End
    End If
    timer_count_1 = timer_count_1 + 1
Loop
start_item = DataObj.GetText
If start_item = 0 Then
    End
End If
Call ClearClipboard
start_item = start_item / 10
num_iterations = total_items - start_item
sub_id_count = start_item - 1
Call open_edit   'Open the edit material window
For current_iteration = 1 To num_iterations
NextSubIDBreak:
sub_id_count = sub_id_count + 1
'===============================================================================================================================================================
    'Send sub id number
    Call send_sub_id(sub_id_count)
'===============================================================================================================================================================
    'Get ID number
    Call get_id_num

    'If DataObj does not exist, create the object
    If DataObj Is Nothing Then
        Set DataObj = New MSForms.DataObject
    End If
    DataObj.GetFromClipboard        'Assign the clipboard information to DataObj
    ReDim Preserve arr_vis_mfg(1 To 4, 1 To current_iteration)  'Create an array that stores the id number, mfg number, and qty
    timer_count_1 = 1     'Initialize timer count
    
    'Wait until the clipboard is populated
    Do While DataObj.GetFormat(1) = False
        Sleep (100)
        If timer_count_1 = 20 Then
            current_iteration = current_iteration + 1
            If current_iteration > num_iterations Then
                Exit For
            Else
                GoTo NextSubIDBreak
            End If
        End If
        timer_count_1 = timer_count_1 + 1
    Loop
    
    'Assign id number to the array and clear the clipboard
    arr_vis_mfg(1, current_iteration) = DataObj.GetText
    Call ClearClipboard
'===============================================================================================================================================================
    'Get QTY
    Call get_qty
    'If DataObj does not exist, create the object
    If DataObj Is Nothing Then
        Set DataObj = New MSForms.DataObject
    End If
    DataObj.GetFromClipboard        'Assign the clipboard information to DataObj
    
    'Wait until the clipboard is populated
    Do While DataObj.GetFormat(1) = False
    Loop
    'Assign qty to the array and clear the clipboard
    arr_vis_mfg(2, current_iteration) = DataObj.GetText
    Call ClearClipboard
'===============================================================================================================================================================
    'Vendor ID
    Call get_vendor_id
    'If DataObj does not exist, create the object
    If DataObj Is Nothing Then
        Set DataObj = New MSForms.DataObject
    End If
    DataObj.GetFromClipboard        'Assign the clipboard information to DataObj
    
    'Loop until clipboard is populated or until timer reaches 1 second
    timer_count_2 = 0
    Do While DataObj.GetFormat(1) = False
        If timer_count_2 = 10 Then
            Exit Do
        End If
        Sleep (100)
        timer_count_2 = timer_count_2 + 1
    Loop
    
    'If clipboard data is populated assign it to the array otherwise give the array a blank value and clear the clipboard
    If DataObj.GetFormat(1) = True Then
        arr_vis_mfg(3, current_iteration) = DataObj.GetText
    Else
        arr_vis_mfg(3, current_iteration) = ""
    End If
    Call ClearClipboard

'===============================================================================================================================================================
    'Vendor Part #
    Call get_vendor_part_id
    'If DataObj does not exist, create the object
    If DataObj Is Nothing Then
        Set DataObj = New MSForms.DataObject
    End If
    DataObj.GetFromClipboard        'Assign the clipboard information to DataObj
    
    'Loop until clipboard is populated or until timer reaches 1 second
    timer_count_2 = 0
    Do While DataObj.GetFormat(1) = False
        If timer_count_2 = 10 Then
            Exit Do
        End If
        Sleep (100)
        timer_count_2 = timer_count_2 + 1
    Loop
    
    'If clipboard data is populated assign it to the array otherwise give the array a blank value and clear the clipboard
    If DataObj.GetFormat(1) = True Then
        arr_vis_mfg(4, current_iteration) = DataObj.GetText
    Else
        arr_vis_mfg(4, current_iteration) = ""
    End If
    Call ClearClipboard
Next current_iteration

'===============================================================================================================================================================
'Switch to Excel
current_iteration_1 = 0
AppActivate excel_win       'Activate excel
For current_iteration_1 = 1 To num_iterations
    current_row_1 = start_row   'initialize the current row to start at the first id number
    id_num = arr_vis_mfg(1, current_iteration_1)    'assign array value storing the id number to a variable
    qty = arr_vis_mfg(2, current_iteration_1)       'assign array value storing the qty to a variable
    vendor_id = arr_vis_mfg(3, current_iteration_1)   'assign array value storing the mfg number to a variable
    vendor_part_id = arr_vis_mfg(4, current_iteration_1)
    Do While current_row_1 <= last_row
        If id_num = CStr(Range("B" & current_row_1)) And id_num <> "" Then    'If the id number from visual matches the id number of the current row
            Range("Q" & current_row_1).Value = qty
            Range("R" & current_row_1).Value = vendor_id
            Range("S" & current_row_1).Value = vendor_part_id
            vendor_name = Range("T" & current_row_1).Value
            Call row_eval(current_row_1, vendor_part_id, qty, vendor_name)      'Evaluate if the mfg, mfg number, and qty match in visual and the excel BoM
            If CInt(qty) > Range("D" & current_row_1).Value Then    'If the qty is greater than the value stored in column C of the active variable there must be another instance to evaluate
                For current_row_2 = current_row_1 + 1 To last_row
                    If id_num = Range("B" & current_row_2).Value Then   'If the id_number matches column B of the active row
                        qty = 0     'make qty = 0 since the total qty has been assigned to the Value in current_row_1
                        Range("Q" & current_row_2).Value = qty
                        Range("R" & current_row_2).Value = vendor_id
                        Range("S" & current_row_2).Value = vendor_part_id
                        vendor_name = Range("T" & current_row_1).Value
                        Call row_eval(current_row_2, vendor_part_id, qty, vendor_name)   'Evaluate whether or not the vendor_id and qty match
                    End If
                Next current_row_2
            End If
            GoTo NextBreak
        End If
    current_row_1 = current_row_1 + 1
    Loop
NextBreak:
Next current_iteration_1
NUM_On
End Sub

Private Sub row_eval(ByVal current_row As Integer, ByVal vendor_part_id As String, ByVal qty As String, ByVal vendor_name As String)
    If CStr(Range("I" & current_row).Value) <> vendor_name Then
        Range("I" & current_row).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
    If (CStr(Range("N" & current_row).Value) <> vendor_part_id And vendor_part_id <> "") Then
        Range("N" & current_row).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
    If CStr(Range("P" & current_row).Value) <> qty Then
        Range("P" & current_row).Select
        With Selection.Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 65535
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
End Sub
Private Sub ClearClipboard()
    OpenClipboard (0&)
    EmptyClipboard
    CloseClipboard
End Sub

Private Sub window_get(ByRef Wintext As String)
    Dim hwnd As Long
    Dim L As Long
    hwnd = GetForegroundWindow()
    Wintext = String(255, vbNullChar)
    L = GetWindowText(hwnd, Wintext, 255)
    Wintext = Left(Wintext, InStr(1, Wintext, vbNullChar) - 1)
End Sub
Private Sub generate_qty(ByVal start_row As Integer, ByVal last_row As Integer)
Dim qty As Variant
Dim current_row, z, eval_num As Integer
Dim source_book, source_sheet, id_number As String
    current_row = start_row                                     'Acts as a counter later on in the code
    source_book = ActiveWorkbook.Name                           'name of the workbook
    source_sheet = ActiveSheet.Name                             'name of the sheet
    Do While current_row <= last_row                            'While current row is less than or equal to last row
        id_number = Range("B" & current_row).Value              'Assign ID number to variable
        If id_number <> "" And IsNumeric(Left(id_number, 1)) = True Then  'If ID number is not blank then run the following code
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
Private Sub extract_num_iterations_1()
    SendKeys "^m", True
    Application.Wait Now + TimeSerial(0, 0, 2)
    SendKeys "+{TAB}", True
    SendKeys "+{TAB}", True
    SendKeys "^(c)", True
    SendKeys "{ESC}", True
End Sub
Private Sub extract_num_iterations_2()
    SendKeys "^e", True
    Application.Wait Now + TimeSerial(0, 0, 2)
    SendKeys "+{TAB}", True
    SendKeys "^(c)", True
    SendKeys "{ESC}", True
End Sub
Private Sub open_edit()
    Application.Wait Now + TimeSerial(0, 0, 0.5)
    SendKeys "^e", True
    Application.Wait Now + TimeSerial(0, 0, 2)
End Sub

Private Sub send_sub_id(ByVal sub_id_count As Integer)
Dim sub_id As Integer
    sub_id = 10 * sub_id_count
    SendKeys "+{TAB}", True
    SendKeys sub_id, True
End Sub

Private Sub get_id_num()
    SendKeys "{TAB}", True
    SendKeys "^(c)", True
End Sub
Private Sub get_qty()
    SendKeys "{F3}", True
    SendKeys "{TAB}", True
    SendKeys "{TAB}", True
    SendKeys "^(c)", True
End Sub
Private Sub get_vendor_id()
    SendKeys "{F6}", True
    SendKeys "{TAB}", True
    SendKeys "^c", True
End Sub
Private Sub get_vendor_part_id()
    SendKeys "{TAB}", True
    SendKeys "^c", True
    SendKeys "+{TAB}", True
    SendKeys "+{TAB}", True
    SendKeys "+{TAB}", True
End Sub
Private Sub NUM_On()  'Turn NUM-Lock on
  If Not (GetKeyState(vbKeyNumlock) = 1) Then
    keybd_event VK_NUMLOCK, 1, 0, 0
    keybd_event VK_NUMLOCK, 1, KEYEVENTF_KEYUP, 0
  End If
End Sub








