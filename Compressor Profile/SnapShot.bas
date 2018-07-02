Attribute VB_Name = "SnapShot"
Option Explicit
Public Sub run_program()
    Dim sht_start_temp As String
    Dim source_wb_name As String
    Dim source_ws_name As String
    Dim comp_wb_name As String
    Dim comp_sht_name As String
    Dim sht_end_temp As String
    Dim temp_unit As String
    
    Dim start_temp As Integer
    Dim end_temp As Integer
    Dim suc_temp As Integer
    Dim cond_temp As Integer
    Dim max_temp As Integer
    Dim min_temp As Integer
    Dim i As Integer
    
    
    Dim ans As Variant
    
    Dim sys_cap As Double
    Dim pull_down_time As Double
    
    Application.DisplayAlerts = False

    Do Until (temp_unit = "F" Or temp_unit = "C")
        temp_unit = UCase(InputBox("Is the temperature unit of measurement in F or C?", "Temperature Unit"))
        If temp_unit = "" Then
            End
        End If
    Loop
    source_wb_name = ActiveWorkbook.Name
    source_ws_name = ActiveSheet.Name
    pull_down_time = 0
    cond_temp = 95
    max_temp = -500
    min_temp = 500
    i = 3
NextIteration:
     Call unit_eval(temp_unit, start_temp, end_temp)
     Call name_sheet(start_temp, sht_start_temp, temp_unit, end_temp, sht_end_temp)
     If start_temp > max_temp Then
        max_temp = start_temp
        'Application.ScreenUpdating = False
        Workbooks(source_wb_name).Sheets("Graph").Range("A" & i).Value = start_temp
        Workbooks(source_wb_name).Sheets("Graph").Range("B" & i).Value = pull_down_time
        'Sheets("Graph").Select
        'Application.ScreenUpdating = True
        i = i + 1
     End If
     If end_temp < min_temp Then
        min_temp = end_temp
    End If
     
    'Identify all workbooks
    comp_wb_name = "Master - Smart Compressor Capacity Sheet"
    source_wb_name = ActiveWorkbook.Name
    source_ws_name = ActiveSheet.Name

    Call master_input(suc_temp, end_temp, cond_temp)
    Application.ScreenUpdating = False
    Call activate_master(comp_wb_name, comp_sht_name)

    Call master_sht_op(comp_wb_name, comp_sht_name, suc_temp, cond_temp, sys_cap, start_temp)
    Workbooks(source_wb_name).Activate
    Application.ScreenUpdating = True
    Call GoalSeek_2(sys_cap, pull_down_time, source_wb_name, source_ws_name)
    ans = MsgBox("Would you like to do another iteration?", vbYesNo)
    Select Case ans
        Case vbYes
            pull_down_time = Round(pull_down_time, 0)
            Workbooks(source_wb_name).Sheets("Graph").Range("A" & i).Value = end_temp
            Workbooks(source_wb_name).Sheets("Graph").Range("B" & i).Value = pull_down_time
            Workbooks(source_wb_name).Sheets("Graph").Range("C" & i).Value = sys_cap
            i = i + 1
            Call insert_sheet(source_wb_name, source_ws_name)
            GoTo NextIteration
        Case vbNo
            pull_down_time = Round(pull_down_time, 0)
            Workbooks(source_wb_name).Sheets("Graph").Range("A" & i).Value = end_temp
            Workbooks(source_wb_name).Sheets("Graph").Range("B" & i).Value = pull_down_time
            Workbooks(source_wb_name).Sheets("Graph").Range("C" & i).Value = sys_cap
            i = i + 1
            MsgBox ("It will take approximately " & pull_down_time & " minutes to go from " & max_temp & temp_unit & " to " & min_temp & temp_unit)
    End Select
    Application.DisplayAlerts = True
End Sub

Private Sub unit_eval(ByVal temp_unit As String, ByRef start_temp As Integer, ByRef end_temp As Integer)
    If temp_unit = "F" Then
        start_temp = InputBox("Enter the From value located in cell B9", "Start Temp", end_temp)
        end_temp = InputBox("Enter the To value located in cell B10", "End Temp", start_temp - 10)
        Call GoalSeek_1(start_temp, end_temp)
    Else
        start_temp = InputBox("Enter the From value located in cell E3", "Start Temp", end_temp)
        Range("E3").Value = start_temp
        end_temp = InputBox("Enter the To value located in cell E4", "End Temp")
        Range("E4").Value = end_temp
    End If
End Sub

Private Sub GoalSeek_1(ByVal start_temp As Integer, ByVal end_temp As Integer)
    Range("B9").GoalSeek Goal:=start_temp, ChangingCell:=Range("E3")
    Range("B10").GoalSeek Goal:=end_temp, ChangingCell:=Range("E4")
End Sub

Private Sub name_sheet(ByVal start_temp As Integer, ByRef sht_start_temp As String, ByVal temp_unit As String, ByVal end_temp As Integer, ByRef sht_end_temp As String)
    If start_temp > 0 Then
        sht_start_temp = "+" & start_temp
    ElseIf start_temp = 0 Then
        sht_start_temp = 0
    Else
        sht_start_temp = start_temp
    End If
    If end_temp > 0 Then
        sht_end_temp = "+" & end_temp
    ElseIf end_temp = 0 Then
        sht_end_temp = 0
    Else
        sht_end_temp = end_temp
    End If
    ActiveSheet.Name = "THERMAL-WM (SS " & sht_start_temp & temp_unit & " to " & sht_end_temp & temp_unit & ")"
End Sub

Private Sub master_input(ByRef suc_temp As Integer, ByVal end_temp As Integer, ByRef cond_temp As Integer)
    If Range("B10").Value > 10 Then
        suc_temp = 0
    Else
        suc_temp = Range("B10").Value - 10
    End If
    suc_temp = InputBox("Input the suction temperature you wish to evaluate the system at", "Suction Temp", suc_temp)
    cond_temp = InputBox("Input the condensing temperature you wish to evaluate the system at", "Condensing Temp", cond_temp)
End Sub

Private Sub activate_master(ByVal comp_wb_name As String, ByRef comp_sht_name As String)
    Workbooks(comp_wb_name).Activate
    comp_sht_name = ActiveSheet.Name
End Sub

Private Sub master_sht_op(ByVal comp_wb_name As String, ByVal comp_sht_name As String, ByVal suc_temp As Integer, ByVal cond_temp As Integer, ByRef sys_cap As Double, _
                          ByVal start_temp As Integer)
    Workbooks(comp_wb_name).Sheets(comp_sht_name).Range("R7").Value = start_temp
    Workbooks(comp_wb_name).Sheets(comp_sht_name).Range("C7").Value = suc_temp
    Call activate_master(comp_wb_name, comp_sht_name)
    Workbooks(comp_wb_name).Sheets(comp_sht_name).Range("F7").Value = cond_temp
    Call activate_master(comp_wb_name, comp_sht_name)
    sys_cap = Workbooks(comp_wb_name).Sheets(comp_sht_name).Range("J7").Value
    sys_cap = Round(sys_cap, 2)
End Sub

Private Sub GoalSeek_2(ByVal sys_cap As Double, ByRef pull_down_time As Double, ByVal source_wb_name As String, ByVal source_ws_name As String)
    Workbooks(source_wb_name).Sheets(source_ws_name).Range("B83").GoalSeek Goal:=sys_cap, ChangingCell:=Workbooks(source_wb_name).Sheets(source_ws_name).Range("E6")
    pull_down_time = Workbooks(source_wb_name).Sheets(source_ws_name).Range("E6") + pull_down_time
End Sub

Private Sub insert_sheet(ByVal source_wb_name As String, ByVal source_ws_name As String)
    Workbooks(source_wb_name).Sheets(source_ws_name).Copy After:=Workbooks(source_wb_name).Sheets(source_ws_name)
End Sub
