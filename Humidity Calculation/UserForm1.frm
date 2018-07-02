VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5865
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub sbTemp_Change()
    If UserForm1.lblUnitTemp = "°F" Then
        UserForm1.lblUnitTemp.Caption = "°C"
    Else
        UserForm1.lblUnitTemp.Caption = "°F"
    End If
End Sub

Private Sub cbAdd_Click()
    Dim counter_1 As Integer
    Dim counter_2 As Integer

    Dim profile_list As Variant
    Dim swap_1 As Variant
    Dim swap_2 As Variant

    With UserForm1.lbProfiles
        .AddItem UserForm1.txtTemp.Value & UserForm1.lblUnitTemp.Caption
        .List(.ListCount - 1, 1) = UserForm1.txtRH.Value & UserForm1.lblUnitRH.Caption

        profile_list = .List
        For counter_1 = LBound(profile_list, 1) To UBound(profile_list, 1)
            For counter_2 = counter_1 + 1 To UBound(profile_list, 1)
                If profile_list(counter_1, 0) <= profile_list(counter_2, 0) And profile_list(counter_1, 1) <= profile_list(counter_2, 1) Then
                    swap_1 = profile_list(counter_1, 0)
                    swap_2 = profile_list(counter_1, 1)
                    profile_list(counter_1, 0) = profile_list(counter_2, 0)
                    profile_list(counter_1, 1) = profile_list(counter_2, 1)
                    profile_list(counter_2, 0) = swap_1
                    profile_list(counter_2, 1) = swap_2
                End If
            Next counter_2
        Next counter_1
        
        .Clear
        .List = profile_list
    End With
    With UserForm1
        .txtTemp.Value = ""
        .txtRH.Value = ""
        .txtTemp.SetFocus
    End With
End Sub

Private Sub cbDelete_Click()
    UserForm1.lbProfiles.RemoveItem (UserForm1.lbProfiles.ListIndex)
End Sub

Private Sub lbProfiles_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Then
        cbDelete_Click
    End If
End Sub

Private Sub cbRun_Click()
    
    Dim profile_list As Variant
    Dim list_counter As Integer
    Dim profile_name As String
    Dim profile_temp As Double
    Dim profile_temp_units As String
    Dim profile_RH As Double

    Do While ActiveSheet.Scenarios.Count > 0
        ActiveSheet.Scenarios(1).Delete
    Loop
    profile_list = UserForm1.lbProfiles.List
    For list_counter = LBound(profile_list, 1) To UBound(profile_list, 1)
        profile_name = profile_list(list_counter, 0) & ", " & profile_list(list_counter, 1)
        profile_temp_units = Right(profile_list(list_counter, 0), 2)
        If profile_temp_units = "°C" Then
            profile_temp = Mid(profile_list(list_counter, 0), 1, Len(profile_list(list_counter, 0)) - 2)
            profile_temp = profile_temp * 1.8 + 32
        Else
            profile_temp = Mid(profile_list(list_counter, 0), 1, Len(profile_list(list_counter, 0)) - 2)
        End If
        profile_RH = Mid(profile_list(list_counter, 1), 1, Len(profile_list(list_counter, 1)) - 1)
        profile_RH = profile_RH / 100
        ActiveSheet.Scenarios.Add Name:=profile_name, ChangingCells:=Range("C17:C18") _
            , Values:=Array(profile_temp, profile_RH), Comment:="Created by Cody Baker on 4/6/2018" _
            , Locked:=False, Hidden:=False
    Next list_counter
    ActiveSheet.Scenarios.CreateSummary ReportType:=xlStandardSummary, _
        ResultCells:=Range("C25")
    UserForm1.Hide
End Sub

