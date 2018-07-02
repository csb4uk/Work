VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13995
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lblContinue_Click()
    UserForm1.Hide
    Call ExportBoMVisual.ex_vis
End Sub

Private Sub lblDelete_Click()
    Dim removed_item As String
    Dim counter_1 As Integer
    Dim counter_2 As Integer
    Dim num_items As Integer
    Dim li_int As Integer
    Dim arr As Variant
    
    num_items = 0
    li_int = UserForm1.ListBox1.ListIndex
    removed_item = UserForm1.ListBox1.List(UserForm1.ListBox1.ListIndex, 1) 'Store the removed item

    UserForm1.ListBox1.RemoveItem (UserForm1.ListBox1.ListIndex)    'Remove the Item from the list
    arr = UserForm1.ListBox1.List   'Store the Listbox as an array
    ReDim Preserve arr(UBound(arr, 1), UserForm1.ListBox1.ColumnCount - 1)  'Trim the array so only elements used are in it

    'See if another instance of the removed item exists
    For counter_1 = LBound(arr, 1) To UBound(arr, 1)
        If removed_item = arr(counter_1, 1) Then
            Exit For
        End If
    Next
    'If it does find the number of times it still appears
    If counter_1 < UBound(arr, 1) Then
        For counter_2 = LBound(arr, 1) To UBound(arr, 1)
            If removed_item = arr(counter_2, 1) Then
                num_items = num_items + 1
            End If
        Next
        arr(counter_1, 6) = num_items
    End If
    UserForm1.ListBox1.List = arr   'Store the Listbox as an array
    UserForm1.ListBox1.Selected(li_int) = True
End Sub

Private Sub ListBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Then
        lblDelete_Click
    End If
End Sub
