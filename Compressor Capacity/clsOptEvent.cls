VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsOptEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public WithEvents OptionButtonEvents As MSForms.OptionButton
Attribute OptionButtonEvents.VB_VarHelpID = -1

Private Sub OptionButtonEvents_Click()
    'In the event one of the initial option buttons is selected this sub is run so that the compressor combinations Copeland allows are able to be selected through each
    'iteration.  For example if you select 60Hz, only compressor with voltages and phases offered in the 60Hz will be available.  The rest will be locked and grayed out.
    Dim opt_ctrl_2 As Control
    Dim comp As String  'Public declaration of comp does not work in the class module
    Dim num_rows_ref, count_5, count_6, count_7, count_8, count_9, count_10 As Integer
    Dim new_comp_options_arr(), comp_log_arr, cond_arr_1() As Variant

    'Turn off for faster running
     Application.ScreenUpdating = False

    Call ClearingModule.clear_opt_buttons
        
    'Setup information to populate to the userform from the compressor log on the R:\ drive.  This is the same code from the UserForm1 code
    comp = UserForm1.comp_selection
    ReDim cond_arr_1(0 To 4)
    Call PopulateData.comp_pop_data(num_rows_ref, comp_log_arr)

    'Convert all elements of the comp_log_arr
    For count_7 = 1 To num_rows_ref - 1
        For count_8 = 1 To 24
            comp_log_arr(count_7, count_8) = CStr(comp_log_arr(count_7, count_8))
        Next count_8
    Next count_7
        
    'Frames only allow for one option button to be selected.  The following statements loop through every option button in each frame in the optional compressor
    'controls frame and assign it to an array which will be evaluated later so only possible combinations will be provided

    Call DataExtract.ob_eval(cond_arr_1)
        If cond_arr_1(3) = "" Then
            Exit Sub
        End If
    count_5 = 0   'column counter for new_comp_options_arr
    count_6 = 1   'Array row counter

    Call DataExtract.data_extract_2(count_5, count_6, num_rows_ref, comp, comp_log_arr, cond_arr_1, new_comp_options_arr)
    'Turn back on
     Application.ScreenUpdating = True
     
End Sub
