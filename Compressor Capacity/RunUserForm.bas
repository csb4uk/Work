Attribute VB_Name = "RunUserForm"
Option Explicit

Sub Run_UserForm()

Dim ref_book, ref_sheet_1 As String
Dim num_comp, center, offset As Integer

'Gather list of possible compressors from the compressor log on the R:\ drive
Workbooks.Open _
    Filename:="R:\NEW R DRIVE\Refrigeration Compressors\Compressor log.xlsm", _
    ReadOnly:=True
ref_book = "Compressor log"     'Reference Workbook
ref_sheet_1 = "List"            'Reference Sheet List that has the list of all the compressors
num_comp = Workbooks(ref_book).Worksheets(ref_sheet_1).Cells(Rows.Count, 2).End(xlUp).Row   'number of compressors
UserForm1.comp_selection.List = Workbooks(ref_book).Worksheets(ref_sheet_1).Range("B2:B" & num_comp).Value  'Populate the drop down list of compressors on the UserForm
Workbooks(ref_book).Close SaveChanges:=False    'Close the compressor log workbook, do not save changes

UserForm1.Width = 468
UserForm1.Height = 408

'Setup UserForm1
center = (UserForm1.Width) / 2
offset = 12

With UserForm1
    With .comp_selection
        .Top = offset
        .Height = 18
        .Width = 114
        .Left = (UserForm1.Width) / 2 - .Width - 5
        .TextAlign = fmTextAlignLeft
    End With
    With .comp_label_1
        .Top = offset
        .Height = 18
        .Width = 66
        .Left = UserForm1.comp_selection.Left - .Width - 5
        .TextAlign = fmTextAlignRight
    End With
    With .clear_command
        .Top = offset
        .Height = 24
        .Width = 132
        .Left = (UserForm1.Width) / 2 + 5
    End With
    With .comp_frame
        .Top = UserForm1.comp_selection.Top + UserForm1.comp_selection.Height + offset
        .Height = 66
        .Width = 114
        .Left = center - ((.Width + 12 + 246) / 2)
    End With
    With .rpm_frame
        .Top = UserForm1.comp_selection.Top + UserForm1.comp_selection.Height + offset
        .Height = 66
        .Width = 246
        .Left = UserForm1.comp_frame.Left + UserForm1.comp_frame.Width + offset
    End With
    With .comp_control_frame
        .Top = UserForm1.comp_frame.Height + UserForm1.comp_frame.Top + offset
        .Height = 264
        .Width = UserForm1.Width - 24
        .Left = offset
    End With
    With .prim_ref_frame
        .Top = offset
        .Height = 78
        .Width = 276
        .Left = offset
    End With
    With .casc_ref_frame
        .Top = offset
        .Height = 76
        .Width = 132
        .Left = UserForm1.prim_ref_frame.Left + UserForm1.prim_ref_frame.Width + offset
    End With
    With .v_frame
        .Top = UserForm1.prim_ref_frame.Top + UserForm1.prim_ref_frame.Height + offset
        .Height = 108
        .Width = 306
        .Left = offset
    End With
    With .hz_frame
        .Top = UserForm1.v_frame.Top
        .Height = 52
        .Width = 92
        .Left = UserForm1.comp_control_frame.Width - (.Width + offset)
    End With
    With .ph_frame
        .Height = 52
        .Top = UserForm1.v_frame.Top + UserForm1.v_frame.Height - .Height
        .Width = 92
        .Left = UserForm1.comp_control_frame.Width - (.Width + offset)
    End With
    With .cmd_pop_sheets
        .Height = 24
        .Top = UserForm1.v_frame.Top + UserForm1.v_frame.Height + offset
        .Width = 174
        .Left = (UserForm1.comp_control_frame.Width / 2) - (.Width / 2)
    End With
    .Top = Application.Top + 125
    .Left = Application.Left + 25
End With
UserForm1.Show
End Sub
