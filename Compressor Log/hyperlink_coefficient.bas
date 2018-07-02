Attribute VB_Name = "hyperlink_coefficient"
Sub hyperlink_coefficient()

'Turn stuff off so the program can run faster
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False


act_row = ActiveCell.Row  'Starting location to add compatible refrigerants
num_rows = Workbooks("Compressor log").Sheets("Sheet1").Cells(Rows.Count, 2).End(xlUp).Row 'Finds the number of rows used in the doc
'If you only want to do one iteration uncomment the following line
num_rows = act_row

Do While act_row < num_rows + 1  'While all the cells have not been evaluated
    If ActiveSheet.Range("N" & act_row).Value = "" Then
        Const FOLDER_PATH As String = "R:\NEW R DRIVE\Refrigeration Compressors\Copeland" 'assign master file location
        
        comptype = ActiveSheet.Range("A" & (act_row)).Value  'Assign the compressor type (i.e. Hermetic, Semi-Hermetic, or Scroll to a variable
        comp = ActiveSheet.Range("B" & (act_row)).Value      'Assign the actual compressor name as a variable
        wb_path = Application.ActiveWorkbook.FullName
        If comptype = "Hermetic" Then
            file_path = FOLDER_PATH & "\" & comptype & "\" & comp & "\Master Compressor Capacity Information\"  'using the compressor type and compressor go out to the network drive where the matching_
                                                                                                                'information is kept
        Else
            file_path = FOLDER_PATH & "\" & comptype & "\Low Temperature\" & comp & "\Master Compressor Capacity Information\"
        End If
        
        comp_hz = ActiveSheet.Range("G" & (act_row)).Value   'store the compressor Hz as a variable since that is located in the file to be hyperlinked
        comp_code = ActiveSheet.Range("C" & (act_row)).Value 'store the compressor voltage code as a variable since that is located in the file to be hyperlinked
        file_name = Dir(file_path & "*" & comp_code & "*-" & comp_hz & ".csv") 'This searches a directory that matches the following form
            'R:\NEW R DRIVE\Refrigeration Compressors\Copeland\comptype\comp\Master Compressor Capacity Information
            'and finds a file that uses the comp_code as a wildcard and uses "-comp_hz.csv" as an anchor to match (i.e. PFV-50.csv).  The reason for the wildcard is because the TF5, TFD, TFC
            'or TFE can all have the same compressor coefficients
        
        If Len(file_name) > 0 Then                      'If there is no file name to be found in the directory
            With Sheet1                                     'If there is a file name, insert the hyperlink to cell N of the active row
                .Hyperlinks.Add Anchor:=.Range("N" & (act_row)), _
                Address:=file_path & file_name, _
                ScreenTip:="Compressor Coefficients", _
                TextToDisplay:="Coefficients"
            End With
        End If
    End If
    act_row = act_row + 1
Loop
'Turn the things on that were previously turned off to improve performance
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
End Sub
