Attribute VB_Name = "Init_TCalc_form"
Option Explicit

Public Sub Initialize_TCalc_form()
    Dim last_voltage_row As Integer
    Dim source_book As String
    Dim chamber_manu_arr As Variant
    Dim master_collection As Collection
    Set master_collection = New Collection

    '=========================================================================================================
    'Turn off Screen Updating and Extract Data from the "T-Calc User Interface" file on NEW R DRIVE
    '=========================================================================================================
    Application.ScreenUpdating = False
    source_book = ActiveWorkbook.Name
    gather_data master_collection

    '=========================================================================================================
    'Turn on Screen Updating and Activate the T-Calc Workbook
    '=========================================================================================================
    Application.ScreenUpdating = True
    Workbooks(source_book).Activate


    With formTCalc
        '=========================================================================================================
        'Assign the Voltage Code data to the Voltage Code Combo Box in the Userform
        '=========================================================================================================
        With .cmbVoltageCode
            .List = master_collection(5)
            .ColumnCount = 4
            .ColumnWidths = "0.25in;0.75in;0.25in;0.25in"
            .TextAlign = 2
        End With
        '=========================================================================================================
        'Assign the Chamber Manufacturer data to the Chamber Manufacturer Combo Box in the Userform
        'Only Assign the unique values up front for the User to select and add a lookup fuction later
        '=========================================================================================================
        With .cmbChamberManufacturer
            chamber_manu_arr = id_unique_values(master_collection(3))
            .List = chamber_manu_arr
            .TextAlign = 2
        End With
        '=========================================================================================================
        'Assign the Plenum data to the Plenum Combo Box in the Userform
        'Only Assign the unique values up front for the User to select and add a lookup fuction later
        '=========================================================================================================
        With .cmbPlenumType
            .List = id_unique_values(master_collection(6))
            .TextAlign = 2
        End With
        '=========================================================================================================
        'Assign the column headers of the profiles to the listbox in the Userform
        '=========================================================================================================
        With .lbProfiles
            .ColumnCount = 3
            .ColumnWidths = "2 in; 2 in; 1 in"
            .AddItem "Starting Temp"
            .List(0, 1) = "Final Temp"
            .List(0, 2) = "Time"
        End With
        '=========================================================================================================
        'Populate the insulation data combo box
        '=========================================================================================================
        With .cmbInsulation
            .List = Array("Fiberglass", "Foam")
            .TextAlign = 2
        End With
    End With
    formTCalc.Show
End Sub
Private Sub gather_data(ByRef master_collection)

    Dim file_path As String
    Dim file_name As String
    Dim ui_wb As String

    Dim last_row As Integer
    Dim last_col As Integer
    Dim count As Integer

    Dim arr As Variant

    Dim wks As Worksheet

    '=========================================================================================================
    'Open the "T-Calc User Interface" file on NEW R DRIVE
    '=========================================================================================================
    file_path = "I:\engineering\Thermal Calculator\Thermal Calculator v2.00\"
    file_name = "T-Calc User Interface.xlsm"
    Workbooks.Open _
        Filename:=file_path & file_name, _
        ReadOnly:=True
    ui_wb = ActiveWorkbook.Name


    '==================================================================================================================
    'Go through each sheet and extract all of the information from cell A1 to the last row and column with data in it
    '==================================================================================================================
    For Each wks In Worksheets
        With Worksheets(wks.Name)
            last_row = .Cells(.Rows.count, 1).End(xlUp).Row
            last_col = .Cells(2, .Columns.count).End(xlToLeft).Column
            arr = .Range(.Cells(2, 2), .Cells(last_row, last_col))
        End With
        master_collection.Add arr
    Next

    '====================================================================================================================
    'After all data has been gathered into a large array made up of all the smaller arrays assign it the the gather_data
    'function and close the User Interface workbook
    '====================================================================================================================
    Workbooks(ui_wb).Close SaveChanges:=False
End Sub

Private Function id_unique_values(ByRef arr)
    Dim arr_counter As Integer
    Dim new_arr_counter As Integer
    Dim element_counter As Integer
    Dim new_arr() As Variant
    Dim bool_exists As Boolean
    ReDim Preserve new_arr(0)

    '====================================================================================================================
    'Evaluate all elements in the passed array to extract only unique values and remove duplicates, assign the unique
    'values to the id_unique_values fuction
    '====================================================================================================================
    element_counter = -1
    For arr_counter = 1 To UBound(arr, 1)
        For new_arr_counter = 0 To UBound(new_arr)
            If new_arr(new_arr_counter) = arr(arr_counter, 1) Then
                bool_exists = True
                Exit For
            Else
                bool_exists = False
            End If
        Next
        If bool_exists = False Then
            element_counter = element_counter + 1
            ReDim Preserve new_arr(0 To element_counter)
            new_arr(element_counter) = arr(arr_counter, 1)
        End If
    Next
    id_unique_values = new_arr
End Function

