VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formTCalc 
   Caption         =   "T-Calc Interface"
   ClientHeight    =   11325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   21915
   OleObjectBlob   =   "formTCalc.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formTCalc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
    ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const VK_NUMLOCK = &H90
Private Const KEYEVENTF_KEYUP = &H2
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer

'===============================================================================================================================================
'Customer Information Controls
'===============================================================================================================================================
Private Sub cmbVoltageCode_AfterUpdate()
    pop_voltage_code
End Sub
Private Sub cmbVoltageCode_DropButtonClick()
    pop_voltage_code
End Sub
Private Sub pop_voltage_code()
    '===============================================================================================================================================
    'The combo box will naturally only display the information in the first column, so this will build a string that shows the voltage code as
    ' (Code Letter) - Voltage/Phase/Hz
    '===============================================================================================================================================
    With cmbVoltageCode
        If Len(.Value) = 1 Then
            .Text = .List(.ListIndex, 0) & " - " & .List(.ListIndex, 1) & "/" & .List(.ListIndex, 2) & "/" & .List(.ListIndex, 3)
        End If
    End With
End Sub
Private Sub txtCustomerLocationCity_AfterUpdate()
    formTCalc.txtCustomerLocationCity.Value = StrConv(formTCalc.txtCustomerLocationCity.Value, vbProperCase)
    clear_weather_data
End Sub
Private Sub txtCustomerLocationState_AfterUpdate()
    formTCalc.txtCustomerLocationState.Value = StrConv(formTCalc.txtCustomerLocationState.Value, vbProperCase)
    clear_weather_data
End Sub
Private Sub chboxIntl_Click()
    clear_weather_data
End Sub
Private Sub clear_weather_data()
    '===============================================================================================================================================
    'When the user populates a city value, or needs to change the city value
        'Clear the Elevation, Ambient Value, and Weather Station
    '===============================================================================================================================================
    With formTCalc
        .txtElevation.Value = ""
        .txtOutsideAmbient.Value = ""
        .cmbWeatherStation.Caption = "Weather Station"
    End With
End Sub
Private Sub cmbRetrieveAshrae_Click()
    '===============================================================================================================================================
    'This code is reponsible for the following:
        'Run a google search on the city and state/country to find the elevation
        'Run a google search on the city and state/country to find the longitude and latitude
        'If those searches return values find the closest weather station to the longitude and latitude of the city from the ASHRAE Data
            'found in the T-Calc User Interface
    '===============================================================================================================================================
    Dim http As Object
    Dim html As New HTMLDocument
    Dim topics As Object

    Dim ui_wb_ws_arr As Variant
    Dim pdf_arr As Variant

    Dim file_path As String
    Dim file_name As String
    Dim ui_wb As String
    Dim wp As String
    Dim city As String
    Dim state_or_country As String
    Dim lat_long As String
    Dim loc_latitude As String
    Dim loc_longitude As String
    Dim loc_latitude_dir As String
    Dim loc_longitude_dir As String
    Dim sb As String
    Dim ss As String
    Dim ws_latitude As String
    Dim ws_longitude As String
    Dim ws_latitude_dir As String
    Dim ws_longitude_dir As String
    Dim weather_station As String
    Dim ws_folder_path As String
    Dim ws_file_name As String
    Dim ws_file_exists As String

    Dim ws_latitude_num As Double
    Dim ws_longitude_num As Double
    Dim loc_latitude_num As Double
    Dim loc_longitude_num As Double
    Dim lat_diff As Double
    Dim long_diff As Double
    Dim dist_diff As Double
    Dim min_diff As Double
    Dim dry_bulb_temp As Double

    Dim ws_ws_last_row As Integer
    Dim ws_ws_last_col As Integer
    Dim row_counter_1 As Integer
    Dim col_counter_1 As Integer
    Dim ui_wb_ws_col_latitude As Integer
    Dim ui_wb_ws_col_longitude As Integer
    Dim ui_wb_ws_col_sn As Integer
    Dim min_diff_row As Integer
    Dim ui_wb_ws_last_row As Integer
    Dim ui_wb_ws_last_col As Integer
    Dim pdf_row_count As Integer
    Dim pdf_col_count As Integer
    Dim pdf_last_row As Integer
    Dim pdf_last_col As Integer

    '===============================================================================================================================================
    'Store the current workbook and worksheet as values. The focus will change and we will want to return here after this subroutine runs
    '===============================================================================================================================================
    sb = ActiveWorkbook.Name
    ss = ActiveSheet.Name

    '===============================================================================================================================================
    'Make sure the user has entered a valid city and state/country (not a blank or default value).  If it is not valid, give them a message box
    'and stop the subroutine.
    '===============================================================================================================================================
    If Trim(formTCalc.txtCustomerLocationState.Text) = "State/Country" Or Trim(formTCalc.txtCustomerLocationState.Text) = "" Then
        MsgBox ("Fill in Customer State/Country")
        Exit Sub
    End If

    '===============================================================================================================================================
    'Remove the current Elevation and Ambient Values in case the user has changed the city/state/country from a previous iteration
    '===============================================================================================================================================
    formTCalc.txtElevation.Value = ""
    formTCalc.txtOutsideAmbient.Value = ""
    
    '===============================================================================================================================================
    'Turn Screen Updating off to speed up the program
    '===============================================================================================================================================
    formTCalc.Hide
    Application.ScreenUpdating = False
    Application.Visible = False

    '===============================================================================================================================================
    'Run a Google Search on the Elevation.  When an elevation is found it is kept under the html class tag Z0LcW.  You can verify this by selecting
    'the text in the search and inspecting the element.
    '===============================================================================================================================================
    Set http = CreateObject("MSXML2.XMLHTTP")
    city = formTCalc.txtCustomerLocationCity.Value
    state_or_country = txtCustomerLocationState.Value
    If city <> "" Then
        wp = "https://www.google.com/search?q=" & city & "+" & state_or_country & "elevation"
    Else
        wp = "https://www.google.com/search?q=" & state_or_country & "elevation"
    End If
    http.Open "GET", wp, False
    http.send
    html.body.innerHTML = http.responseText
    Set topics = html.getElementsByClassName("Z0LcW")(0)
    If Not topics Is Nothing Then
        formTCalc.txtElevation.Value = topics.innerText
    End If
    '===============================================================================================================================================
    'Run a Google Search on the Longitude and Latitude.  When an latitude and longitude is found it is kept under the html class tag Z0LcW.
    'You can verify this by selecting the text in the search and inspecting the element.
    '===============================================================================================================================================
    Set http = CreateObject("MSXML2.XMLHTTP")
    If city <> "" Then
        wp = "https://www.google.com/search?q=" & city & "+" & state_or_country & "longitude+latitude"
    Else
        wp = "https://www.google.com/search?q=" & state_or_country & "longitude+latitude"
    End If
    http.Open "GET", wp, False
    http.send
    html.body.innerHTML = http.responseText
    Set topics = html.getElementsByClassName("Z0LcW")(0)
    If Not topics Is Nothing Then
        lat_long = topics.innerText
    End If

    '===============================================================================================================================================
    'If the google search returns a value for the longitude and latitude find the closest weather station from the ASHRAE tables in the
    'T-Calc User Interface workbook
    '===============================================================================================================================================
    If lat_long <> "" Then
        '===============================================================================================================================================
        'The longitude and latitude comes in as one string such as "39.1031° N, 84.5120° W".  We need to break this up to have:
            'latitude number = 39.1031
            'latitude direction = N
            'longitude number = 84.5120
            'longitude direction = W
        '===============================================================================================================================================
        loc_latitude = Mid(lat_long, 1, InStr(1, lat_long, ",") - 1)
        loc_longitude = Mid(lat_long, InStr(1, lat_long, ", ") + 2)
        
        loc_latitude_dir = Trim(Mid(loc_latitude, InStr(1, loc_latitude, " ") + 1))
        loc_longitude_dir = Trim(Mid(loc_longitude, InStr(1, loc_latitude, " ") + 1))
        loc_latitude_num = Mid(loc_latitude, 1, InStr(1, loc_latitude, "°") - 1)
        loc_longitude_num = Mid(loc_longitude, 1, InStr(1, loc_latitude, "°") - 1)


        '===============================================================================================================================================
        'Open the T-Calc User Interface workbook and read all of the ASHRAE data into an array to speed up processing, close the workbook when finished
        '===============================================================================================================================================
        file_path = "I:\engineering\Thermal Calculator\Thermal Calculator v2.00\"
        file_name = "T-Calc User Interface.xlsm"
        ws_folder_path = "R:\NEW R DRIVE\ASHRAE Weather Data\STATIONS\"
        Workbooks.Open _
            Filename:=file_path & file_name, _
            ReadOnly:=True
        ui_wb = ActiveWorkbook.Name
        If formTCalc.chboxIntl.Value = False Then
            With Workbooks(ui_wb).Sheets("Weather Station (US)")
                ui_wb_ws_last_row = .Cells(.Rows.count, 1).End(xlUp).Row
                ui_wb_ws_last_col = .Cells(1, .Columns.count).End(xlToLeft).Column
                ui_wb_ws_arr = .Range(.Cells(1, 1), .Cells(ui_wb_ws_last_row, ui_wb_ws_last_col))
            End With
            Workbooks(ui_wb).Close SaveChanges:=False
        Else
            With Workbooks(ui_wb).Sheets("Weather Station(Intl)")
                ui_wb_ws_last_row = .Cells(.Rows.count, 1).End(xlUp).Row
                ui_wb_ws_last_col = .Cells(1, .Columns.count).End(xlToLeft).Column
                ui_wb_ws_arr = .Range(.Cells(1, 1), .Cells(ui_wb_ws_last_row, ui_wb_ws_last_col))
            End With
            Workbooks(ui_wb).Close SaveChanges:=False
        End If

        '===============================================================================================================================================
        'Identify the columns that the WMO (used to find the pdf of the weather station), Latitude and Longitude are kept
        '===============================================================================================================================================
        For col_counter_1 = LBound(ui_wb_ws_arr, 2) To UBound(ui_wb_ws_arr, 2)
            If ui_wb_ws_arr(1, col_counter_1) = "Latitude(°)" Then
                ui_wb_ws_col_latitude = col_counter_1
            ElseIf ui_wb_ws_arr(1, col_counter_1) = "Longitude(°)" Then
                ui_wb_ws_col_longitude = col_counter_1
            ElseIf ui_wb_ws_arr(1, col_counter_1) = "WMO #" Then
                ui_wb_ws_col_sn = col_counter_1
            End If
        Next

        '===============================================================================================================================================
        'Go through each row of the ASHRAE data.  Compare each weather station longitude and latitude to the values returned by the google search
        'to find the minimum distance to the closest weather station.  This is done through triangulation or the pythagorean theorem.  If the distance
        'of the current weather station is closer than the previous minimum distance then make sure that weather station has an available pdf
        'to pull from. For some reason we do not have PDFs for all of the weather stations listed
        '===============================================================================================================================================
        min_diff = 1000
        For row_counter_1 = LBound(ui_wb_ws_arr, 1) + 1 To UBound(ui_wb_ws_arr, 1)
            ws_latitude = ui_wb_ws_arr(row_counter_1, ui_wb_ws_col_latitude)
            ws_longitude = ui_wb_ws_arr(row_counter_1, ui_wb_ws_col_longitude)
            ws_latitude_num = CDbl(Mid(ws_latitude, 1, Len(ws_latitude) - 1))
            ws_latitude_dir = Right(ws_latitude, 1)
            ws_longitude_num = CDbl(Mid(ws_longitude, 1, Len(ws_longitude) - 1))
            ws_longitude_dir = Right(ws_longitude, 1)

            If ws_latitude_dir = loc_latitude_dir And ws_longitude_dir = loc_longitude_dir Then
                lat_diff = Abs(loc_latitude_num - ws_latitude_num)
                long_diff = Abs(loc_longitude_num - ws_longitude_num)
                dist_diff = (((lat_diff) ^ 2) + ((long_diff) ^ 2)) ^ (1 / 2)
                If dist_diff < min_diff Then
                    ws_file_name = ui_wb_ws_arr(row_counter_1, ui_wb_ws_col_sn) & "_p*"
                    ws_file_exists = Dir(ws_folder_path & ws_file_name)
                    If Len(ws_file_exists) > 0 Then
                        min_diff = dist_diff
                        min_diff_row = row_counter_1
                        weather_station = ui_wb_ws_arr(row_counter_1, ui_wb_ws_col_sn)
                    End If
                End If
            End If
        Next
        '===============================================================================================================================================
        'Write the weather station to the cmbWeatherStation combo box
        '===============================================================================================================================================
        formTCalc.cmbWeatherStation.Caption = weather_station

        '===============================================================================================================================================
        'Find the file name of the closest weather station
        'Open the pdf using FollowHyperlink
        'Extract all the text from the file with the extract_adobe subroutine and paste into a new worksheet in the Excel workbook
        '===============================================================================================================================================
        ws_file_exists = Dir(ws_folder_path & weather_station & "_p*")
        ActiveWorkbook.FollowHyperlink (ws_folder_path & ws_file_exists)
        extract_adobe

        '===============================================================================================================================================
        'Find all the data that was extracted from the pdf
        '===============================================================================================================================================
        With Workbooks(sb).Sheets("ASHRAE data")
            pdf_last_row = .Cells(.Rows.count, 1).End(xlUp).Row
            pdf_last_col = .Cells(1, .Columns.count).End(xlToLeft).Column
            pdf_arr = .Range(.Cells(1, 1), .Cells(pdf_last_row, pdf_last_col))
        End With

        '===============================================================================================================================================
        'Run through the extracted pdf text to find where Annual Cooling is. Loop through the columns searching the text 2 rows below the row where
        'Annual Cooling was found to find text that says "9A". When that is found look one row under that and that SHOULD be the 0.4% Dry Bulb Temp
        '===============================================================================================================================================
        For pdf_row_count = LBound(pdf_arr, 1) To UBound(pdf_arr, 1)
            If pdf_arr(pdf_row_count, 1) = "Annual" And pdf_arr(pdf_row_count, 2) = "Cooling" Then
                For pdf_col_count = LBound(pdf_arr, 2) To UBound(pdf_arr, 2)
                    If UCase(pdf_arr(pdf_row_count + 2, pdf_col_count)) = "9A" Then
                        dry_bulb_temp = pdf_arr(pdf_row_count + 3, pdf_col_count)
                        GoTo WriteForm
                    End If
                Next
            End If
        Next
        '===============================================================================================================================================
        'If a dry bulb temp is found with the above method, write it to txtOutsideAmbient, Activate the original sheet, and turn screen updating on.
        '===============================================================================================================================================
WriteForm:
        If dry_bulb_temp <> 0 Then
            formTCalc.txtOutsideAmbient.Value = dry_bulb_temp
        End If
        Sheets("Thermal Calculator - Input ").Activate
        Application.ScreenUpdating = True
    Else
        MsgBox ("Location not found through Google")
    End If
    NUM_On
    formTCalc.Show
End Sub
Private Sub NUM_On()  'Turn NUM-Lock on
    Application.Wait Now + TimeSerial(0, 0, 0.5)    'Wait 0.5 seconds
    If Not (GetKeyState(vbKeyNumlock) = 1) Then
      keybd_event VK_NUMLOCK, 1, 0, 0
      keybd_event VK_NUMLOCK, 1, KEYEVENTF_KEYUP, 0
    End If
End Sub
Private Sub extract_adobe()
    Dim wks As Worksheet
    '===============================================================================================================================================
    'Use send keys to select all the text, copy it to the clipboard and close the Adobe Reader
    '===============================================================================================================================================
    Application.ScreenUpdating = False
    Application.Wait Now + TimeSerial(0, 0, 1)  'Wait one second
    Application.SendKeys ("^a")     'Highlight all text
    Application.SendKeys ("^c")     'Copy all text
    Application.Wait Now + TimeSerial(0, 0, 1)  'Wait one second
    Application.SendKeys ("^q")     'Close Adobe Reader
    Application.Wait Now + TimeSerial(0, 0, 0.5)    'Wait 0.5 seconds

    '===============================================================================================================================================
    'Activate Excel.  If a sheet name ASHRAE data exists, delete it
    '===============================================================================================================================================
    Application.Visible = True
    AppActivate "Microsoft Excel"   'Activate excel again
    For Each wks In Worksheets
        If wks.Name Like "ASHRAE Data" Then
            Application.DisplayAlerts = False
            Sheets("ASHRAE data").Delete
            Application.DisplayAlerts = True
        End If
    Next
    '===============================================================================================================================================
    'Add a new excel sheet name ASHRAE data and paste in the Adobe Text in the clipboard, and format it
    '===============================================================================================================================================
    Sheets.Add
    ActiveSheet.Name = "ASHRAE Data"
    ActiveSheet.PasteSpecial Format:="Unicode Text", link:=False, _
        DisplayAsIcon:=False, NoHTMLFormatting:=True
    Application.Wait Now + TimeSerial(0, 0, 1)  'Wait one second
    If Range("B5").Value = "" Then
        Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=True, _
            Semicolon:=False, Comma:=True, Space:=True ', Other:=False, FieldInfo
    End If
    With Range("A:Z")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
        .Columns.EntireColumn.AutoFit
    End With
    Application.ScreenUpdating = True
End Sub
Private Sub cmbWeatherStation_Click()
    Dim folder_path As String
    Dim file_name As String
    Dim file_exists As String

    '===============================================================================================================================================
    'This code will open the Weather Data sheet to extract the Outside Ambient where the unit is going
    'The Caption stored here is populated during the txtCustomerLocationState_AfterUpdate Code by running a triangulation (pythagorean theorem) on
    'the longitude and latitude found through a google search on the values entered in the the customer location/region
    '===============================================================================================================================================

    folder_path = "R:\NEW R DRIVE\ASHRAE Weather Data\STATIONS\"
    file_name = cmbWeatherStation.Caption & "_p*"

    file_exists = Dir(folder_path & file_name)
    '===============================================================================================================================================
    'If a file is found under the above folder path then the length of the file name will be greater than 0.  If the file exists then use the
    'FollowHyperlink function to open the pdf for viewing
    '===============================================================================================================================================
    If Len(file_exists) > 0 Then
        ActiveWorkbook.FollowHyperlink (folder_path & file_exists)
    End If
End Sub
'===============================================================================================================================================
'Transition Information Controls
'===============================================================================================================================================
Private Sub sbTempUnit1_Change()
    temp_change
End Sub
Private Sub sbTempUnit2_Change()
    temp_change
End Sub
Private Sub sbTempUnit3_Change()
    temp_change
End Sub
Private Sub temp_change()
    '===============================================================================================================================================
    'This code is tied to the spin button in the transition information.   It allows the user to select °C or °F for the transition profiles
    'The code changes both of the text boxes to °C or °F to prevent the user from entering two different scales
    '===============================================================================================================================================
    If formTCalc.txtTempUnit1.Text = "°C" Then
        formTCalc.txtTempUnit1.Text = "°F"
        formTCalc.txtTempUnit2.Text = "°F"
        formTCalc.txtTempUnit3.Text = "°F"
    Else
        formTCalc.txtTempUnit1.Text = "°C"
        formTCalc.txtTempUnit2.Text = "°C"
        formTCalc.txtTempUnit3.Text = "°C"
    End If
End Sub
Private Sub sbTransitionType_Change()
    '===============================================================================================================================================
    'Switch between Rate of Change and Time for Transition Type
    '===============================================================================================================================================
    Select Case formTCalc.txtTransitionType
        Case "Rate of Change"
            formTCalc.txtTransitionType.Text = "Time"
            formTCalc.txtTransitionTypeUnits.Text = "Minutes"
        Case "Time"
            formTCalc.txtTransitionType.Text = "Rate of Change"
            formTCalc.txtTransitionTypeUnits.Text = "°C/Min"
    End Select
End Sub
Private Sub sbTransitionTypeUnits_Change()
    '===============================================================================================================================================
    'This code allows the user to change the Rate of Change from °C/min to °F/min.  If the Transition Type is in Time, the user is only permitted
    'to select minutes
    '===============================================================================================================================================
    Select Case formTCalc.txtTransitionType
        Case "Rate of Change"
            If formTCalc.txtTransitionTypeUnits.Text = "°C/Min" Then
                formTCalc.txtTransitionTypeUnits.Text = "°F/Min"
            Else
                formTCalc.txtTransitionTypeUnits.Text = "°C/Min"
            End If
        Case "Time"
            formTCalc.txtTransitionTypeUnits.Text = "Minutes"
    End Select
End Sub
Private Sub cmbAddProfile_Click()
    Dim st As Variant
    Dim et As Variant
    Dim transition_time As Integer

    '===============================================================================================================================================
    'Take the starting and ending temperature as well as the transition information and move it to the text box to keep a list of all the profiles
    'This will be used later on to run all of the customer profiles in a loop to find the "worst case scenario"
    '===============================================================================================================================================
    st = formTCalc.txtStartTemp.Value
    et = formTCalc.txtEndTemp.Value
    transition_time = formTCalc.txtTransitionInterval.Value
    '===============================================================================================================================================
    'Add the above values to the listbox, once the values are added clear all of the values from the Userform and return to the "Start Temp"
    'Box to add more profiles
    '===============================================================================================================================================
    If IsNumeric(st) = True And IsNumeric(et) = True And IsNumeric(transition_time) = True Then
        With formTCalc.lbProfiles
            .AddItem st & txtTempUnit1.Text
            .List(.ListCount - 1, 1) = et & txtTempUnit2.Text
            If txtTransitionTypeUnits.Text = "Minutes" Then
                .List(.ListCount - 1, 2) = transition_time & "m"
            Else
                .List(.ListCount - 1, 2) = transition_time & txtTransitionTypeUnits.Text
            End If
        End With
        With formTCalc
            .txtStartTemp.Value = ""
            .txtEndTemp.Value = ""
            .txtTransitionInterval = ""
        End With
    Else
        MsgBox ("Transitions are not input as numbers")
        Exit Sub
    End If
    formTCalc.txtStartTemp.SetFocus
End Sub
Private Sub lbProfiles_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyDelete Then
        formTCalc.lbProfiles.RemoveItem (formTCalc.lbProfiles.ListIndex)    'Remove the Item from the list
    End If
End Sub
'===============================================================================================================================================
'Chamber Construction Controls
'===============================================================================================================================================
Private Sub txtChamberDepth_AfterUpdate()
    chamber_volume
End Sub
Private Sub sbChamberDepth_Change()
    chamber_units
End Sub
Private Sub txtChamberWidth_AfterUpdate()
    chamber_volume
End Sub
Private Sub sbChamberWidth_Change()
    chamber_units
End Sub
Private Sub txtChamberHeight_AfterUpdate()
    chamber_volume
End Sub
Private Sub sbChamberHeight_Change()
    chamber_units
End Sub
Private Sub chamber_volume()
    '===============================================================================================================================================
    'This code changes the calculated volume equation to match the units.  The year is 2018 and the US has yet to embrace the metric system a
    'and may never do so, so it is currently unknown if this will ever be needed.
    '===============================================================================================================================================
    If formTCalc.txtChamberDepthUnits.Text = "in." Then
        formTCalc.txtChamberVolumeUnits.Value = "ft."
        If formTCalc.txtChamberDepth.Value <> "" And formTCalc.txtChamberWidth.Value <> "" And formTCalc.txtChamberHeight.Value <> "" Then
            formTCalc.txtChamberVolume.Value = Format((formTCalc.txtChamberDepth.Value / 12) * (formTCalc.txtChamberWidth.Value / 12) * (formTCalc.txtChamberHeight.Value / 12), "#,##0.00")
        Else
            formTCalc.txtChamberVolume.Value = ""
        End If
    Else
        formTCalc.txtChamberVolumeUnits.Value = "m."
        If formTCalc.txtChamberDepth.Value <> "" And formTCalc.txtChamberWidth.Value <> "" And formTCalc.txtChamberHeight.Value <> "" Then
            formTCalc.txtChamberVolume.Value = Format((formTCalc.txtChamberDepth.Value * 0.01) * (formTCalc.txtChamberWidth.Value * 0.01) * (formTCalc.txtChamberHeight.Value * 0.01), "#,##0.00")
        Else
            formTCalc.txtChamberVolume.Value = ""
        End If
    End If
End Sub
Private Sub chamber_units()
    '===============================================================================================================================================
    'This code allows the user to switch the units on the chamber construction from inches to cm.  It also changes the calculated volume equation
    'to match the units.  The year is 2018 and the US has yet to embrace the metric system and may never do so, so it is currently unknown if this
    'will ever be needed.
    '===============================================================================================================================================
    If formTCalc.txtChamberDepthUnits.Text = "in." Then
        formTCalc.txtChamberDepthUnits.Text = "cm."
        formTCalc.txtChamberWidthUnits.Text = "cm."
        formTCalc.txtChamberHeightUnits.Text = "cm."
        formTCalc.txtChamberVolumeUnits.Value = "m."
        If formTCalc.txtChamberDepth.Value <> "" And formTCalc.txtChamberWidth.Value <> "" And formTCalc.txtChamberHeight.Value <> "" Then
            formTCalc.txtChamberVolume.Value = Format((formTCalc.txtChamberDepth.Value * 0.01) * (formTCalc.txtChamberWidth.Value * 0.01) * (formTCalc.txtChamberHeight.Value * 0.01), "#,##0.00")
        Else
            formTCalc.txtChamberVolume.Value = ""
        End If
    Else
        formTCalc.txtChamberDepthUnits.Text = "in."
        formTCalc.txtChamberWidthUnits.Text = "in."
        formTCalc.txtChamberHeightUnits.Text = "in."
        formTCalc.txtChamberVolumeUnits.Value = "ft."
        If formTCalc.txtChamberDepth.Value <> "" And formTCalc.txtChamberWidth.Value <> "" And formTCalc.txtChamberHeight.Value <> "" Then
            formTCalc.txtChamberVolume.Value = Format((formTCalc.txtChamberDepth.Value / 12) * (formTCalc.txtChamberWidth.Value / 12) * (formTCalc.txtChamberHeight.Value / 12), "#,##0.00")
        Else
            formTCalc.txtChamberVolume.Value = ""
        End If
    End If
End Sub
'===============================================================================================================================================
'Plenum Controls
'===============================================================================================================================================
Private Sub cmbPlenumType_AfterUpdate()

    Dim file_path As String
    Dim file_name As String
    Dim ui_wb As String
    Dim ui_ws As String
    
    Dim last_row As Integer
    Dim last_col As Integer
    Dim row_count As Integer
    Dim col_count As Integer

    Dim plenum_type As String
    Dim blade_material As String
    Dim non_sparking As Boolean

    Dim master_collection As Collection
    Set master_collection = New Collection
    
    Dim ss_plenum_dict As Object
    Set ss_plenum_dict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    plenum_type = formTCalc.cmbPlenumType.Value
    non_sparking = formTCalc.cbNonSparkingBlower.Value

    If non_sparking = True Then
        blade_material = "Al"
    Else
        blade_material = "SS"
    End If
    '===============================================================================================================================================
    'Open the T-Calc User Interface workbook and gather data
    '===============================================================================================================================================
    gather_data master_collection

    '===============================================================================================================================================
    'Get Plenum data to insert the Blower and Motor Type
    '===============================================================================================================================================
    pop_ui_dict master_collection(6), plenum_type, "FAN MATERIAL", ss_plenum_dict, blade_material

    With formTCalc
        .txtBlower.Value = ss_plenum_dict("Fan Blade")
        .txtMotor.Value = ss_plenum_dict("Motor")
    End With
    Application.ScreenUpdating = True
End Sub
Private Sub cbNonSparkingBlower_Click()
    cmbPlenumType_AfterUpdate
End Sub
'===============================================================================================================================================
'Run Button
'===============================================================================================================================================
Private Sub cmbRun_Click()
    Dim source_book As String
    Dim source_sheet As String
    Dim ui_wb As String
    Dim file_path As String
    Dim file_name As String
    Dim msg As String

    Dim numeric_coll As Collection
    Set numeric_coll = New Collection
    
    Dim empty_value_coll As Collection
    Set empty_value_coll = New Collection
    
    Dim collection_item As Variant
    Dim master_collection As Collection
    Set master_collection = New Collection

    control_loop empty_value_coll, numeric_coll
    If empty_value_coll.count > 0 Then
        msg = "The following items do not have valid inputs to run the program:" & vbNewLine
        For Each collection_item In empty_value_coll
            msg = msg & vbTab & collection_item & vbNewLine
        Next collection_item
        MsgBox (msg)
        Exit Sub
    End If
    '===============================================================================================================================================
    'Run the program.  Store the T-Calc workbook as a variable because we will be changing focus to another workbook at points in the code
    '===============================================================================================================================================
    source_book = ActiveWorkbook.Name
    source_sheet = ActiveSheet.Name

    gather_data master_collection

    '===============================================================================================================================================
    'Activate the T-Calc workbook
    '===============================================================================================================================================
    Workbooks(source_book).Activate

    '===============================================================================================================================================
    'Add customer information from T-Calc to Excel Sheet
    '===============================================================================================================================================
    write_customer_info source_book, source_sheet

    '===============================================================================================================================================
    'Add Chamber Data
    '===============================================================================================================================================
    chamber_data source_book, source_sheet, master_collection
    
    '===============================================================================================================================================
    'Add Plenum Data
    '===============================================================================================================================================
    plenum_data source_book, source_sheet, master_collection

    '===============================================================================================================================================
    'Add transition summary profiles
    '===============================================================================================================================================
    transition_summary source_book, source_sheet

End Sub
Private Sub control_loop(empty_value_coll, numeric_coll)
    Dim form_controls As Control
    Dim type_name As String
    Dim name_value As Variant

    Dim input_dict As Object
    Set input_dict = CreateObject("Scripting.Dictionary")

    '===============================================================================================================================================
    'Create a dictionary of easy to read inputs for the user
    '===============================================================================================================================================
    form_inputs_dict input_dict
    pop_numeric_collecion numeric_coll

    '===============================================================================================================================================
    'Loop through each control.  If the control is a TextBox or ComboBox, the value is Blank, and it is not part of the transition information
    'Add the value to the collection.  If the value should be a number and is not, add it to the collection
    '===============================================================================================================================================
    For Each form_controls In formTCalc.Controls
        type_name = TypeName(form_controls)
        If TypeName(form_controls) = "TextBox" Or TypeName(form_controls) = "ComboBox" Then
            If form_controls.Value = "" And _
                (form_controls.Name <> "txtStartTemp" And form_controls.Name <> "txtEndTemp" And form_controls.Name <> "txtTransitionInterval") Then
                
                empty_value_coll.Add input_dict(form_controls.Name)
            
            ElseIf Exists(numeric_coll, Trim(form_controls.Name)) Then
                If IsNumeric(form_controls.Value) = False Then
                    empty_value_coll.Add input_dict(form_controls.Name)
                End If
            End If
        End If
    Next
End Sub
Private Sub form_inputs_dict(input_dict)
    input_dict.Add "txtCustomerName", "Customer Name"
    input_dict.Add "txtQuote", "Quote"
    input_dict.Add "cmbVoltageCode", "Voltage Code"
    input_dict.Add "txtCustomerLocationCity", "Customer Location (City)"
    input_dict.Add "txtCustomerLocationState", "Customer Location (State)"
    input_dict.Add "txtElevation", "Elevation"
    input_dict.Add "txtOutsideAmbient", "Outside Ambient"
    input_dict.Add "txtCustomerAmbient", "Customer Ambient"
    input_dict.Add "txtHighTemp", "High Temp"
    input_dict.Add "txtChamberDepth", "Chamber Depth"
    input_dict.Add "txtChamberWidth", "Chamber Width"
    input_dict.Add "txtChamberHeight", "Chamber Height"
    input_dict.Add "txtChamberVolume", "Chamber Volume"
    input_dict.Add "cmbChamberManufacturer", "Chamber Manufacturer"
    input_dict.Add "cmbPlenumType", "Plenum Type"
End Sub
Private Sub pop_numeric_collecion(numeric_coll)
    numeric_coll.Add "txtHighTemp"
    numeric_coll.Add "txtChamberDepth"
    numeric_coll.Add "txtChamberHeight"
    numeric_coll.Add "txtChamberWidth"
    numeric_coll.Add "txtChamberVolume"
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
            arr = .Range(.Cells(1, 2), .Cells(last_row, last_col))
        End With
        master_collection.Add arr
    Next

    '====================================================================================================================
    'After all data has been gathered into a large array made up of all the smaller arrays assign it the the gather_data
    'function and close the User Interface workbook
    '====================================================================================================================
    Workbooks(ui_wb).Close SaveChanges:=False
End Sub
Function Exists(coll, key) As Boolean
    Dim coll_count As Integer
    For coll_count = 1 To coll.count
        If coll(coll_count) = key Then
            Exists = True
            Exit For
        End If
    Next
End Function
'===============================================================================================================================================
'Unload Customer Info
'===============================================================================================================================================
Private Sub write_customer_info(ByVal source_book, ByVal source_sheet)
    '===============================================================================================================================================
    'Add customer information from T-Calc to Excel Sheet
    '===============================================================================================================================================
    With formTCalc
        Workbooks(source_book).Sheets(source_sheet).Range("E2").Value = StrConv(.txtCustomerName.Value, vbProperCase)
        Workbooks(source_book).Sheets(source_sheet).Range("E3").Value = .txtQuote.Value
        Workbooks(source_book).Sheets(source_sheet).Range("E4").Value = .cmbVoltageCode.Value
        If .txtCustomerLocationCity.Value <> "" Then
            Workbooks(source_book).Sheets(source_sheet).Range("E5").Value = StrConv(.txtCustomerLocationCity.Value & ", " & .txtCustomerLocationState.Value, vbProperCase)
        Else
            Workbooks(source_book).Sheets(source_sheet).Range("E5").Value = StrConv(.txtCustomerLocationState.Value, vbProperCase)
        End If
        Workbooks(source_book).Sheets(source_sheet).Range("E6").Value = .txtElevation.Value
        Workbooks(source_book).Sheets(source_sheet).Range("E7").Value = .txtOutsideAmbient.Value
        Workbooks(source_book).Sheets(source_sheet).Range("E8").Value = .txtCustomerAmbient.Value
    End With
End Sub
'===============================================================================================================================================
'Unload Chamber Info
'===============================================================================================================================================
Private Sub chamber_data(sb, ss, master_collection)
    
    Dim chamber_manufacturer As String
    Dim high_temp As String

    Dim cm_dict As Object
    Set cm_dict = CreateObject("Scripting.Dictionary")

    Dim wall_thickness As Integer

    '===============================================================================================================================================
    'Assign the Depth, Width and Height of the chamber to the appropriate cells
    '===============================================================================================================================================
    With formTCalc
        Workbooks(sb).Sheets(ss).Range("C24").Value = .txtChamberDepth.Value
        Workbooks(sb).Sheets(ss).Range("D24").Value = .txtChamberWidth.Value
        Workbooks(sb).Sheets(ss).Range("E24").Value = .txtChamberHeight.Value
    End With

    '===============================================================================================================================================
    'Determine the thickness of the chamber walls.  If the highest temperature of the box is greater than 86°C then the thickness is 5", anything
    'below that and it is 4"
    '===============================================================================================================================================
    Select Case txtTempUnit3.Value
        Case "°C"
            Select Case txtHighTemp.Value
                Case Is < 86
                    wall_thickness = 4
                Case Is >= 86
                    wall_thickness = 5
            End Select
        Case "°F"
            Select Case txtHighTemp.Value
                Case Is < 186.8
                    wall_thickness = 4
                Case Is >= 186.8
                    wall_thickness = 5
            End Select
    End Select

    '===============================================================================================================================================
    'Populate the Chamber Manufacturer Dictionary
    '===============================================================================================================================================
    pop_ui_dict master_collection(3), formTCalc.cmbChamberManufacturer.Text, "WALL THICKNESS (IN)", cm_dict, wall_thickness

    If cm_dict.count = 0 Then
        Exit Sub
    End If
    '===============================================================================================================================================
    'Take all the values from the chamber manufacturer dictionary and put them in to the T-Calc workbook
    '===============================================================================================================================================
    unload_chamber_dictionary cm_dict, sb, ss, master_collection
End Sub
Private Sub unload_chamber_dictionary(cm_dict, sb, ss, master_collection)
    Dim k_factor As Double
    Dim wall_thickness As Integer
    Dim n As Variant

    Dim ss_sm_walls_dict As Object
    Set ss_sm_walls_dict = CreateObject("Scripting.Dictionary")

    Dim ss_sm_floors_dict As Object
    Set ss_sm_floors_dict = CreateObject("Scripting.Dictionary")
    '===============================================================================================================================================
    'Take all the values from the chamber manufacturer dictionary and put them in to the T-Calc workbook
    '===============================================================================================================================================
    With Workbooks(sb).Sheets(ss)
        .Range("G29").Value = cm_dict("Wall Thickness (in)")
        .Range("Q25").Value = cm_dict("Insulation Type")
        .Range("Q27").Value = cm_dict("Insulation Density (lbs/cubic foot)")
        k_factor = cm_dict("Insulation (K factor)") / cm_dict("Wall Thickness (in)")
        .Range("Q29").GoalSeek Goal:=k_factor, ChangingCell:=Range("Q26")
        .Range("F33").Value = cm_dict("Wall & Ceiling Material")
        '===============================================================================================================================================
        'Populate the Sheet Metal Dictionary for wall and ceiling
        '===============================================================================================================================================
        pop_ui_dict master_collection(4), cm_dict("Wall & Ceiling Thickness (ga)"), "GAUGE NUMBER", ss_sm_walls_dict

        .Range("F34").Value = ss_sm_walls_dict("SS Specific Heat (BTU/lb/F)")
        .Range("F35").Value = ss_sm_walls_dict("SS Thermal Conductivity (BTU/in/sq ft/F)")
        .Range("F36").Value = ss_sm_walls_dict("SS Density (lbs/cubic foot)")
        .Range("F37").Value = "Sheet Metal"
        .Range("F38").Value = ss_sm_walls_dict("Gauge Number")
        .Range("F39").Value = ss_sm_walls_dict("Sheet Thickness (in)")
        .Range("F40").Value = ss_sm_walls_dict("Sheet Density (lb/sq ft)")
        .Range("F43").Value = cm_dict("Wall Thickness (in)")
        
        .Range("Q33").Value = cm_dict("Floor Liner Material")
        '===============================================================================================================================================
        'Populate the Sheet Metal Dictionary for floor
        '===============================================================================================================================================
        pop_ui_dict master_collection(4), cm_dict("Floor Liner Thickness(ga)"), "GAUGE NUMBER", ss_sm_floors_dict
        .Range("Q34").Value = ss_sm_floors_dict("SS Specific Heat (BTU/lb/F)")
        .Range("Q35").Value = ss_sm_floors_dict("SS Thermal Conductivity (BTU/in/sq ft/F)")
        .Range("Q36").Value = ss_sm_floors_dict("SS Density (lbs/cubic foot)")
        .Range("Q37").Value = "Sheet Metal"
        .Range("Q38").Value = ss_sm_floors_dict("Gauge Number")
        .Range("Q39").Value = ss_sm_floors_dict("Sheet Thickness (in)")
        .Range("Q40").Value = ss_sm_floors_dict("Sheet Density (lb/sq ft)")
        .Range("Q43").Value = cm_dict("Wall Thickness (in)")
    End With
End Sub

Private Sub pop_ui_dict(ByVal collection_arr, ByVal lookup_row_value, ByVal lookup_col_name, ByRef ui_dict, Optional ByVal lookup_col_value)
    Dim col_count_1 As Integer
    Dim col_count_2 As Integer
    Dim lookup_col As Integer
    Dim lookup_row As Integer

    '===============================================================================================================================================
    'Each chamber manufacturer offers 4" and 5" wall thickness.  Loop through each column to find the "Wall Thickness" column.
    '===============================================================================================================================================
    For col_count_1 = LBound(collection_arr, 2) To UBound(collection_arr, 2)
        If UCase(collection_arr(1, col_count_1)) = lookup_col_name Then
            lookup_col = col_count_1
            Exit For
        End If
    Next

    '===============================================================================================================================================
    'Find the matching row that contains the desired chamber manufacturer and wall thickness
    '===============================================================================================================================================
    
    If IsMissing(lookup_col_value) = False Then
        For lookup_row = LBound(collection_arr, 1) To UBound(collection_arr, 1)
            If collection_arr(lookup_row, 1) = lookup_row_value And collection_arr(lookup_row, lookup_col) = CStr(lookup_col_value) Then
                Exit For
            End If
        Next
    Else
        For lookup_row = LBound(collection_arr, 1) To UBound(collection_arr, 1)
            If CStr(collection_arr(lookup_row, lookup_col)) = lookup_row_value Then
                Exit For
            End If
        Next
    End If
    If lookup_row > UBound(collection_arr, 1) Then
        MsgBox ("No match found for " & lookup_row_value)
        Exit Sub
    End If
    '===============================================================================================================================================
    'Add the Column Header and Values to a Dictionary to be used later
    '===============================================================================================================================================
    For col_count_2 = LBound(collection_arr, 2) To UBound(collection_arr, 2)
        ui_dict.Add Trim(collection_arr(1, col_count_2)), Trim(collection_arr(lookup_row, col_count_2))
    Next
End Sub
'===============================================================================================================================================
'Unload Plenum Info
'===============================================================================================================================================
Private Sub plenum_data(sb, ss, master_collection)

    Dim plenum_data_dict As Object
    Set plenum_data_dict = CreateObject("Scripting.Dictionary")

    Dim fan_blade_dict As Object
    Set fan_blade_dict = CreateObject("Scripting.Dictionary")

    Dim motor_dict As Object
    Set motor_dict = CreateObject("Scripting.Dictionary")

    Dim insulation_dict As Object
    Set insulation_dict = CreateObject("Scripting.Dictionary")

    Dim plenum_ss_dict As Object
    Set plenum_ss_dict = CreateObject("Scripting.Dictionary")

    Dim plenum_type As String
    Dim fan_blade As String
    Dim motor As String
    Dim insulation As String

    With formTCalc
        Workbooks(sb).Sheets(ss).Range("F49").Value = .cmbPlenumType.Value
        plenum_type = .cmbPlenumType.Value
        fan_blade = .txtBlower.Value
        motor = .txtMotor.Value
        insulation = .cmbInsulation.Value
    End With

    '===============================================================================================================================================
    'Gather All Dictionary Information
    '===============================================================================================================================================
    pop_ui_dict master_collection(6), plenum_type, "FAN BLADE", plenum_data_dict, fan_blade
    pop_ui_dict master_collection(7), fan_blade, "PART ID", fan_blade_dict
    pop_ui_dict master_collection(8), motor, "MOTOR", motor_dict
    pop_ui_dict master_collection(9), insulation, "TYPE", insulation_dict
    pop_ui_dict master_collection(4), plenum_data_dict("Sheet Metal (ga)"), "GAUGE NUMBER", plenum_ss_dict
    '===============================================================================================================================================
    'Unload All Data to Sheet
    '===============================================================================================================================================
    unload_plenum_dict plenum_data_dict, sb, ss
    unload_fan_dict fan_blade_dict, sb, ss
    unload_motor_dict motor_dict, sb, ss
    unload_ins_dict insulation_dict, sb, ss
    unload_plenum_ss_dict plenum_ss_dict, sb, ss
End Sub
Private Sub unload_plenum_dict(dict, sb, ss)
    With Workbooks(sb).Sheets(ss)
        .Range("C52").Value = dict("Depth")
        .Range("D52").Value = dict("Width")
        .Range("E52").Value = dict("Height")
        .Range("E58").Value = dict("Air Outlet Area")
        .Range("E59").Value = dict("Air Inlet Area")
        .Range("Q54").Value = dict("Fan Qty")
    End With
End Sub
Private Sub unload_fan_dict(dict, sb, ss)
    With Workbooks(sb).Sheets(ss)
        .Range("Q51").Value = dict("Fan Type")
        .Range("Q52").Value = dict("Specs")
        .Range("Q53").Value = dict("Flow Rate (CFM)")
        .Range("Q55").Value = dict("Weight (lb)")
        .Range("Q57").Value = dict("Material")
        .Range("Q58").Value = dict("Specific Heat (BTU/lb/F)")
    End With
End Sub
Private Sub unload_motor_dict(dict, sb, ss)
    With Workbooks(sb).Sheets(ss)
        .Range("Q62").Value = dict("HP")
        .Range("Q63").Value = dict("RPM")
    End With
End Sub
Private Sub unload_ins_dict(dict, sb, ss)
    With Workbooks(sb).Sheets(ss)
        .Range("F64").Value = dict("Type")
        .Range("F65").Value = dict("Thermal Resistance (BTU/in/sf/F)")
        .Range("F66").Value = dict("Density (lb/cubic feet)")
    End With
End Sub
Private Sub unload_plenum_ss_dict(dict, sb, ss)
    With Workbooks(sb).Sheets(ss)
        .Range("F73").Value = "Stainless Steel"
        .Range("P73").Value = "Stainless Steel"
        .Range("F74").Value = dict("SS Specific Heat (BTU/lb/F)")
        .Range("P74").Value = dict("SS Specific Heat (BTU/lb/F)")
        .Range("F75").Value = dict("SS Thermal Conductivity (BTU/in/sq ft/F)")
        .Range("P75").Value = dict("SS Thermal Conductivity (BTU/in/sq ft/F)")
        .Range("F76").Value = dict("SS Density (lbs/cubic foot)")
        .Range("P76").Value = dict("SS Density (lbs/cubic foot)")
        .Range("F77").Value = "Sheet Metal"
        .Range("P77").Value = "Sheet Metal"
        .Range("F78").Value = dict("Gauge Number")
        .Range("P78").Value = dict("Gauge Number")
        .Range("F79").Value = dict("Sheet Thickness (in)")
        .Range("P79").Value = dict("Sheet Thickness (in)")
        .Range("F80").Value = dict("Sheet Density (lb/sq ft)")
        .Range("P80").Value = dict("Sheet Density (lb/sq ft)")
        .Range("F83").Value = .Range("F43").Value
        .Range("P83").Value = .Range("F43").Value
    End With
End Sub
Private Sub transition_summary(ByVal source_book, ByVal source_sheet)
    '===============================================================================================================================================
    'Add transition summary profiles
    '===============================================================================================================================================
    Dim starting_temp_units As String
    Dim ending_temp_units As String
    Dim transition_units As String

    Dim starting_temp_number As Double
    Dim ending_temp_number As Double
    Dim transition_number As Double
    Dim lb_counter As Integer

    With formTCalc
        For lb_counter = 1 To (.lbProfiles.ListCount - 1)
            split_text starting_temp_number, .lbProfiles.List(lb_counter, 0), starting_temp_units
            split_text ending_temp_number, .lbProfiles.List(lb_counter, 1), ending_temp_units
            split_text transition_number, .lbProfiles.List(lb_counter, 2), transition_units
            Select Case starting_temp_units
                Case "°C"
                    Workbooks(source_book).Sheets(source_sheet).Range("E12").Value = starting_temp_number
                Case "°F"
                    Workbooks(source_book).Sheets(source_sheet).Range("G12").GoalSeek Goal:=starting_temp_number, ChangingCell:=Workbooks(source_book).Sheets(source_sheet).Range("E12")
            End Select
            Select Case ending_temp_units
                Case "°C"
                    Workbooks(source_book).Sheets(source_sheet).Range("E13").Value = ending_temp_number
                Case "°F"
                    Workbooks(source_book).Sheets(source_sheet).Range("G13").GoalSeek Goal:=ending_temp_number, ChangingCell:=Workbooks(source_book).Sheets(source_sheet).Range("E13")
            End Select
            Select Case transition_units
                Case "m"
                    Workbooks(source_book).Sheets(source_sheet).Range("K12").Value = transition_number
                    Workbooks(source_book).Sheets(source_sheet).Range("K13").Value = 0
                Case "°C/Min"
                    Workbooks(source_book).Sheets(source_sheet).Range("K12").Value = 0
                    Workbooks(source_book).Sheets(source_sheet).Range("K13").Value = transition_number
                Case "°F/min"
                    Workbooks(source_book).Sheets(source_sheet).Range("K12").Value = 0
                    Workbooks(source_book).Sheets(source_sheet).Range("E18").GoalSeek Goal:=transition_number, ChangingCell:=Workbooks(source_book).Sheets(source_sheet).Range("K13")
            End Select
        Next
    End With
End Sub
Private Sub split_text(ByRef temp_number, ByVal userform_value, ByRef temp_units)

    Dim str_counter As Integer

    For str_counter = 1 To Len(userform_value)
        If IsNumeric(Mid(userform_value, str_counter, 1)) = False And Mid(userform_value, str_counter, 1) <> "-" And Mid(userform_value, str_counter, 1) <> "." Then
            Exit For
        End If
    Next
    temp_number = Mid(userform_value, 1, str_counter - 1)
    temp_units = Mid(userform_value, str_counter)
End Sub

