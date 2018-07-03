Attribute VB_Name = "CrossCheck"
Sub cross_check()

Dim arr1(), arr2(), arr3() As Variant
Dim Loc As String
Dim i, j, n As Integer
ReDim Preserve arr3(2, 1)

'Turn stuff off so the program can run faster
Application.Calculation = xlCalculationManual
Application.ScreenUpdating = False
act_sheet = ActiveSheet.name

arr1 = Workbooks("Master - Smart Compressor Capacity Sheet").Sheets(act_sheet).Range("C6:AL80").Value

'Workbooks.Open _
'    Filename:="C:\Users\cbaker\Desktop\ZF18K4E_cross check - Smart Compressor Capacity Sheet.xlsm", _
'    ReadOnly:=True, _
'    Password:=Workbooks("Master - Smart Compressor Capacity Sheet").Sheets("Compressor Summary").Range("B3").Value
Workbooks.Open _
    Filename:="C:\Users\cbaker\Desktop\3DB3F33KE_cross check - Smart Compressor Capacity Sheet.xlsm", _
    ReadOnly:=True, _
    Password:=Workbooks("Master - Smart Compressor Capacity Sheet").Sheets("Compressor Summary").Range("B3").Value
'Test procedures
'arr2 = Workbooks("ZF18K4E_cross check - Smart Compressor Capacity Sheet.xlsm").Sheets("R404A - 60 Smart ").Range("C6:AL80").Value
'arr2 = Workbooks("ZF18K4E_cross check - Smart Compressor Capacity Sheet.xlsm").Sheets("R404A - 50 Smart").Range("C6:AL80").Value
'arr2 = Workbooks("3DB3F33KE_cross check - Smart Compressor Capacity Sheet.xlsm").Sheets("R404A - 60 Smart ").Range("C6:AL80").Value
arr2 = Workbooks("3DB3F33KE_cross check - Smart Compressor Capacity Sheet.xlsm").Sheets("R404A - 50 Smart").Range("C6:AL80").Value
'arr2 = Workbooks("Master - Smart Compressor Capacity Sheet").Sheets("Sheet3").Range("C6:AL80").Value

'Workbooks("ZF18K4E_cross check - Smart Compressor Capacity Sheet").Close SaveChanges:=False
Workbooks("3DB3F33KE_cross check - Smart Compressor Capacity Sheet").Close SaveChanges:=False
n = 0
For i = 1 To 75
    For j = 1 To 36
        Select Case j
            Case 8, 9, 13, 20
                If Abs(Round(arr1(i, j), 1) - Round(arr2(i, j), 1)) > 2 Then
                    v1 = arr1(i, j)
                    v2 = arr2(i, j)
                    ind = "arr(" & i & ", " & j & ")"
                    arr3(0, n) = ind
                    arr3(1, n) = v1
                    arr3(2, n) = v2
                    n = n + 1
                    ReDim Preserve arr3(2, n)
                End If
            Case Else
                If Round(arr1(i, j), 0.1) <> Round(arr2(i, j), 0.1) Then
                    v1 = arr1(i, j)
                    v2 = arr2(i, j)
                    ind = "arr(" & i & ", " & j & ")"
                    arr3(0, n) = ind
                    arr3(1, n) = v1
                    arr3(2, n) = v2
                    n = n + 1
                    ReDim Preserve arr3(2, n)
                End If
            End Select
    Next j
Next i
'Turn the things on that were previously turned off to improve performance
Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True
ReDim Preserve arr3(2, n - 1)
End Sub

