Attribute VB_Name = "EvaluateBudgetHours"
Option Private Module
Option Explicit

Public Sub evaluate_budget_array(ByRef current_ws_array As Variant, ByRef customer_name As Variant, ByRef model_number As Variant, _
                                    ByRef cab_hours As Variant, ByRef electrical_hours As Variant, ByRef refrigeration_hours As Variant)
    Dim array_counter As Integer

    customer_name = current_ws_array(1, 2)
    model_number = current_ws_array(2, 2)

    For array_counter = LBound(current_ws_array, 1) To UBound(current_ws_array, 1)

        If current_ws_array(array_counter, 1) = "Total Cabinetry" Then
            cab_hours = current_ws_array(array_counter, 7)
        ElseIf current_ws_array(array_counter, 1) = "Cabinet Total" Then
            cab_hours = current_ws_array(array_counter, 7)
        ElseIf current_ws_array(array_counter, 1) = "Total Electrical" Then
            electrical_hours = current_ws_array(array_counter, 7)
        ElseIf current_ws_array(array_counter, 1) = "Total Refrigeration" Then
            refrigeration_hours = current_ws_array(array_counter, 7) + refrigeration_hours
        ElseIf current_ws_array(array_counter, 1) = "Total Humidity" Then
            refrigeration_hours = current_ws_array(array_counter, 7) + refrigeration_hours
        End If
    Next array_counter

End Sub
