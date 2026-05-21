Option Explicit


Sub deleteCharts()
    Dim sh As Worksheet
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set sh = ThisWorkbook.Sheets(2)
    sh.Activate
    If sh.Shapes.Count > 0 Then
        sh.Shapes.SelectAll
        Selection.Delete
    End If

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "deleteCharts error " & Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub


Sub resizeCharts()
    Dim sh As Worksheet
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set sh = ThisWorkbook.Sheets(2)
    sh.Activate
    If sh.Shapes.Count > 0 Then
        sh.Shapes.SelectAll
        Selection.ShapeRange.Height = 104.4
        Selection.ShapeRange.Width = 375.12
    End If

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "resizeCharts error " & Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub
