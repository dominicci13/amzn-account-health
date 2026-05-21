Option Explicit

' --- deleteCharts ----------------------------------------------------------
' Clears every shape on the dashboard sheet (Sheet 2) so the Python step can
' paste a fresh set of speed-distribution chart screenshots without colliding
' with last run's images. Sheet activation + SelectAll is required because
' Selection.Delete only operates on the active sheet's selection.
'
' Performance toggles disable screen updates, automatic calculation, events,
' and alerts during the delete; the original Application state is restored
' in the Cleanup block whether the Sub succeeded or raised an error.
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


' --- resizeCharts ----------------------------------------------------------
' Normalizes every chart screenshot on the dashboard sheet (Sheet 2) to the
' same height/width so the layout stays tidy regardless of the source DPI of
' the clipboard image Python pasted in.
'
' Dimensions (104.4 x 375.12) are sized to fit the dashboard cell grid; do
' not change without re-aligning the per-account chart_cell anchors in
' `config/accounts.json -> account_health_dashboard`.
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
