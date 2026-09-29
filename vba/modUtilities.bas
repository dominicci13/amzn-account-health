Option Explicit

' --- deleteCharts ----------------------------------------------------------
' Clears every shape on the dashboard sheet (Sheet 2) so the Python step can
' paste a fresh set of speed-distribution chart screenshots without colliding
' with last run's images. Sheet activation + SelectAll is required because
' Selection.Delete only operates on the active sheet's selection.
'
' Returns "" on success, or a one-line reason on failure. It must never
' MsgBox: Excel runs hidden, so a modal waits forever on an invisible window
' and the run hangs without ever reaching the crash handler. The Python
' caller raises on a non-empty result.
'
' Performance toggles disable screen updates, automatic calculation, events,
' and alerts during the delete; the original Application state is restored
' in the Cleanup block whether the delete succeeded or raised an error.
' Cleanup snapshots Err BEFORE restoring that state, because a restore that
' itself fails would otherwise overwrite the real reason.
Function deleteCharts() As String
    Dim sh As Worksheet
    Dim prevCalc As Long
    Dim errNum As Long
    Dim errDesc As String

    prevCalc = Application.Calculation

    On Error GoTo Cleanup

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
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo -1  ' end the active handler first, or the Resume Next below is ignored on the error path

    On Error Resume Next
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    On Error GoTo 0

    If errNum <> 0 Then
        deleteCharts = "deleteCharts failed - error " & errNum & ": " & errDesc
    End If
End Function


' --- resizeCharts ----------------------------------------------------------
' Normalizes every chart screenshot on the dashboard sheet (Sheet 2) to the
' same height/width so the layout stays tidy regardless of the source DPI of
' the clipboard image Python pasted in.
'
' Dimensions (104.4 x 375.12) are sized to fit the dashboard cell grid; do
' not change without re-aligning the per-account chart_cell anchors in
' `config/accounts.json -> account_health_dashboard`.
'
' Returns "" on success, or a one-line reason on failure; never a MsgBox
' (see deleteCharts).
Function resizeCharts() As String
    Dim sh As Worksheet
    Dim prevCalc As Long
    Dim errNum As Long
    Dim errDesc As String

    prevCalc = Application.Calculation

    On Error GoTo Cleanup

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
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo -1  ' end the active handler first, or the Resume Next below is ignored on the error path

    On Error Resume Next
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    On Error GoTo 0

    If errNum <> 0 Then
        resizeCharts = "resizeCharts failed - error " & errNum & ": " & errDesc
    End If
End Function
