Attribute VB_Name = "Module5"
Sub Refresh_Dashboard()

    Dim pc As PivotCache
    Dim ws As Worksheet

    Application.ScreenUpdating = False

    ' Refresh all pivot caches
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc

    ' Optional: Recalculate all formulas
    Application.Calculate

    ' Optional: Activate Dashboard sheet
    On Error Resume Next
    Set ws = Sheets("Dashboard")
    If Not ws Is Nothing Then ws.Activate
    On Error GoTo 0

    Application.ScreenUpdating = True

End Sub

