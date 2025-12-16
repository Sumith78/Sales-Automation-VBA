Attribute VB_Name = "Module6"
Sub Write_Audit_Log(runStatus As String)

    Dim wsLog As Worksheet
    Dim wsRaw As Worksheet
    Dim wsClean As Worksheet
    Dim wsEx As Worksheet

    Set wsLog = Sheets("Audit_Log")
    Set wsRaw = Sheets("Raw_Transactions")
    Set wsClean = Sheets("Clean_Transactions")
    Set wsEx = Sheets("Exception_Report")

    ' Find next empty row in Audit_Log
    Dim logRow As Long
    logRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1

    ' Count raw & clean rows (excluding header)
    Dim rawCount As Long
    Dim cleanCount As Long
    rawCount = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row - 1
    cleanCount = wsClean.Cells(wsClean.Rows.Count, 1).End(xlUp).Row - 1

    ' Count UNIQUE Trade_IDs in Exception_Report
    Dim dictEx As Object
    Set dictEx = CreateObject("Scripting.Dictionary")

    Dim i As Long, lastExRow As Long
    lastExRow = wsEx.Cells(wsEx.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastExRow
        If Trim(wsEx.Cells(i, 1).Value) <> "" Then
            dictEx(Trim(wsEx.Cells(i, 1).Value)) = 1
        End If
    Next i

    Dim exceptionTradeCount As Long
    exceptionTradeCount = dictEx.Count

    ' Write audit log
    wsLog.Cells(logRow, 1).Value = Now
    wsLog.Cells(logRow, 2).Value = rawCount
    wsLog.Cells(logRow, 3).Value = cleanCount
    wsLog.Cells(logRow, 4).Value = exceptionTradeCount
    wsLog.Cells(logRow, 5).Value = runStatus

End Sub

