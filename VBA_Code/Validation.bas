Attribute VB_Name = "Module1"
Sub Validate_And_Clean_Transactions()

    Dim wsRaw As Worksheet, wsClean As Worksheet, wsEx As Worksheet, wsMaster As Worksheet
    Set wsRaw = Sheets("Raw_Transactions")
    Set wsClean = Sheets("Clean_Transactions")
    Set wsEx = Sheets("Exception_Report")
    Set wsMaster = Sheets("Instrument_Master")

    wsClean.Cells.Clear
    wsEx.Cells.Clear

    wsRaw.Rows(1).Copy wsClean.Rows(1)
    wsEx.Range("A1:C1").Value = Array("Trade_ID", "Issue_Type", "Description")

    Dim lastRow As Long, lastMaster As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, "A").End(xlUp).Row
    lastMaster = wsMaster.Cells(wsMaster.Rows.Count, "A").End(xlUp).Row

    Dim cleanRow As Long: cleanRow = 2
    Dim exRow As Long: exRow = 2

    Dim i As Long
    For i = 2 To lastRow

        Dim tradeID As String
        tradeID = Trim(wsRaw.Cells(i, 1).Value)

        Dim buySell As String
        buySell = UCase(Trim(wsRaw.Cells(i, 7).Value))

        Dim qty As Double
        qty = Val(wsRaw.Cells(i, 8).Value)

        Dim price As Double
        price = Val(wsRaw.Cells(i, 9).Value)

        Dim instCode As String
        instCode = wsRaw.Cells(i, 6).Value

        ' ---- VALIDATION RULES ----
        If tradeID = "" Then
            Call LogException(wsEx, exRow, tradeID, "Missing Trade_ID", "Trade ID is blank")
            exRow = exRow + 1

        ElseIf Not IsDate(wsRaw.Cells(i, 2).Value) Then
            Call LogException(wsEx, exRow, tradeID, "Invalid Date", "Trade date is invalid")
            exRow = exRow + 1

        ElseIf buySell <> "BUY" And buySell <> "SELL" Then
            Call LogException(wsEx, exRow, tradeID, "Invalid Buy/Sell", "Buy_Sell value is incorrect")
            exRow = exRow + 1

        ElseIf qty <= 0 Then
            Call LogException(wsEx, exRow, tradeID, "Invalid Quantity", "Quantity must be > 0")
            exRow = exRow + 1

        ElseIf price <= 0 Then
            Call LogException(wsEx, exRow, tradeID, "Invalid Price", "Price must be > 0")
            exRow = exRow + 1

        ElseIf Application.WorksheetFunction.CountIf(wsMaster.Range("A2:A" & lastMaster), instCode) = 0 Then
            Call LogException(wsEx, exRow, tradeID, "Invalid Instrument", "Instrument not found in master")
            exRow = exRow + 1

        Else
            wsRaw.Rows(i).Copy wsClean.Rows(cleanRow)
            cleanRow = cleanRow + 1
        End If

    Next i

    MsgBox "Validation & Cleaning Completed", vbInformation

End Sub



Sub LogException(ws As Worksheet, r As Long, tradeID As String, issue As String, desc As String)
    ws.Cells(r, 1).Value = tradeID
    ws.Cells(r, 2).Value = issue
    ws.Cells(r, 3).Value = desc
End Sub

