Attribute VB_Name = "Module2"
Sub Calculate_Trade_Metrics()

    Dim wsClean As Worksheet, wsCalc As Worksheet
    Set wsClean = Sheets("Clean_Transactions")
    Set wsCalc = Sheets("Calculated_Metrics")

    ' Clear previous output
    wsCalc.Cells.Clear

    ' Write headers for calculated sheet
    wsCalc.Range("A1:J1").Value = Array( _
        "Trade_ID", _
        "Instrument_Code", _
        "Buy_Sell", _
        "Quantity", _
        "Signed_Quantity", _
        "Price", _
        "Trade_Value", _
        "Desk", _
        "Region", _
        "Trader_Name" _
    )

    ' Identify last row in clean data
    Dim lastRow As Long
    lastRow = wsClean.Cells(wsClean.Rows.Count, "A").End(xlUp).Row

    Dim outputRow As Long
    outputRow = 2

    Dim i As Long
    For i = 2 To lastRow

        ' Read required values from Clean_Transactions
        Dim tradeID As String
        tradeID = wsClean.Cells(i, 1).Value

        Dim instCode As String
        instCode = wsClean.Cells(i, 6).Value

        Dim buySell As String
        buySell = UCase(Trim(wsClean.Cells(i, 7).Value))

        Dim qty As Double
        qty = wsClean.Cells(i, 8).Value

        Dim price As Double
        price = wsClean.Cells(i, 9).Value

        ' Apply BUY / SELL sign logic
        Dim signedQty As Double
        If buySell = "SELL" Then
            signedQty = qty * -1
        Else
            signedQty = qty
        End If

        ' Recalculate Trade Value
        Dim tradeValue As Double
        tradeValue = signedQty * price

        ' Write calculated output
        wsCalc.Cells(outputRow, 1).Value = tradeID
        wsCalc.Cells(outputRow, 2).Value = instCode
        wsCalc.Cells(outputRow, 3).Value = buySell
        wsCalc.Cells(outputRow, 4).Value = qty
        wsCalc.Cells(outputRow, 5).Value = signedQty
        wsCalc.Cells(outputRow, 6).Value = price
        wsCalc.Cells(outputRow, 7).Value = tradeValue
        wsCalc.Cells(outputRow, 8).Value = wsClean.Cells(i, 12).Value   ' Desk
        wsCalc.Cells(outputRow, 9).Value = wsClean.Cells(i, 13).Value   ' Region
        wsCalc.Cells(outputRow, 10).Value = wsClean.Cells(i, 16).Value  ' Trader_Name

        outputRow = outputRow + 1

    Next i

    MsgBox "Trade calculations completed successfully", vbInformation

End Sub


