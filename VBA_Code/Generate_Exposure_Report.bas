Attribute VB_Name = "Module3"
Sub Generate_Exposure_Report()

    Dim wsCalc As Worksheet, wsExp As Worksheet
    Set wsCalc = Sheets("Calculated_Metrics")
    Set wsExp = Sheets("Exposure_Report")

    wsExp.Cells.Clear
    wsExp.Range("A1:C1").Value = Array("Region", "Total_Exposure", "Trade_Count")

    Dim lastRow As Long
    lastRow = wsCalc.Cells(wsCalc.Rows.Count, 1).End(xlUp).Row

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 2 To lastRow

        Dim region As String
        region = Trim(UCase(wsCalc.Cells(i, 9).Value))   ' Region

        If region <> "" Then

            Dim tradeValue As Double
            tradeValue = CDbl(wsCalc.Cells(i, 7).Value)  ' Trade_Value

            If dict.Exists(region) Then
                dict(region) = Array( _
                    dict(region)(0) + tradeValue, _
                    dict(region)(1) + 1 _
                )
            Else
                dict.Add region, Array(tradeValue, 1)
            End If

        End If

    Next i

    Dim r As Long: r = 2
    Dim k As Variant
    For Each k In dict.Keys
        wsExp.Cells(r, 1).Value = k
        wsExp.Cells(r, 2).Value = dict(k)(0)
        wsExp.Cells(r, 3).Value = dict(k)(1)
        r = r + 1
    Next k

End Sub


