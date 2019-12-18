Sub itWasVBA_HW()

    Dim ws As Worksheet

    For Each ws In Worksheets

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"

        Dim tickerSymbol As String

        Dim Total_vol As Double
        Total_vol = 0

        Dim counter As Long
        counter = 2

        Dim yearOpen As Double
        yearOpen = 0

        Dim yearClose As Double
        yearClose = 0

        Dim yearChange As Double
        yearChange = 0

        Dim percentChange As Double
        percentChange = 0

        Dim lastrow As Long
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        For i = 2 To lastrow

        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then

        yearOpen = ws.Cells(i, 3).Value

        End If

        Total_vol = Total_vol + ws.Cells(i, 7)

        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then

        ws.Cells(counter, 9).Value = ws.Cells(i, 1).Value

        ws.Cells(counter, 12).Value = Total_vol

        yearClose = ws.Cells(i, 6).Value

        yearChange = yearClose - yearOpen
        ws.Cells(counter, 10).Value = yearChange

        If yearChange >= 0 Then
            ws.Cells(counter, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(counter, 10).Interior.ColorIndex = 3
        End If

        If yearOpen = 0 And yearClose = 0 Then

            percentChange = 0
            ws.Cells(counter, 11).Value = percentChange
            ws.Cells(counter, 11).NumberFormat = "0.00%"

        ElseIf yearOpen = 0 Then

            Dim percentNS As String
            percentNS = "New Stock"
            ws.Cells(counter, 11).Value = percentChange

        Else

            percentChange = yearChange / yearOpen
            ws.Cells(counter, 11).Value = percentChange
            ws.Cells(counter, 11).NumberFormat = "0.00%"

        End If

        counter = counter + 1

        totalVol = 0
        yearOpen = 0
        yearClose = 0
        yearChange = 0
        percentChange = 0

    End If

    Next i

    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row

    Dim bestStock As String
    Dim bestValue As Double

    bestValue = ws.Cells(2, 11).Value

    Dim worstStock As String
    Dim worstValue As Double

    worstValue = ws.Cells(2, 11).Value

    Dim mostVolStock As String
    Dim mostVolValue As Double

    mostVolValue = ws.Cells(2, 11).Value

    For j = 2 To lastrow

        If ws.Cells(j, 11).Value > bestValue Then
        bestValue = ws.Cells(j, 11).Value
        bestStock = ws.Cells(j, 9).Value

    End If

    If ws.Cells(j, 11).Value < worstValue Then
        worstValue = ws.Cells(j, 11).Value
        worstStock = ws.Cells(i, 9).Value

    End If
    Next j

    ws.Cells(2, 16).Value = bestStock
    ws.Cells(2, 17).Value = bestValue
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 16).Value = worstStock
    ws.Cells(3, 17).Value = worstValue
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 16).Value = mostVolStock
    ws.Cells(4, 17).Value = mostVolValue

    ws.Columns("I:L").EntireColumn.AutoFit
    ws.Columns("O:Q").EntireColumn.AutoFit

    Next ws

End Sub
Sub HW()

End Sub
