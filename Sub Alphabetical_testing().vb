    Sub Alphabetical_testing()
For Each ws In Worksheets
    Dim worksheetName As String
    Dim lastrow As Long
    Dim VolStock As Long
    Dim Vol As Long
    Dim tableRow As Long
    Dim i As Long
    Dim ticker As String
    Dim yearopen As Double
    Dim yearclose As Double
    VolStock = 0
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, Columns.Count).End(xlToLeft).Column
    tableRow = 2
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly_Change"
    ws.Cells(1, 11).Value = "Percent_Change"
    ws.Cells(1, 12).Value = "Total_Stock_Volume"
    For i = 2 To lastrow
        If yearopen = 0 Then
           yearopen = ws.Cells(i, 3).Value
      End If
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i - 1, 1) = ws.Cells(i, 1) Then
            yearclose = Cells(i, 6).Value
            yearlychange = yearclose - yearopen
            yearlypercent = yearlychange / yearopen
            Vol = ws.Cells(i, 7).Value
            VolStock = VolStock + Vol
            Range("K" & tableRow).NumberFormat = "0.00%"
            Range("K" & tableRow).Value = yearlypercent
            Range("I" & tableRow).Value = ws.Cells(i, 1)
            Range("J" & tableRow).Value = yearlychange
            Range("L" & tableRow).Value = CLng(VolStock)
            tableRow = tableRow + 1
            VolStock = 0
        End If
Next i
Next ws
End Sub