Attribute VB_Name = "Module1"
Sub StockTicker()

Dim total_volume As Double
'total_volume = 0

Dim year_change As Double
Dim perc_change As Double

'keep track of where ticker starts for each stock
Dim openingrow As Double
openingrow = 2

'keep track of the row to print to are
Dim sumrow As Double
sumrow = 2

lastrow = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

    total_volume = total_volume + Cells(i, 7).Value

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'print the total volume
        Cells(sumrow, 12).Value = total_volume
        'print ticker
        Cells(sumrow, 9).Value = Cells(i, 1).Value
        
        'calculate change
        year_change = Cells(i, 6).Value - Cells(openingrow, 3).Value
        'print and format change
        Cells(sumrow, 10).Value = year_change
            If year_change < 0 Then
                Cells(sumrow, 10).Interior.ColorIndex = 3
            Else
                Cells(sumrow, 10).Interior.ColorIndex = 4
            End If
        'calculate percent change
            If Cells(openingrow, 3).Value = 0 Then
                perc_change = 0
                Cells(sumrow, 11).Value = perc_change
                Cells(sumrow, 11).NumberFormat = "0.00%"
            Else
                perc_change = (year_change / Cells(openingrow, 3).Value)
            End If
                    
        'print and format percent change
        Cells(sumrow, 11).Value = perc_change
        Cells(sumrow, 11).NumberFormat = "0.00%"
        
        'update sumrow to be next row down
        sumrow = sumrow + 1
        'resetting total volume amount
        total_volume = 0
        'resetting opening row
        openingrow = i + 1
    End If

Next i

End Sub

