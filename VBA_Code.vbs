
'Option Explicit

Sub Process()

For Each Worksheet In ThisWorkbook.Sheets
        Worksheet.Activate
  
    'Assign names to new cells
    [I1] = "Ticker"
    [J1] = "Yearly Change"
    [K1] = "Percent Change"
    [L1] = "Total Stock Volume"
    [O2] = "Greatest % Increase"
    [O3] = "Greatest % Decrease"
    [O4] = "Greatest Total Volume"
    [P1] = "Ticker"
    [Q1] = "Value"
    
    'Create variables for the [I1]
    'Set an initial variable for holding the ticker symbol
    
    Dim iRowCount As Double: iRowCount = Cells(Rows.Count, "A").End(xlUp).row
    
    Dim iThickerIndex As Integer: iThickerIndex = 2
    Dim CurrentSymbol As String: CurrentSymbol = Cells(2, 1).Value 'ROW 2 COLUMN 1
    
    Dim open_price As Double: open_price = Cells(2, 3).Value
    'Dim OpePrice As Double: OpePrice = Range("C2").Value 'ROW COLUMN C
    Dim close_price As Double
    Dim volume As Double: volume = 0
    Dim iRow As Double: iRow = 2
    For iRow = 2 To (iRowCount + 1)
    
        
        'CHECK IF WE HAVE A NEW SYMBOL
        If Cells(iRow, 1).Value <> CurrentSymbol Then 'WE HAVE A NEW SYMBOL
            Cells(iThickerIndex, 9).Value = CurrentSymbol 'PRINT SYMBOL
            CurrentSymbol = Cells(iRow, 1).Value 'UPDATE CURRENT SYMBOL
            
            'CALCULATE OPEN/CLOSE
            Dim year_change As Double
            Dim per_change As Double
            close_price = Cells(iRow - 1, 6).Value
            year_change = close_price - open_price
            Cells(iThickerIndex, 10).Value = year_change
            
            per_change = (close_price - open_price) / open_price
            Cells(iThickerIndex, 11).Value = per_change
            
            open_price = Cells(iRow, 3).Value ' update the open price for new ticker
            
            Cells(iThickerIndex, 12).Value = volume
            volume = Cells(iRow, 7).Value
            iThickerIndex = iThickerIndex + 1 'SET PRINT CELL TO NEXT ROW
        Else
            volume = volume + Cells(iRow, 7).Value
        End If
    
    Next
    
    Dim GreatestTotalVolume As Double: GreatestTotalVolume = 0
    Dim GreatestIncrease As Double: GreatestIncrease = 0
    Dim GreatestDecrease As Double: GreatestDecrease = 0
    
    iRowCount = Cells(Rows.Count, "L").End(xlUp).row
    For i = 2 To iRowCount
        Dim current_volume As Double
        current_volume = Cells(i, 12).Value
        If current_volume > GreatestTotalVolume Then
            GreatestTotalVolume = current_volume
            Range("P4").Value = Cells(i, 9).Value
            'put ticker in ticker bucket
        End If
        
        If Cells(i, 11).Value > GreatestIncrease Then
            GreatestIncrease = Cells(i, 11).Value
            Range("Q2").Value = Cells(i, 11).Value ' update the value
            Range("P2").Value = Cells(i, 9).Value ' update the ticker
        End If
        
        If Cells(i, 11).Value < GreatestDecrease Then
            GreatestDecrease = Cells(i, 11).Value
            Range("Q3").Value = Cells(i, 11).Value ' update the value
            Range("P3").Value = Cells(i, 9).Value ' update the ticker
        End If
    Next
    Range("Q4").Value = GreatestTotalVolume ' update the value

    For i = 2 To iRowCount
        If Cells(i, 10).Value >= 0.01 Then
            Cells(i, 10).Interior.ColorIndex = 4
        ElseIf Cells(i, 10) < 0.01 Then
            Cells(i, 10).Interior.ColorIndex = 3
        End If
    Next i
    
    Columns("A:Q").AutoFit
        Columns("K:K").NumberFormat = "0.00%"
    Next Worksheet
    
End Sub


