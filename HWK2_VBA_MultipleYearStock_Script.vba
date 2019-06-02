Sub stock_analysis()

'Create a loop to cycle through the worksheets
Dim ws As Worksheet

For Each ws In Worksheets

    'Create Summary Table of Stocks
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("I1:L1").Font.Bold = True
    
    'Set variable to hold ticker symbol
    Dim ticker As String
    
    'set variable to hold the row for each ticker symbol in the summary table
    Dim ticker_count As Long
    ticker_count = 2 'first ticker stored on the 2nd row
    
    'set variable to hold the total volume of each ticker
    Dim total_vol As Double
    total_vol = 0
    
    'set variable to count the total rows in each worksheet
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'set variable to hold year open price, year close price, yearly change and percent change for each ticker
    Dim year_open_price As Double
    Dim year_close_price As Double
    Dim year_change As Double
    Dim percent_change As Double
    

    'loop through the table
    
    For i = 2 To lastrow
        'loop through to grab year open for each ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            year_open_price = ws.Cells(i, 3).Value
        End If
    
        total_vol = total_vol + ws.Cells(i, 7).Value
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ws.Cells(ticker_count, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(ticker_count, 12).Value = total_vol
            year_close_price = ws.Cells(i, 6).Value
            year_change = year_close_price - year_open_price
            ws.Cells(ticker_count, 10).Value = year_change
            
            'conditional to format cell
            If year_change >= 0 Then
                ws.Cells(ticker_count, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(ticker_count, 10).Interior.ColorIndex = 3
            End If
            
            
            'calculate percent change
            'conditional for the case of new stock
            If year_open_price = 0 And year_close_price = 0 Then
                percent_change = 0
                ws.Cells(ticker_count, 11).Value = percent_change
                ws.Cells(ticker_count, 11).NumberFormat = "0.00%"
            ElseIf year_open_price = 0 Then
                Dim percent_change_NA As String
                percent_change_NA = "New Stock"
                ws.Cells(ticker_count, 11).Value = percent_change
            Else
                percent_change = year_change / year_open_price
                ws.Cells(ticker_count, 11).Value = percent_change
                ws.Cells(ticker_count, 11).NumberFormat = "0.00%"
            End If
            
            
            'reset
            ticker_count = ticker_count + 1
            total_vol = 0
            year_open_price = 0
            year_close_price = 0
            year_change = 0
            percent_change = 0
            
        End If
        
    Next i
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("O2:O4").Font.Bold = True
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("P1:Q1").Font.Bold = True
    
    lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    Dim best_value As Double
    Dim best_stock As String
    best_value = ws.Cells(2, 11).Value
    
    Dim worst_value As Double
    Dim worst_stock As String
    worst_value = ws.Cells(2, 11).Value
    
    Dim most_vol As Double
    Dim most_stock As String
    most_vol = ws.Cells(2, 12).Value
    
    
    For j = 2 To lastrow
        If ws.Cells(j, 11).Value > best_value Then
            best_value = ws.Cells(j, 11).Value
            best_stock = ws.Cells(j, 9).Value
        End If
        
        If ws.Cells(j, 11).Value < worst_value Then
            worst_value = ws.Cells(j, 11).Value
            worst_stock = ws.Cells(j, 9).Value
        End If
        
        If ws.Cells(j, 12).Value > most_vol Then
            most_vol = ws.Cells(j, 12).Value
            most_stock = ws.Cells(j, 9).Value
        End If
    Next j
    
    ws.Range("P2").Value = best_stock
    ws.Range("P3").Value = worst_stock
    ws.Range("Q2").Value = best_value
    ws.Range("Q3").Value = worst_value
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Range("P4").Value = most_stock
    ws.Range("Q4").Value = most_vol
    
    'Autofit table columns
    ws.Columns("A:G").EntireColumn.AutoFit
    ws.Columns("I:L").EntireColumn.AutoFit
    ws.Columns("O:Q").EntireColumn.AutoFit
         
        
        
 Next ws
    

End Sub



