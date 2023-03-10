Attribute VB_Name = "Module1"
Sub StockInfo():

'Print the headers of the new columns
    
    'Print header of Ticker column
    Range("I1").Value = "Ticker"
    'Print header of Yearly Change column
    Range("J1").Value = "Yearly Change"
    'Print header of Percent Change column
    Range("K1").Value = "Percent Change"
    'Print header of Total Stock Volume column
    Range("L1").Value = "Total Stock Volume"
    
'Create a counter to count the number of used rows in the <ticker> column
Dim rowcount As Integer

'Create a counter to count the number of unique tickers in <ticker> column
Dim ucounter As Integer

'Loop through the rows in the <ticker> column
    
    For i = 1 To 22771
        
        'Loop through unique tickers from the <ticker> column
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        'Count the number of unique tickers in <ticker> column
        ucounter = ucounter + 1
        
        'Print the unique tickers in the Ticker column
        Cells(ucounter + 1, 9).Value = Cells(i + 1, 1).Value
        End If
        
    Next i

'Loop through the rows in the Ticker column

    For j = 2 To ucounter
    
    
    
    
    Next j
    
'Repeat these steps for each ticker in the Ticker column

    'Create an array for the first ticker
    'Look up the row with the oldest <date> and get the oldest open value from the <open> column
    'Look up the row with the newest <date> and get the newest close value from the <close> column

    'Calculate the price change: newest close minus oldest open
    'Print the price change in the Yearly Change column

    'Calculate the percent change: (newest close/oldest open) minus 1
    'Print the percent change in the Percent Change column

    'Sum all the values in the <vol> column and print the summed value in the Total Stock Volume column

'Create the summary table

    'Get the largest value from the Percent Change column
    'Print this largest value in the cell: (column Value, row Greatest % Increase)
    'Find the largest value in the Percent Change column
    'In the same row where this value was found, get the ticker from the Ticker column
    'Print the ticker in the cell: (column Ticker, row Greatest % Increase)

    'Get the smallest value from the Percent Change column
    'Print this smallest value in the cell: (column Value, row Greatest % Decrease)
    'Find the smallest value in the Percent Change column
    'In the same row where this value was found, get the ticker from the Ticker column
    'Print the ticker in the cell: (column Ticker, row Greatest % Decrease)

    'Get the largest value in the Total Stock Volume column
    'Print this value in the cell: (column Value, row Greatest Total Volume)
    'Find the largest value in the Total Stock Volume column
    'In the same row where this value was found, get the ticker from the Ticker column
    'Print the ticker in the cell: (column Ticker, row Greatest % Increase)

End Sub
Sub stockinfo2()

'Get the tickers and print the headers

    'Copy <ticker> column into column I
    Range("A:A").Copy Range("I:I")
    'Remove duplicates from Ticker column
    Range("I:I").RemoveDuplicates Columns:=1, Header:=xlYes
    
    'Print header of Ticker column
    Range("I1").Value = "Ticker"
    'Print header of Yearly Change column
    Range("J1").Value = "Yearly Change"
    'Print header of Percent Change column
    Range("K1").Value = "Percent Change"
    'Print header of Total Stock Volume column
    Range("L1").Value = "Total Stock Volume"

'Repeat these steps for each ticker in the Ticker column

    'Create an array for the first ticker
    'Look up the row with the oldest <date> and get the oldest open value from the <open> column
    'Look up the row with the newest <date> and get the newest close value from the <close> column

    'Calculate the price change: newest close minus oldest open
    'Print the price change in the Yearly Change column

    'Calculate the percent change: (newest close/oldest open) minus 1
    'Print the percent change in the Percent Change column

    'Sum all the values in the <vol> column and print the summed value in the Total Stock Volume column

'Create the summary table

    'Get the largest value from the Percent Change column
    'Print this largest value in the cell: (column Value, row Greatest % Increase)
    'Find the largest value in the Percent Change column
    'In the same row where this value was found, get the ticker from the Ticker column
    'Print the ticker in the cell: (column Ticker, row Greatest % Increase)

    'Get the smallest value from the Percent Change column
    'Print this smallest value in the cell: (column Value, row Greatest % Decrease)
    'Find the smallest value in the Percent Change column
    'In the same row where this value was found, get the ticker from the Ticker column
    'Print the ticker in the cell: (column Ticker, row Greatest % Decrease)

    'Get the largest value in the Total Stock Volume column
    'Print this value in the cell: (column Value, row Greatest Total Volume)
    'Find the largest value in the Total Stock Volume column
    'In the same row where this value was found, get the ticker from the Ticker column
    'Print the ticker in the cell: (column Ticker, row Greatest % Increase)

End Sub
End Sub
