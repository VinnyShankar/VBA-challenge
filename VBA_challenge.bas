Attribute VB_Name = "Module1"
Sub Reset():

    'Reset the worksheet to it's original state
    
        'Clear the cells
        Columns("I:L").Clear
        Columns("P:R").Clear
        
        'Reset the widths of the columns
        Columns("J:L").ColumnWidth = 8.43
        Columns("P").ColumnWidth = 8.43
        
        'Remove the number formats
        Range("J2:J91").ClearFormats
        Range("R2:R3").ClearFormats

End Sub
Sub Stocks():

    'Print the column labels for table one
    
        'Print the Ticker column label
        Range("I1").Value = "Ticker"
        
        'Print the Yearly Change column label
        Range("J1").Value = "Yearly Change"
        
        'Print the Percentage Change column label
        Range("K1").Value = "Percentage Change"
        
        'Print the Total Stock Volume column label
         Range("L1").Value = "Total Stock Volume"
    
    'Populate table one
    
        'Count the number of rows in the dataset
        Dim rowcount As Long
        rowcount = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Keep track of the row in table one
        Dim tickercount As Long
        tickercount = 2
        
        'Store the sum of the volume for the current <ticker>
        'This variable needs to be a Double or the value will overflow since stock volume is a very large integer
        Dim volsum As Double
        volsum = 0
        
        'Store the value of the first <open>
        Dim firstopen As Double
        firstopen = Cells(2, 3).Value
        
        'Store the value of the Greatest % Increase
        Dim greatestincrease As Double
        greatestincrease = 0
        
        'Store the ticker with the Greatest % Increase
        Dim tickerincrease As String
        tickerincrease = 0
        
        'Store the value of the Greatest % Decrease
        Dim greatestdecrease As Double
        greatestdecrease = 0
        
        'Store the ticker with the Greatest % Decrease
        Dim tickerdecrease As String
        tickerdecrease = 0
        
        'Store the value of the Greatest Total Volume
        Dim greatestvol As Double
        greatestvol = 0
        
        'Store the ticker with the Greatest Total Volume
        Dim tickervol As String
        tickervol = Cells(2, 9).Value
    
            'Loop through the rows
            For a = 2 To rowcount
        
                'Add the <vol> of the current row to volume sum
                volsum = volsum + Cells(a, 7).Value
            
                'If the current <ticker> is not the same as the next <ticker>
                If Cells(a, 1).Value <> Cells(a + 1, 1).Value Then
    
                    'Print the current <ticker> in the Ticker column of table one
                    Cells(tickercount, 9).Value = Cells(a, 1).Value
                    
                    'Print the (last <close> - first <open>) in the Yearly Change column of table one and give it two decimal places
                    Cells(tickercount, 10).Value = FormatNumber(Cells(a, 6).Value - firstopen, 2)
                    'Format the value to have Excel's "Number" format
                    Cells(tickercount, 10).NumberFormat = "0.00"
                    
                        'If the Yearly Change is positive
                        If Cells(tickercount, 10).Value > 0 Then
                
                            'Fill the cell with a green color
                            Cells(tickercount, 10).Interior.ColorIndex = 4
                    
                        'If the Yearly Change is negative
                        ElseIf Cells(tickercount, 10).Value < 0 Then
                        
                            'Fill the cell with a red color
                            Cells(tickercount, 10).Interior.ColorIndex = 3
                
                        End If
                    
                    'Print the ((last <close>/first <open>)-1) in the Percentage Change column
                    Cells(tickercount, 11).Value = FormatPercent((Cells(a, 6).Value / firstopen) - 1)
                    
                        'If the next Percentage Change is bigger than the current biggest Percentage Change
                        If Cells(tickercount, 11).Value > greatestincrease Then
                        
                            'Store the value of the bigger Percentage Change
                            greatestincrease = Cells(tickercount, 11).Value
                            
                            'Store the Ticker with the bigger Percentage Change
                            tickerincrease = Cells(tickercount, 9).Value
                        
                        'But if the next Percentage Change is smaller than the current smallest Percentage Change
                        ElseIf Cells(tickercount, 11).Value < greatestdecrease Then
                        
                            'Store the value of the smaller Percentage Change
                            greatestdecrease = Cells(tickercount, 11).Value
                            
                            'Store the Ticker with the smaller Percentage Change
                            tickerdecrease = Cells(tickercount, 9).Value
            
                        End If
                    
                    'Print the volume sum in the Total Stock Volume Column
                    Cells(tickercount, 12).Value = volsum
                    
                        'If the volume sum that was just printed is bigger than the current biggest Total Stock Volume
                        If Cells(tickercount, 12).Value > greatestvol Then
                        
                            'Store the value of that volume sum as the new biggest Total Stock Volume
                            greatestvol = Cells(tickercount, 12).Value
                            
                            'Store the ticker associated with that volume sum
                            tickervol = Cells(tickercount, 9).Value
            
                        End If
                        
                    'Set the firstopen variable to the first <open> of the new symbol
                    firstopen = Cells(a + 1, 3).Value
                    
                    'Reset the volume sum
                    volsum = 0
                    
                    'Shift down by one row in table one
                    tickercount = tickercount + 1
            
                End If
                
            Next a
            
        'Auto adjust the width of the columns in table one
        Columns("J:L").AutoFit
        
    'Print the column labels and row labels for table two
        
        'Print the Ticker column label
        Range("Q1").Value = "Ticker"
        
        'Print the Value column label
        Range("R1").Value = "Value"
        
        'Print the Greatest & Increase row label
        Range("P2").Value = "Greatest % Increase"
        
        'Print the Greatest & Decrease row label
        Range("P3").Value = "Greatest % Decrease"
        
        'Print the Greatest Total Volume row label
        Range("P4").Value = "Greatest Total Volume"
        
        'Auto adjust the width of the column with the row labels in table two
        Columns("P").AutoFit
        
    'Populate table two
        
        'Print the ticker with the Greatest % Increase
        Range("Q2").Value = tickerincrease
        
        'Print the value of the Greatest % Increase
        Range("R2").Value = FormatPercent(greatestincrease)
        
        'Print the ticker with the Greatest % Decrease
        Range("Q3").Value = tickerdecrease
            
        'Print the value of the Greatest % Decrease
        Range("R3").Value = FormatPercent(greatestdecrease)
        
        'Print the ticker with the Greatest Total Volume
        Range("Q4").Value = tickervol
        
        'Print the value of the Greatest Total Volume
        Range("R4").Value = greatestvol

End Sub
