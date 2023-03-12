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

    'Print the column headers for table one
    
        'Print the Ticker column header
        Range("I1").Value = "Ticker"
        
        'Print the Yearly Change column header
        Range("J1").Value = "Yearly Change"
        
        'Print the Percentage Change column header
        Range("K1").Value = "Percentage Change"
        
        'Print the Total Stock Volume column header
         Range("L1").Value = "Total Stock Volume"
    
    'Print the column headers and row headers for table two
        
        'Print the Ticker column header
        Range("Q1").Value = "Ticker"
        
        'Print the Value column header
        Range("R1").Value = "Value"
        
        'Print the Greatest & Increase row header
        Range("P2").Value = "Greatest % Increase"
        
        'Print the Greatest & Decrease row header
        Range("P3").Value = "Greatest % Decrease"
        
        'Print the Greatest Total Volume row header
        Range("P4").Value = "Greatest Total Volume"
    
    'Populate table one
    
        'Count the number of rows in the dataset
        Dim rowcount As Long
        rowcount = Cells(Rows.Count, 1).End(xlUp).Row
        
        'Keep track of the row in table one
        Dim tickercount As Long
        tickercount = 2
        
        'Store the sum of the volume for the current <ticker>
        'This variable needs to be a Double or the value will overflow since stock volume is a very large integer
        Dim volcount As Double
        volcount = 0
        
        'Store the value of the first <open>
        Dim firstopen As Double
        firstopen = Cells(2, 3).Value
        
        'Store the value of the biggest Percentage Change
        Dim greatestincrease As Double
        greatestincrease = 0
        
        'Store the ticker with the biggest Percentage Change
        Dim tickerincrease As String
        tickerincrease = 0
        
        'Store the value of the smallest Percentage Change
        Dim greatestdecrease As Double
        greatestdecrease = 0
        
        'Store the ticker with the smallest Percentage Change
        Dim tickerdecrease As String
        tickerdecrease = 0
        
        'Store the value of the biggest Total Stock Volume
        Dim greatestvol As Double
        greatestvol = 0
        
        'Store the ticker with the biggest Total Stock Volume
        Dim tickervol As String
        tickervol = Cells(2, 9).Value
    
            'Loop through the rows
            For a = 2 To rowcount
        
                'Add the <vol> of the current row to volume counter
                volcount = volcount + Cells(a, 7).Value
            
                'If the current <ticker> is not the same as the next <ticker>
                If Cells(a, 1).Value <> Cells(a + 1, 1).Value Then
    
                    'Print the current <ticker> in the Ticker column of table one
                    Cells(tickercount, 9).Value = Cells(a, 1).Value
                    
                    'Print the (last <close> - first <open>) in the Yearly Change column of table one with two decimal places
                    Cells(tickercount, 10).Value = FormatNumber(Cells(a, 6).Value - firstopen, 2)
                    'Format the value to have Excel's "Number" format
                    Cells(tickercount, 10).NumberFormat = "0.00"
                    
                        'If the Yearly Change is positive
                        If Cells(tickercount, 10).Value > 0 Then
                
                            'Fill the cell with a green color
                            Cells(tickercount, 10).Interior.Color = vbGreen
                    
                        'If the Yearly Change is negative
                        ElseIf Cells(tickercount, 10).Value < 0 Then
                        
                            'Fill the cell with a red color
                            Cells(tickercount, 10).Interior.Color = vbRed
                
                        End If
                    
                    'Print the ((last <close>/first <open>)-1) in the Percentage Change column
                    Cells(tickercount, 11).Value = FormatPercent((Cells(a, 6).Value / firstopen) - 1)
                    
                        'If the next Percentage Change is bigger than the current biggest Percentage Change
                        If Cells(tickercount, 11).Value > greatestincrease Then
                        
                            'Store the value of the bigger Percentage Change
                            greatestincrease = Cells(tickercount, 11).Value
                            
                            'Store the Ticker with the bigger Percentage Change
                            tickerincrease = Cells(tickercount, 9).Value
                        
                        'Else if the next Percentage Change is smaller than the current smallest Percentage Change
                        ElseIf Cells(tickercount, 11).Value < greatestdecrease Then
                        
                            'Store the value of the smaller Percentage Change
                            greatestdecrease = Cells(tickercount, 11).Value
                            
                            'Store the Ticker with the smaller Percentage Change
                            tickerdecrease = Cells(tickercount, 9).Value
            
                        End If
                    
                    'Print the volume count in the Total Stock Volume Column
                    Cells(tickercount, 12).Value = volcount
                    
                        'If the new Total Stock Volume is bigger than the current biggest Total Stock Volume
                        If Cells(tickercount, 12).Value > greatestvol Then
                        
                            'Store the value of the bigger Total Stock Volume
                            greatestvol = Cells(tickercount, 12).Value
                            
                            'Store the value of the new ticker
                            tickervol = Cells(tickercount, 9).Value
            
                        End If
                        
                    'Set the firstopen index to the first <open> of the new symbol
                    firstopen = Cells(a + 1, 3).Value
                    
                    'Set the volume counter to the oldest <open> of the next ticker
                    volcount = 0
                    
                    'Shift down by one row in table one
                    tickercount = tickercount + 1
            
                End If
                
            Next a
            
        'Print the greatest % increase
        Range("R2").Value = FormatPercent(greatestincrease)
            
        'Print the ticker of the greatest & increase
        Range("Q2").Value = tickerincrease
            
        'Print the greatest % decrease
        Range("R3").Value = FormatPercent(greatestdecrease)
            
        'Print the ticker of the greatest & decrease
        Range("Q3").Value = tickerdecrease
        
        'Print the greatest total volume
        Range("R4").Value = greatestvol
        
        'Print the ticker of the greatest Total Volume
        Range("Q4").Value = tickervol
    
    'Format table one
            
        'Auto adjust the width of the columns in the new table
        Columns("J:L").AutoFit
        Columns("P").AutoFit
    
    'Populate table two
    'IDEA: Loop through the column and if the new value is bigger/smaller, keep it; at the end, print the value in the appropriate cell
    
        'Store the greatest Percentage Change increase
        'Dim greatestincrease As Double
        'greatestincrease = Cells(2, 11).Value
        
        'Store the greatest increase ticker
        'Dim tickerincrease As String
        'tickerincrease = Cells(2, 9).Value
        
        'Store the greatest Percentage Change decrease
        'Dim greatestdecrease As Double
        'greatestdecrease = Cells(2, 11).Value
        
        'Store the greatest decrease ticker
        'Dim tickerdecrease As String
        'tickerdecrease = Cells(2, 9).Value
        
        'Store the greatest total volume
        'Dim greatestvol As Double
        'greatestvol = Cells(2, 12).Value
        
        'Store the greatest total volume ticker
        'Dim tickervol As String
        'tickervol = Cells(2, 9).Value
    
        'Loop through the rows
        'For c = 2 To tickercount
        
            'If the next Yearly Change is bigger than the current Yearly Change, then
            'If Cells(c + 1, 11).Value > greatestincrease Then
            
            'Store the value of the bigger Year Change
            'greatestincrease = Cells(c + 1, 11).Value
            
            'Store the value of the new ticker
            'tickerincrease = Cells(c + 1, 9).Value
            
            'Else if the next Yearly Change is smaller than the current Yearly Change, then
            'ElseIf Cells(c + 1, 11).Value < greatestdecrease Then
            
            'Store the value of the smaller Yearly Change
            'greatestdecrease = Cells(c + 1, 11).Value
            
            'Store the value of the new ticker
            'tickerdecrease = Cells(c + 1, 9).Value
            
            'End If
            
            'If the next Total Stock Volume is bigger than the current Total Stock Volume
            'If Cells(c + 1, 12).Value > greatestvol Then
            
            'Store the value of the bigger Total Stock Volume
            'greatestvol = Cells(c + 1, 12).Value
            
            'Store the value of the new ticker
            'tickervol = Cells(c + 1, 9).Value
            
           ' End If
            
        'Next c
        
        'Print the greatest % increase
        'Range("R2").Value = FormatPercent(greatestincrease)
        
        'Print the ticker of the greatest & increase
        'Range("Q2").Value = tickerincrease
        
        'Print the greatest % decrease
        'Range("R3").Value = FormatPercent(greatestdecrease)
        
        'Print the ticker of the greatest & decrease
        'Range("Q3").Value = tickerdecrease
        
        'Print the ticker of the greatest Total Volume
        'Range("Q4").Value = tickervol
    
    'Print the biggest Total Stock Volume in table two
    'Range("R4").Value = Application.WorksheetFunction.Max(Range("L2:L91"))

End Sub
