Attribute VB_Name = "Module1"
Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    
    
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    
    
    'Create a header row
    
    Cells(3, 1).Value = "Year"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
    
    
    
    Worksheets("2018").Activate
    
    
    
    'set initial volume to zero
    
    totalVolume = 0
    
    
    
    Dim startingPrice As Double
    
    Dim endingPrice As Double
    
    
    
        
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    
    'establish the number of rows to loop over
    
    rowStart = 2
    
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    

    
    'loop over all the rows
    
    For i = rowStart To rowEnd
    
        'increase totalVolume if ticker is "DQ"
        
        If Cells(i, 1).Value = "DQ" Then
        
            'increase totalVolume by the value in the current row
        
            totalVolume = totalVolume + Cells(i, 8).Value
            
        End If
            
        'set starting price
        
        If Cells(i, 1).Value = "DQ" And Cells(i - 1, 1).Value <> "DQ" Then
        
            startingPrice = Cells(i, 6).Value
        
        End If
        
        'set ending price
        
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then
        
            endingPrice = Cells(i, 6).Value
        
        End If
    
    Next i
    
    
    
    Worksheets("DQ Analysis").Activate
    
    Cells(4, 1).Value = 2018
    
    Cells(4, 2).Value = totalVolume
    
    Cells(4, 3).Value = endingPrice / startingPrice - 1
    
    
End Sub


Sub AllStocksAnalysis()

    
    
    'Format the output sheet on the "All Stocks Analysis" worksheet.
    
    Worksheets("All Stocks Analysis").Activate
    
    'create a title
    
    Dim yearValue As String
    
    yearValue = InputBox("What year would you like to run the analysis on?")
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'create a header row
    
    Cells(3, 1).Value = "Ticker"
    
    Cells(3, 2).Value = "Total Daily Volume"
    
    Cells(3, 3).Value = "Return"
    
    
    
    'Initialize an array of all tickers.
    
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    
    tickers(1) = "CSIQ"
    
    tickers(2) = "DQ"
    
    tickers(3) = "ENPH"
    
    tickers(4) = "FSLR"
    
    tickers(5) = "HASI"
    
    tickers(6) = "JKS"
    
    tickers(7) = "RUN"
    
    tickers(8) = "SEDG"
    
    tickers(9) = "SPWR"
    
    tickers(10) = "TERP"
    
    tickers(11) = "VSLR"

    
    
    'Prepare for the analysis of tickers.
    
    'Initialize variables for the starting price and ending price.
    
    Dim startingPrice As Double
    
    Dim endingPrice As Double

    'Activate the data worksheet.
    
    Worksheets("yearValue").Activate
    
    'Find the number of rows to loop over.
    
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row



    'Loop through the tickers.
    
    For i = 0 To 11
    
        ticker = tickers(i)
    
        totalVolume = 0
    
    
    
        'Loop through rows in the data.
    
        Worksheets("yearValue").Activate
        
        'Start an inner loop to go through all the rows of the data
        
        For j = 2 To RowCount
        
            'Find the total volume for the current ticker
        
            '(increase the value in total volume by the volume in the selected row, if the ticker matches the current ticker from the outer loop)
        
            If Cells(j, 1).Value = ticker Then
        
                totalVolume = totalVolume + Cells(j, 8).Value
        
                End If
        
            'Find the starting price for the current ticker.
            
            '(check if the current row is the first row with the selected ticker. If so, assign the current price to startingPrice)
            
            If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
               startingPrice = Cells(j, 6).Value
               
               End If
    
            'Find the ending price for the current ticker.
            
            If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
            
                endingPrice = Cells(j, 6).Value
                
                End If
                
        Next j

    
    
    'Output the data for the current ticker.
    
    Worksheets("All Stocks Analysis").Activate
    
    Cells(4 + i, 1).Value = ticker
    
    Cells(4 + i, 2).Value = totalVolume
    
    Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
    
    Next i

End Sub

Sub formatAllStocksAnalysisTable()

    'Formatting
    
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1,A3:C3").Font.Bold = True
    
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    
    Range("B4:B15").NumberFormat = "#,##0"
    
    Range("C4:C15").NumberFormat = "0.0%"
    
    Columns("B").AutoFit
    
    dataRowStart = 4
    
    dataRowEnd = 15
    
    
    
    For i = dataRowStart To dataRowEnd
    
        If Cells(i, 3) > 0 Then
        
            'Change cell color to green
            
            Cells(i, 3).Interior.Color = vbGreen
            
        ElseIf Cells(i, 3) < 0 Then
        
            'Change cell gcolor to red
            
            Cells(i, 3).Interior.Color = vbRed
            
        Else
        
            'Clear the cell color
            
            Cells(i, 3).Interior.Color = xlNone
            
        End If
        
    Next i
    
    
End Sub


    



