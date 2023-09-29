Attribute VB_Name = "Module1"
Sub stockCalc():

' was able to run this code okay in the test sheet, but am running into overflow error when
' finding last row of data

For Each ws In ThisWorkbook.Worksheets
   
      ' populate summary table headers
    
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        
        ' set variables
        Dim ticker As String
        Dim yearlyChange, percentChange, openPrice, high, low, closePrice, difference, increaseIndex, greatestIncrease As Double
        Dim volume, minDate, maxDate As Long
        
     
        volume = 0
    
        Dim summaryTableRows As Integer
        summaryTableRows = 2
        
        Dim tickerStart As Integer
        tickerStart = 2
        
        ' keep getting overflow error here
        Dim lastrow As Integer
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        ' loop through all the rows of data
        For Row = 2 To lastrow
            
            'check if ticker is the same or has changed to the next row
            If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                ' set ticker value
                ticker = Cells(Row, 1).Value
                ' calculate total volume
                volume = volume + Cells(Row, 7).Value
                ' set yearly closing price
                closePrice = Cells(Row, 6).Value
                ' set yearly opening price
                openPrice = Cells(tickerStart, 3).Value
                ' find price differences
                difference = openPrice - closePrice
                ' find percent change
                percentChange = difference / closePrice
                
                ' populate summary table rows with values
                ws.Cells(summaryTableRows, 9).Value = ticker
                
                ws.Cells(summaryTableRows, 10).Value = difference
                
                ws.Cells(summaryTableRows, 11).Value = percentChange
                
                ws.Cells(summaryTableRows, 12).Value = volume
                
                    For findValue = tickerStart To Row
                        If Cells(findValue, 3).Value <> 0 Then
                            tickerStart = findValue
                        End If
                    Next findValue
                
                'conditional formatting to color boxes depending on positive or
                'negative values
                
                If ws.Cells(summaryTableRows, 10).Value > 0 Then
                    ws.Cells(summaryTableRows, 10).Interior.ColorIndex = 4
                ElseIf ws.Cells(summaryTableRows, 10) < 0 Then
                    ws.Cells(summaryTableRows, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(summaryTableRows, 10) = 0 Then
                    ws.Cells(summaryTableRows, 10).Interior.ColorIndex = 0
                End If
                
                ws.Cells(summaryTableRows, 11).Style = "Percent"
                
                summaryTableRows = summaryTableRows + 1
                
                ' reset volume to 0
                volume = 0
                
                'Autofit summary table
                Columns("I:L").AutoFit
                
                
        
            Else
             
                'ticker unchanged, tally volume
                volume = volume + Cells(Row, 7).Value
            
            
            
            End If
            
            
            
        Next Row
        
        ' here i tried to find the max percent change
        greatestIncrease = WorksheetFunction.Max(ws.Cells(summaryTableRows, 10))
        increaseIndex = WorksheetFunction.Match(greatestIncrease, ws.Cells(summaryTableRows, 10))
        Range("P2").Value = ws.Range("I" & increaseIndex + 1).Value
        Range("Q2").Value = greatestIncrease
        
        
    Next ws
    
    
    
End Sub
