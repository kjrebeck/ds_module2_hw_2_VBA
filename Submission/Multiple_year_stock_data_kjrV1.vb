
Sub stockticker():

    'define variables
    Dim ticker As String 'ticker, stock code - column I
    Dim quarterly_change As Double 'decimal column j
    Dim percent_change As Double 'decimal to % column k
    Dim volume As LongLong
    Dim total_stock_volume As LongLong 'large integer column L
    Dim first_open As Double 'stock opening value
    Dim last_close As Double 'stock closing value
    Dim lastRow As Long
    Dim k As LongLong 'output row
    Dim i As Long 'for iteration
    Dim highestper As Double 'greatest percent increas by ticker
    Dim lowestper As Double 'greatest percent decrease by ticker
    Dim highestvol As LongLong 'greatest volume by ticker

    'apply xode to all worksheets in file
    For Each ws In Worksheets

        'define last row, asked chatGPT -
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        'table output row
        k = 2
        
        'start value for each stocks total volume
        total_stock_volume = 0
                    
        'first open of worksheet
        first_open = ws.Cells(2, 3).Value
        
        'Set header names for first table
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Volume"
        'Set Headers for second table
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

            'Loop each row
            'For each stock get ticker code, quarterly change, % change, total volume
            'print values to table using loop
            
                
            For i = 2 To lastRow
                'set first ticker
                ticker = ws.Cells(i, 1).Value
                'set first volume
                volume = ws.Cells(i, 7).Value
                    
                    'is next row same ticker?
                    'If ticker in next row is different then :
                    If ws.Cells(i + 1, 1).Value <> ticker Then
                        'add to total stock volume
                        total_stock_volume = total_stock_volume + volume
                        
                        'get last close, first open defined prior to loop
                        last_close = ws.Cells(i, 6).Value
                        
                        'get quarterly change
                        quarterly_change = last_close - first_open
                        
                        'get percent change
                        'Find Percent - Greatest Increase & Decrease, Greatest Volum
                        'if/else statement for dividing by 0 error
                        If (first_open > 0) Then
                                percent_change = quarterly_change / first_open
                        Else
                                percent_change = 0
                        End If
                            
                        'first output table values
                        'print ticker to table
                        ws.Cells(k, 9).Value = ticker
                        'print quarterly change to table
                        ws.Cells(k, 10).Value = quarterly_change
                        'print percent change to table
                        ws.Cells(k, 11).Value = percent_change
                        'print total stock volume to table
                        ws.Cells(k, 12).Value = total_stock_volume
                            
                            'color quarterly change
                        If (quarterly_change > 0) Then
                            ws.Cells(k, 10).Interior.ColorIndex = 4
                            ws.Cells(k, 11).Interior.ColorIndex = 4
                        ElseIf (quarterly_change < 0) Then
                            ws.Cells(k, 10).Interior.ColorIndex = 3
                            ws.Cells(k, 11).Interior.ColorIndex = 3
                        Else
                            ws.Cells(k, 10).Interior.ColorIndex = 2
                            ws.Cells(k, 11).Interior.ColorIndex = 2
                        End If
                
                        If ticker = ws.Cells(2, 1).Value Then
                            highestvol = total_stock_volume
                            highestper = percent_change
                            lowestper = percent_change
                
                        Else
                            'is ticker total volume the greatest?
                            If total_stock_volume > highestvol Then
                                    highestvol = total_stock_volume
                            End If

                            If percent_change > highestper Then
                                    highestper = percent_change
                            End If
                            
                            If percent_change < lowestper Then
                                    lowestper = percent_change
                            End If
                        
                        End If
                                    
                        ws.Cells(2, 17).Value = highestper
                        ws.Cells(3, 17).Value = lowestper
                        ws.Cells(4, 17).Value = highestvol
                            
                'RESET Values
                
                'reset stock volume total
                total_stock_volume = 0
                'Add one to summary table row
                k = k + 1
                'reset first open
                first_open = ws.Cells(i + 1, 3).Value
                            
            Else
                'add to total stock volume
                total_stock_volume = total_stock_volume + volume
                                                    
            End If



        Next i
            ' Style my leaderboard
            ws.Columns("K:K").NumberFormat = "0.00%"
            ws.Columns("I:Q").AutoFit
            ws.Cells(2, 17).NumberFormat = "0.00%"
            ws.Cells(3, 17).NumberFormat = "0.00%"
            
    Next ws

    
            
End Sub