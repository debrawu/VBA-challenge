Attribute VB_Name = "Module21"
Sub tickerStock()

'loop through each sheet
For Each ws In Worksheets

    'create variable to hold ticker name
    Dim ticker As String
    ticker = ws.Cells(2, 1).Value
    'create variable to hold volume
    Dim volume As Variant
    'create variable to hold open year price
    Dim openPrice As Double
    'create variable to hold close year price
    Dim closePrice As Double
        
    'create variable for greatest percent increase
    Dim greatestIncrease As Double
    Dim greatestIncTic As String
    
    'create variable for greatest percent decrease
    Dim greatestDecrease As String
    Dim greatestDecTic As String
    
    'create variable for greatest total volume
    Dim greatestTotalVol As Double
    Dim greatestTotalTic As String
    
            
    'create summary table variable
    Dim summary_table As Integer
    'start it at row 2 so that the columns can be labeled
    summary_table = 2
    
    'find the lastrow in ticker stock
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'dim opening price
    openPrice = ws.Range("C2").Value
    
     
    'label summary table columns
    ws.Range("I1").Value = "Ticker Symbol"
    ws.Range("J1").Value = "YearlyChange"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    
    Dim yearChange As Variant
    Dim percentChange As Variant
    volume = 0
    'loop through stock data
    For i = 2 To lastrow
    
        'if ticker name is not the same, then...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set ticker name
            ticker = ws.Cells(i, 1).Value
                                        
            'place ticker name in summary table
            ws.Range("I" & summary_table).Value = ticker
            'ticker = ws.Cells(Rows.Count, "I").End(xlUp).Row - 1
            
            'add the volume
            volume = volume + ws.Cells(i, 7).Value
            
        
            'sum the volume in summary table
            ws.Range("L" & summary_table).Value = volume

            'grab close price ticker value
            closePrice = ws.Cells(i, 6).Value
            
                If Cells(summary_table, 3).Value = 0 Then
                    For find_value = summary_table To i
                    If Cells(find_value, 3).Value <> 0 Then
                    summary_table = find_value
                    Exit For
                    End If
                    Next find_value
                    End If
                    
                    
                
            openPrice = ws.Cells(summary_table, 3)
            'yearly change calc
            yearChange = closePrice - openPrice
            
            
            'put values in column j
            ws.Range("J" & summary_table).Value = yearChange
            
                 'make new if statement for conditional formatting for red/green change in yearly change
                'create syntax
                If yearChange < 0 Then
                ws.Cells(summary_table, 10).Interior.ColorIndex = 3
                ElseIf yearChange > 0 Then
                ws.Cells(summary_table, 10).Interior.ColorIndex = 4
                End If
            
            'calc percentchange
            percentChange = yearChange / openPrice
            
            'print in column k
            ws.Range("K" & summary_table).Value = percentChange
                  
                   
            'format to percent
            ws.Range("K2:K" & summary_table).NumberFormat = "0.00%"
            
                                  
            'create new row for next ticker
            summary_table = summary_table + 1

            
            'make sure to reset the volume to 0 so it doesn't add the previous ticker volume
            volume = 0
            
                        
            'grab open price ticker value
            openPrice = ws.Cells(i + 1, 3).Value
            
          
            
        End If
        
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        
            'add volume of tickers together
            volume = volume + ws.Cells(i, 7).Value
           'sum the volume in summary table
            ws.Range("L" & summary_table).Value = volume
            
            
            
        End If
        

        


  Next i
          'find out the greatest % increase
        greatestIncrease = Application.WorksheetFunction.Max(Range("K:K"))
        
        'Print the; percentage; into; the; cell
        Range("R2").Value = greatestIncrease
        
        'format to percent
        ws.Range("R2").NumberFormat = "0.00%"
        
        'find out the greatest % decrease
        greatestDecrease = Application.WorksheetFunction.Min(Range("K:K"))
        
        'Print the; percentage; into; the; cell
        ws.Range("R3").Value = greatestDecrease
        
        'format to percent
        ws.Range("R3").NumberFormat = "0.00%"
        
        'find out the greatest total volume
        greatestTotalVol = Application.WorksheetFunction.Max(Range("L:L"))
        
        'Print the; percentage; into; the; cell
        ws.Range("R4").Value = greatestTotalVol
        
        'format to regular number
        ws.Range("R4").NumberFormat = "General"
        
        'create if then statements for the tickernames of each value
        If ws.Cells(i, 11).Value = greatestIncrease Then
        ws.Cells(i, 9).Value = greatestIncTic
        ws.Range("Q2").Value = greatestIncTic
        
        End If
        
        If ws.Cells(i, 11).Value = greatestDecrease Then
        ws.Cells(i, 9).Value = greatestDecTic
        ws.Range("Q3").Value = greatestDecTic
        
        End If
        
        If ws.Cells(summary_table, 12).Value = greatestTotalVol Then
        ws.Cells(i, 9).Value = greatestTotalTic
        ws.Range("Q4").Value = greatestTotalTic
        
        End If
  
Next ws

End Sub
