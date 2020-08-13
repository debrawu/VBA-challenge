Attribute VB_Name = "Module21"
Sub tickerStock()

'loop through each sheet
For Each ws In Worksheets

    'create variable to hold ticker name
    Dim ticker As String
    
    'create variable to hold volume
    Dim volume As Long
    'set variable to 0
    volume = 0
    
    'create variable to hold open year price
    Dim openPrice2 As Double
    'set to zero
    openPrice2 = 0
        
    'create variable to hold close year price
    Dim closePrice As Double
    'set to zero
    closePrice = 0
    
    'create variable for yearly change
    Dim yearChange As Double
    yearChange = 0
    
    'create variable for percent change
    Dim percentChange As Double
    percentChange = 0
    
    'create variable for greatest percent increase
    Dim greatestIncrease As Double
    
    'create variable for greatest percent decrease
    Dim greatestDecrease As String
    
    'create variable for greatest total volume
    Dim greatestTotalVol As Double
    
            
    'create summary table variable
    Dim summary_table As Integer
    'start it at row 2 so that the columns can be labeled
    summary_table = 2
    
    'find the lastrow in ticker stock
    Dim lastrow As Long
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'dim opening price
    openPrice = ws.Range("C2").Value
    
    'loop through stock data
    For i = 2 To lastrow
    
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
       
               
        
        'if ticker name is not the same, then...
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
            'set ticker name
            ticker = ws.Cells(i, 1).Value
                                        
            'place ticker name in summary table
            ws.Range("I" & summary_table).Value = ticker
            'ticker = ws.Cells(Rows.Count, "I").End(xlUp).Row - 1


            'grab close price ticker value
            closePrice = ws.Cells(i, 6).Value
            
            'yearly change calc
            yearChange = closePrice - openPrice
            
            
            'put values in column j
            ws.Range("J" & summary_table).Value = yearChange
            
            
            'calc percentchange
            percentChange = yearChange / openPrice
            
            'print in column k
            ws.Range("K" & summary_table).Value = percentChange
                  
                   
            'format to percent
            ws.Range("K2:K" & summary_table).NumberFormat = "0.00%"
            
           'add the volume
            volume = volume + ws.Cells(i, 7).Value
            
        
            'sum the volume in summary table
            ws.Range("L" & summary_table).Value = volume
            
                        
            'create new row for next ticker
            summary_table = summary_table + 1

            
            'make sure to reset the volume to 0 so it doesn't add the previous ticker volume
            volume = 0
            
                        
            'grab open price ticker value
            openPrice = ws.Cells(i + 1, 3).Value
            
          
            
        End If
        
        If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1).Value Then
        
        
            'add the volume
            volume = volume + ws.Cells(i, 7).Value
            
            
        End If
        
        
 
        
        'make new if statement for conditional formatting for red/green change in yearly change
        'create syntax
        If ws.Cells(i, 10).Value < 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 3
        ElseIf ws.Cells(i, 10).Value > 0 Then
        ws.Cells(i, 10).Interior.ColorIndex = 4
        
        End If

        'find out the greatest % increase
        greatestIncrease = Application.WorksheetFunction.Max(Range("K:K"))
        
        'Print the; percentage; into; the; cell
        Range("R2").Value = greatestIncrease
        
        'format to percent
        ws.Range("R2").NumberFormat = "0.00%"
        
        'find out the greatest % decrease
        greatestDecrease = Application.WorksheetFunction.Max(Range("K:K"))
        
        'Print the; percentage; into; the; cell
        ws.Range("R3").Value = greatestDecrease
        
        'format to percent
        ws.Range("R3").NumberFormat = "0.00%"
        
        'find out the greatest total volume
        greatestTotalVol = Application.WorksheetFunction.Max(Range("L:L"))
        
        'Print the; percentage; into; the; cell
        ws.Range("R4").Value = greatestTotalVol
        
        'format to percent
        ws.Range("R4").NumberFormat = "General"
        
        'create if then statements for the tickernames of each value
        If ws.Cells(i, 11).Value = ws.Range("R2").Value Then
        ws.Cells(i, 9).Value = ws.Range("Q2").Value
        ElseIf ws.Cells(i, 11).Value = ws.Range("R3").Value Then
        ws.Cells(i, 9).Value = ws.Range("Q3").Value
        ElseIf ws.Cells(i, 12).Value = ws.Range("r4").Value Then
        ws.Cells(i, 9).Value = ws.Range("Q4").Value
        End If
      
                  

  Next i
  
  
Next ws

End Sub
