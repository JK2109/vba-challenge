Option Explicit

Sub stock_price()
    

'Instructions
'Create a script that will loop through all the stocks for one year and output the following information:

'The ticker symbol.
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'The total stock volume of the stock.

    '--------------------------------
    'Creating a Summary Table Header
    '--------------------------------


    Dim ws As Worksheet
    
    For Each ws In Worksheets
        
        
    ws.Range("J:O").ClearContents 'To clear contents for rerun
    ws.Range("J:O").ClearFormats   'To clear format for rerun
    
        Dim Summary_Table_Header(5) As String
            Summary_Table_Header(0) = "Ticker"
            Summary_Table_Header(1) = "Yearly Change"
            Summary_Table_Header(2) = "Percent Change"
            Summary_Table_Header(3) = "Total Stock Volume"
            Summary_Table_Header(4) = "Open Price of the Year"
            Summary_Table_Header(5) = "Close Price of the Year"
            
        Dim Summary_Table_Col As Integer
            Summary_Table_Col = 10         'Defining Summary Table Column Location
        
        Dim k As Integer
            
            For k = 0 To 5
                ws.Cells(1, Summary_Table_Col + k).Value = Summary_Table_Header(k)
                ws.Cells(1, Summary_Table_Col + k).Font.Bold = True
                ws.Columns("J:O").AutoFit
                
            Next k
           
        '------------------------------------------------------
        'Extracting Unique Ticker Symbols and Opening & Closing Price to Summary Table
        'Calculating Total Stock Volume
        '------------------------------------------------------
        
        Dim ticker_symbol As String
        
        Dim Total_Stock_Volume As Double
            Total_Stock_Volume = 0
            
        Dim Summary_Table_Row As Integer
            Summary_Table_Row = 2
        
        
            
        Dim Open_Price As Double
        Dim Close_Price As Double
        
        
        
        Dim i As Long
        Dim LastRow As Long     'Finding the LastRow
            LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            
            
            
            For i = 2 To LastRow
            
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                    ticker_symbol = ws.Cells(i, 1).Value
                    
                    Close_Price = ws.Cells(i, 6).Value
                                                  
                                   
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value 'Add Stock Volume to the total
                                                 
                    ws.Cells(Summary_Table_Row, Summary_Table_Col).Value = ticker_symbol
                    ws.Cells(Summary_Table_Row, Summary_Table_Col + 5).Value = Close_Price
                    ws.Cells(Summary_Table_Row, Summary_Table_Col + 3).Value = Total_Stock_Volume
                                   
                    Summary_Table_Row = Summary_Table_Row + 1
                    
                    
                    Total_Stock_Volume = 0  'Reset the total stock volume
                     
                Else
                    
                    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value 'Add Stock Volume to the total if the row is same ticker
                   
                End If
                
                      
                If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then
                    Open_Price = ws.Cells(i, 3).Value
                    ws.Cells(Summary_Table_Row, Summary_Table_Col + 4).Value = Open_Price
                End If
                              
              
            Next i
           
        
                         
        '---------------------------------------------------------------------------------------------------------------------------------
        'Calculating Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
        'Calculating The percent change from opening price at the beginning of a given year to the closing price at the end of that year
        '---------------------------------------------------------------------------------------------------------------------------------
        Dim j As Long
        Dim LastRow_Table As Long
        Dim Price_Change As Double
        Dim Percent_Change As Double
        
            LastRow_Table = ws.Cells(Rows.Count, Summary_Table_Col).End(xlUp).Row
        
            For j = 2 To LastRow_Table
                
                    Price_Change = ws.Cells(j, Summary_Table_Col + 5).Value - ws.Cells(j, Summary_Table_Col + 4).Value
                    ws.Cells(j, Summary_Table_Col + 1).Value = Price_Change
                
                    If ws.Cells(j, Summary_Table_Col + 4).Value = 0 Or ws.Cells(j, Summary_Table_Col + 5).Value = 0 Then
                        Percent_Change = 0
                        ws.Cells(j, Summary_Table_Col + 2).Value = Percent_Change
                    Else
                        Percent_Change = ws.Cells(j, Summary_Table_Col + 5).Value / ws.Cells(j, Summary_Table_Col + 4).Value - 1
                    ws.Cells(j, Summary_Table_Col + 2).Value = Percent_Change
                    
                   
                    End If
                
            Next j
        
       
    
        '---------------------------------------------------------------------------------------------------------------
        'Solution to return the stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".
        '---------------------------------------------------------------------------------------------------------------
        
        Dim m As Long
        Dim max_inc As Double
        Dim max_dec As Double
        Dim max_vol As Double
        Dim max_inc_ticker, max_dec_ticker, max_vol_ticker As String
        
        
            max_inc = ws.Cells(2, Summary_Table_Col + 2) 'Initializing the Greatest % Increase(max increase) to the first stock
            max_dec = ws.Cells(2, Summary_Table_Col + 2) 'Initializing the Greatest % decrease(max decrease) to the first stock
            max_vol = ws.Cells(2, Summary_Table_Col + 3) 'Initializing the Greatest total volume to the first stock
                
            
            
            For m = 2 To LastRow_Table
                   
                   
                    If ws.Cells(m, Summary_Table_Col + 2).Value >= max_inc Then
                        max_inc = ws.Cells(m, Summary_Table_Col + 2).Value 'resetting the value if stock with higher % increase found
                        max_inc_ticker = ws.Cells(m, Summary_Table_Col).Value 'tagging the ticker with the higher % increase
                        
                    End If
                    
                    If ws.Cells(m, Summary_Table_Col + 2).Value <= max_dec Then
                        max_dec = ws.Cells(m, Summary_Table_Col + 2).Value      'resetting the value if stock with higher % decrease found
                        max_dec_ticker = ws.Cells(m, Summary_Table_Col).Value   'tagging the ticker with the higher % decrease
                        
                    End If
                                   
                    If ws.Cells(m, Summary_Table_Col + 3).Value >= max_vol Then
                        max_vol = ws.Cells(m, Summary_Table_Col + 3).Value      'resetting the value if stock with higher volume found
                        max_vol_ticker = ws.Cells(m, Summary_Table_Col).Value   'tagging the ticker with the higher volume
                    End If
                    
                 
                'Inserting the summary table
                    
                    ws.Range("S1") = "Ticker"
                    ws.Range("T1") = "Value"
                    
                    ws.Range("R2") = "Greatest % Increase"
                    ws.Range("R3") = "Greatest % Decrease"
                    ws.Range("R4") = "Greatest Total Volume"
                    
                    ws.Range("S2") = max_inc_ticker
                    ws.Range("S3") = max_dec_ticker
                    ws.Range("S4") = max_vol_ticker
                    
                    ws.Range("T2") = max_inc
                    ws.Range("T3") = max_dec
                    ws.Range("T4") = max_vol
                
            
            Next m
        
        
         '-------------------------------------------------
        'Summary Table Formatting
        '-------------------------------------------------
        
        
        ws.Range("K:K").NumberFormat = "#,##0.00"
        ws.Range("L:L").NumberFormat = "#,##0.00%"
        ws.Range("N:O").NumberFormat = "#,##0.00"
        ws.Range("M:M").NumberFormat = "#,##0"
        
        ws.Range("T2:T3").NumberFormat = "#,##0.00%"
        ws.Range("T4").NumberFormat = "#,##0"
        ws.Columns("R:T").AutoFit
        
        Dim l As Long
        
            For l = 2 To LastRow_Table
            
                If ws.Cells(l, Summary_Table_Col + 1).Value > 0 Then
                    ws.Cells(l, Summary_Table_Col + 1).Interior.ColorIndex = 4
                
                ElseIf ws.Cells(l, Summary_Table_Col + 1).Value = 0 Then
                    ws.Cells(l, Summary_Table_Col + 1).Interior.ColorIndex = 45
                        
                Else
                    ws.Cells(l, Summary_Table_Col + 1).Interior.ColorIndex = 3
                End If
            Next l
    
        
        
        
    
    Next ws
        MsgBox ("Job Completed")
    

End Sub



