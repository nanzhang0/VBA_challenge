Attribute VB_Name = "Module2"
Sub stock_trade_sumamry()

 'set ws as a worksheet object
    Dim ws As Worksheet
    Dim summary_table_header As Boolean
    summary_table_header = False
    
    Dim spreadsheet As Boolean
    spreadsheet = True
    
    'loop through all of the worksheets in the active workbook
    For Each ws In Worksheets
    
        'declare variable to hold the ticker name
        Dim ticker_name As String
        ticker_name = " "
        
        'delcare variable to hold the total per ticker name
        Dim total_ticker_volume As Double
        total_ticker_volume = 0
        
        'declare variables
        Dim open_price As Double
        open_price = 0
        
        Dim close_price As Double
        close_price = 0
        
        Dim change_price As Double
        change_price = 0
        
        Dim change_percent As Double
        change_percent = 0
        
        Dim max_ticker_name As String
        max_ticker_name = " "
        
        Dim min_ticker_name As String
        min_ticker_name = " "
        
        Dim max_percent As Double
        max_percent = 0
        
        Dim min_percent As Double
        min_percent = 0
        
        Dim max_volume_ticker As String
        max_volume_ticker = " "
        
        Dim max_volume As Double
        max_volume = 0
        '----------------------------------------------------------------
         
       
        'set summary table for worksheets
        Dim summary_table_row As Long
        summary_table_row = 2
        
        'set row count for worksheets
        Dim LastRow As Long
        Dim i As Long
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        
        If summary_table_header Then
            'set titles for the summary table
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Yearly Change"
            ws.Range("K1").Value = "Percent Change"
            ws.Range("L1").Value = "Total Stock Volume"
            
            'set titles for new summary table
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
        Else
           
            summary_table_header = True
        End If
        
        'set open price for the first ticker of ws
        open_price = ws.Cells(2, 3).Value
        
        'loop the worksheet to last row
        For i = 2 To LastRow
      
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'set the ticker name
                ticker_name = ws.Cells(i, 1).Value
                
                'calculate change_price and change_percent
                close_price = ws.Cells(i, 6).Value
                change_price = close_price - open_price
                
                
                If open_price <> 0 Then
                    
                    change_percent = (change_price / open_price) * 100
              
                End If
                
                
                total_ticker_volume = total_ticker_volume + ws.Cells(i, 7).Value
              
                
                
                ws.Range("I" & summary_table_row).Value = ticker_name
               
                ws.Range("J" & summary_table_row).Value = change_price
                
                'formatting
                If (change_price > 0) Then
                    
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 4
                ElseIf (change_price <= 0) Then
                    
                    ws.Range("J" & summary_table_row).Interior.ColorIndex = 3
                End If
                
                
                ws.Range("K" & summary_table_row).Value = (CStr(change_percent) & "%")
                
                ws.Range("L" & summary_table_row).Value = total_ticker_volume
                
                'add 1 to the summary table row count
                summary_table_row = summary_table_row + 1
                
                'reset change_price and change_percent holders
                change_price = 0
               
                close_price = 0
              
                open_price = ws.Cells(i + 1, 3).Value
              
                
                'populate new summary table
              
                If (change_percent > max_percent) Then
                    max_percent = change_percent
                    max_ticker_name = ticker_name
                ElseIf (change_percent < min_percent) Then
                    min_percent = change_percent
                    min_ticker_name = ticker_name
                End If
                       
                If (total_ticker_volume > max_volume) Then
                    max_volume = total_ticker_volume
                    max_volume_ticker = ticker_name
                End If
                
                'reset counters
                change_percent = 0
                total_ticker_volume = 0
                
          
            Else
                
                total_ticker_volume = total_ticker_volume + ws.Cells(i, 7).Value
                
            End If
           
      
        Next i

           
            'record all new counts to the new summary table
            If Not spreadsheet Then
            
                ws.Range("Q2").Value = (CStr(max_percent) & "%")
                ws.Range("Q3").Value = (CStr(min_percent) & "%")
                ws.Range("P2").Value = max_ticker_name
                ws.Range("P3").Value = min_ticker_name
                ws.Range("Q4").Value = max_volume
                ws.Range("P4").Value = max_volume_ticker
                
            Else
                spreadsheet = False
            End If
        
     Next ws
End Sub
