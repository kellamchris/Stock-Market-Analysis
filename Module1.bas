Attribute VB_Name = "Module1"
Sub Module2_challenge()

    'here I'm declaring variables
    Dim ws As Worksheet
    Dim i As Long
    Dim ticker As String
    Dim summary_table_row As Long
    Dim lastRow As Long
    Dim openPrice As Double
    Dim closePrice As Double
    Dim percentChange As Double
    Dim totalVolume As Variant
    Dim firstRow As Long
        
    'loop through each worksheet
    For Each ws In Worksheets
    
        'set total volume for worksheet
        totalVolume = 0
        
        'add headers to the worksheet
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'calculate the last row in Column A of current worksheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'set starting rows for loop
        summary_table_row = 2
        firstRow = 2
        
        'loop through each row in the worksheet to output info
        For i = 2 To lastRow
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                'set the ticker
                ticker = ws.Cells(i, 1).Value
                ws.Cells(summary_table_row, 9).Value = ticker
                                               
                'set the quarterly change
                openPrice = ws.Cells(firstRow, 3).Value
                closePrice = ws.Cells(i, 6).Value
                ws.Cells(summary_table_row, 10).Value = closePrice - openPrice
                
                'set the percentage change
                percentChange = ((closePrice - openPrice) / openPrice)
                ws.Cells(summary_table_row, 11).Value = percentChange
                
                'add to the total volume
                totalVolume = totalVolume + ws.Cells(i, 7).Value
                
                'print total volume in the summary table
                ws.Cells(summary_table_row, 12).Value = totalVolume
               
                'add one to the summary table row
                summary_table_row = summary_table_row + 1
            
                'reset total volume for the next ticker
                totalVolume = 0
                
                'update firstRow for the next ticker
                firstRow = i + 1
                
            Else
                totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            End If
            
        Next i
        
        'add conditional formatting for quarterly change
        Dim lastRowJ As Long
        Dim Zero As Integer
        Dim cell As Range
         
        Zero = 0
         
        'calculate the last row in Column J of current worksheet
        lastRowJ = ws.Cells(ws.Rows.Count, "J").End(xlUp).Row
         
        For Each cell In ws.Range("J2:J" & lastRowJ)
            If cell.Value > 0 Then
                cell.Interior.ColorIndex = 4
                 
            Else
                cell.Interior.ColorIndex = 3
                 
            End If
                 
        Next cell
        
        'add conditional formatting for percent change
        Dim lastRowK As Long
        
        'calculate the last row in Column K of current worksheet
        lastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
        
        For Each cell In ws.Range("K2:K" & lastRowK)
            If cell.Value > 0 Then
                cell.Interior.ColorIndex = 4
                
            Else
                cell.Interior.ColorIndex = 3
                
            End If
            
        Next cell
                 
        'add functionality table
        Dim maxValue As Double
        Dim minValue As Double
        Dim lastRowL As Long
        Dim maxVolume As Variant
        Dim maxTicker As String
        Dim minTicker As String
        Dim maxVolumeTicker As String
        
        'add headers
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
               
        'calculate the last row in Column L of current worksheet
        lastRowL = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
                 
        'Return stock with Greatest % increase
        maxValue = ws.Cells(2, 11).Value
        maxTicker = ws.Cells(2, 9).Value
         
        For Each cell In ws.Range("K2:K" & lastRowK)
            If cell.Value > maxValue Then
                maxValue = cell.Value
                maxTicker = ws.Cells(cell.Row, 9).Value
            End If
             
        Next cell
            
        ws.Cells(2, 16).Value = maxValue
        ws.Cells(2, 15).Value = maxTicker
         
        'Return stock with greatest % decrease
        minValue = ws.Cells(2, 11).Value
        minTicker = ws.Cells(2, 9).Value
         
        For Each cell In ws.Range("K2:K" & lastRowK)
            If cell.Value < minValue Then
                minValue = cell.Value
                minTicker = ws.Cells(cell.Row, 9).Value
            End If
             
        Next cell
         
        ws.Cells(3, 16).Value = minValue
        ws.Cells(3, 15).Value = minTicker
                 
        'Return stock with greatest total volume
        maxVolume = ws.Cells(2, 12).Value
        maxVolumeTicker = ws.Cells(2, 9).Value
         
        For Each cell In ws.Range("L2:L" & lastRowL)
            If cell.Value > maxVolume Then
                maxVolume = cell.Value
                maxVolumeTicker = ws.Cells(cell.Row, 9).Value
                 
            End If
             
        Next cell
         
        ws.Cells(4, 16).Value = maxVolume
        ws.Cells(4, 15).Value = maxVolumeTicker
        
    Next ws
    
    MsgBox ("Script has finished running")
    
End Sub
