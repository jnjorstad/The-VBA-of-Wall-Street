Attribute VB_Name = "Module1"
Sub Stock_Data()
    
    For Each ws In Worksheets
    
        Dim WorksheetName As String
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        WorksheetName = ws.Name
    
        'assign each column accordingly
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'set initial variable for holding ticker
        Dim Ticker As String
        
        'set initial variable for holding the total per ticker
        Dim Total_Stock_Volume As Double
        Total_Stock_Volume = 0
        
        'set initial variables for holding year open, close, yearly change, and percent change
        Dim open_price As Double
        open_price = 0
        Dim close_price As Double
        close_price = 0
        Dim yearly_change As Double
        yearly_change = 0
        Dim percent_change As Double
        percent_change = 0
        
        
        'keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        'set year open
        open_price = Cells(2, 3).Value
        
        'loop through all stocks
        For i = 2 To LastRow
            
            'check if we are still within same ticker, if it's not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'set ticker symbol
                Ticker = ws.Cells(i, 1).Value
    
                'set year close
                close_price = ws.Cells(i, 6).Value
                
                'calculate yearly change
                yearly_change = close_price - open_price
                
                'calculate percent change
                percent_change = Round((yearly_change / open_price) * 100, 2)
                
                'add to ticker total
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
                
                'print ticker symbol in summary table
                ws.Range("I" & Summary_Table_Row).Value = Ticker
                      
                      'print yearly change in summary table
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
                      
                      'print percent change in summary table
                ws.Range("K" & Summary_Table_Row).Value = percent_change
                      
                      'print ticker total in summary table
                ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
                
                'add one to summary table row
                Summary_Table_Row = Summary_Table_Row + 1
                
                'reset ticker total, yearly change, year close, and change year open
                Total_Stock_Volume = 0
                
                yearly_change = 0
                
                close_price = 0
                
                percent_change = 0
                
                open_price = ws.Cells(i + 1, 3).Value
            
            Else
                'add to the ticker total
                Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
                
            End If
            
            'assign colors to yearly change column
            If ws.Cells(i, 10).Value > 0 Then
            
                 ws.Cells(i, 10).Interior.ColorIndex = 4
            
            Else
                
                 ws.Cells(i, 10).Interior.ColorIndex = 3
            
            End If
            
        Next i
    
    Next ws
    
End Sub

