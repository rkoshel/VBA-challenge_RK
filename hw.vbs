Sub Stock_Data()


'Set variables

Dim ticker As String

Dim number_tickers As Integer

Dim LastRow As Long

Dim opening_price As Double

Dim closing_price As Double

Dim yearly_change As Double

Dim percent_change As Double

Dim total_stock_value As Double


'Loop for all worksheets in workbook

For Each ws In Worksheets

    ws.Activate
    
    'Find last row in worksheet
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Set columns
    
    ws.Range("J1").Value = "Ticker"
    
    ws.Range("K1").Value = "Yearly Change"
    
    ws.Range("L1").Value = "Percent Change"
    
    ws.Range("M1").Value = "Total Stock Value"
    
    
    'Initialize variables
    
    number_tickers = 0
    
    ticker = ""
    
    yearly_change = 0
    
    opening_price = 0
    
    percent_change = 0
    
    total_stock_volume = 0
    
    
    For i = 2 To LastRow
    
        ticker = Cells(i, 1).Value
        
        If opening_price = 0 Then
        
        opening_price = Cells(i, 3).Value
        
        End If
        
        
        'Total stock volume
        
        total_stock_volume = total_stock_volume + Cells(i, 7).Value
        
        
        'Make ticker list
        
        If Cells(i + 1, 1).Value <> ticker Then
        
        number_tickers = number_tickers + 1
        
        Cells(number_tickers + 1, 10) = ticker
        
        closing_price = Cells(i, 6)
        
        yearly_change = closing_price - opening_price
        
        
        'Make yearly change list
        
        Cells(number_tickers + 1, 11).Value = yearly_change
        
        
        'Set colors
        
        If yearly_change >= 0 Then
        
        Cells(number_tickers + 1, 11).Interior.ColorIndex = 4
        
        Else
        
        Cells(number_tickers + 1, 11).Interior.ColorIndex = 3
        
        End If
        
        
        'Percent change value
        
        If opening_price = 0 Then
        
        percent_change = 0
        
        Else
        
        percent_change = (yearly_change / opening_price)
        
        
        'Set percent change
        
        Cells(number_tickers + 1, 12).Value = Format(percent_change, "Percent")
        
        opening_price = 0
        
        Cells(number_tickers + 1, 13).Value = total_stock_volume
        
        total_stock_volume = 0
        
        
        End If
        
        End If
        
        Next i
        
        Next ws

End Sub
