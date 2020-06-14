Attribute VB_Name = "Module1"
Sub stock_data():
For Each ws In Worksheets
    Dim ticker As String
    Dim year_change As Double
    Dim percent_change As Double
    Dim stock_close As Double
    Dim stock_open As Double

    Dim total_volume As Double
    total_volume = 0

    Dim row_no As Long
    row_no = 2

    Dim last_row As Long
    last_row = ws.Cells(1, 1).End(xlDown).Row

    Dim ticker_count As Long
    ticker_count = 0

    Dim max_percent As Double
    Dim min_percent As Double
    Dim max_volume As Double


    For I = 2 To last_row
        If ws.Cells(I, 3).Value > 0 Then
            If ws.Cells(I, 1).Value <> ws.Cells(I + 1, 1).Value Then
                ticker = ws.Cells(I, 1).Value
                total_volume = total_volume + ws.Cells(I, 7).Value
                stock_close = ws.Cells(I, 6).Value
                stock_open = ws.Cells((I - ticker_count), 3).Value
        
                year_change = stock_close - stock_open
                percent_change = Round((stock_close - stock_open) / stock_open, 2)
                ws.Cells(row_no, 9).Value = ticker
                ws.Cells(row_no, 12).Value = total_volume
                ws.Cells(row_no, 10).Value = year_change
                ws.Cells(row_no, 11).Value = percent_change
        
                row_no = row_no + 1
                total_volume = 0
                ticker_count = 0
        
            Else
                total_volume = total_volume + ws.Cells(I, 7).Value
                ticker_count = ticker_count + 1
                
            End If
        
        End If
      
    Next I
    
    'Conditional formatting
    Dim summary_table_lastrow As Double
    summary_table_lastrow = ws.Cells(2, 9).End(xlDown).Row
    For I = 2 To summary_table_lastrow
        If ws.Cells(I, 10).Value < 0 Then
            ws.Cells(I, 10).Interior.ColorIndex = 3
    
        Else
            ws.Cells(I, 10).Interior.ColorIndex = 4
        
     
        End If

    'Max and Min values

    max_percent = Application.WorksheetFunction.Max(ws.Range("K:K"))
    min_percent = Application.WorksheetFunction.Min(ws.Range("K:K"))
    max_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))


    ws.Cells(2, 16).Value = max_percent
    max_ticker = Application.Match(ws.Cells(2, 16).Value, ws.Range("I:I"), 0)
    ws.Cells(2, 15).Value = max_ticker
    ws.Cells(3, 16).Value = min_percent
    min_ticker = Application.Match(ws.Cells(3, 16).Value, ws.Range("I:L"), 0)
    ws.Cells(3, 15).Value = min_ticker
    ws.Cells(4, 16).Value = max_volume
    max_volume_ticker = Application.Match(ws.Cells(4, 16).Value, ws.Range("I:L"), 0)
    ws.Cells(4, 15).Value = max_volume_ticker

    Next I
    
Next ws
    

End Sub
