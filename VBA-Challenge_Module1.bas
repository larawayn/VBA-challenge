Attribute VB_Name = "Module1"
Sub VBA_Challenge()


'Declare variables
   Dim ws As Worksheet
   Dim ticker_symbol As String
   Dim last_row As Long
   Dim ticker_index As Integer
   Dim open_price As Double
   Dim close_price As Double
   Dim yearly_change As Double
   Dim percent_change As Double
   Dim stock_volume As Variant
   Dim row_index As Long
   
       
    'Loop through all worksheets
    For Each ws In Worksheets
       
    'Grabbed the worksheet name
       WorksheetName = ws.Name
    'Create row index to store data in created table
        row_index = 2
    'find last row of stock data
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    'set variables
       open_price = ws.Cells(2, 3).Value
       close_price = 0
       yearly_change = 0
       percent_change = 0
       stock_volume = 0

       
        For i = 2 To last_row
        'compile ticker symbols and remove duplicates
            ticker_symbol = ws.Cells(i, 1).Value
            If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
            ws.Cells(row_index, 9) = ticker_symbol
        'find closed price
            close_price = ws.Cells(i, 6).Value
        'calculate yearly change
            yearly_change = close_price - open_price
        'Find any open_price that's zero to avoid dividing by zero when calculating percent_change
        'Then calculate percent_change
                If open_price <> 0 Then
                percent_change = (yearly_change / open_price)
                Else
                percent_change = 0
                End If
        '{lace yearly_change in column 10
            ws.Cells(row_index, 10) = yearly_change
                If ws.Cells(row_index, 10) >= 0 Then
                ws.Cells(row_index, 10).Interior.ColorIndex = 4
                Else
                ws.Cells(row_index, 10).Interior.ColorIndex = 3
                End If
            
        'Place percent_change in column 11
            ws.Cells(row_index, 11) = percent_change
        'Format percent_change
            ws.Cells(row_index, 11).NumberFormat = "0.00%"
        'Reset open_price for next ticker
            open_price = ws.Cells(i + 1, 3).Value
        'Calculate stock_volume, place it in column 12 and reset it to zero for next ticker
            stock_volume = stock_volume + ws.Cells(i, 7).Value
            ws.Cells(row_index, 12) = stock_volume
            stock_volume = 0
        'update row_index volume
            row_index = row_index + 1
        'Continue calculating stock_volume
            Else
                stock_volume = stock_volume + ws.Cells(i, 7).Value
            
            End If
    
        Next i
        
            'Add titles to cells
            ws.Range("I1") = "Ticker"
            ws.Range("J1") = "Yearly Change"
            ws.Range("K1") = "Percent Change"
            ws.Range("L1") = "Total Stock Volume"
            ws.Range("P1") = "Ticker"
            ws.Range("Q1") = "Value"
            ws.Range("O1") = "Bonus Calculations"
            ws.Range("O2") = "Greatest % Increase"
            ws.Range("O3") = "Greatest % Decrease"
            ws.Range("O4") = "Greatest Total Volume"
            
            
            'Find Max, Min, largest total volume
            ws.Range("Q2") = Application.WorksheetFunction.Max(ws.Range("K:K"))
            ws.Range("Q3") = Application.WorksheetFunction.Min(ws.Range("K:K"))
            ws.Range("Q2:Q3").NumberFormat = "0.00%"
            ws.Range("Q4") = Application.WorksheetFunction.Max(ws.Range("L:L"))
            ws.Range("P2") = Application.WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K:K"), 0))
            ws.Range("P3") = Application.WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K:K"), 0))
            ws.Range("P4") = Application.WorksheetFunction.Index(ws.Range("I:I"), WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L:L"), 0))
            ws.Range("O1:O4").ColumnWidth = 20
          

    Next ws

End Sub
