Attribute VB_Name = "Module1"
Sub summerize_stock()
Dim ticker As String
Dim last_row, summary_line As Long
Dim open_price, close_price, total_volume, most_increase, most_decrease, most_volume As Double

'loop through all worksheets
For Each ws In Worksheets

    'Get row count
    last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'initialize variables
    ticker = ws.Cells(2, 1).Value
    open_price = ws.Cells(2, 3).Value
    close_price = 0
    total_volume = 0
    summary_line = 2
    most_increase = 0
    most_decrease = 0
    most_volume = 0
    
    'initialize summary column titles
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    'loop through all rows
    For i = 2 To last_row
        
        If ws.Cells(i + 1, 1).Value <> ticker Then
            'define close price
            close_price = ws.Cells(i, 6).Value
            
            'collect summary data
            ws.Cells(summary_line, 9).Value = ticker
            ws.Cells(summary_line, 10).Value = close_price - open_price
                If close_price - open_price > 0 Then
                    ws.Cells(summary_line, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(summary_line, 10).Interior.ColorIndex = 3
                End If
                
            'caliculate % when open price is not 0
            If open_price = 0 Then
                ws.Range("K" & summary_line).Value = 0
            Else
                ws.Range("K" & summary_line).Value = (close_price - open_price) / open_price
            End If
        
            'print summary lines
            ws.Range("K" & summary_line).NumberFormat = "0.00%"
            ws.Range("L" & summary_line).Value = total_volume + ws.Cells(i, 7).Value
            
            
            'Greatest increase
            If ws.Range("K" & summary_line).Value > most_increase Then
                most_increase_ticker = ticker
                most_increase = ws.Range("K" & summary_line).Value
            End If
    
            'Greatest decrease
            If ws.Range("K" & summary_line).Value < most_decrease Then
                most_decrease_ticker = ticker
                most_decrease = ws.Range("K" & summary_line).Value
            End If
            
            'Greatest volume
            If most_volume < ws.Range("L" & summary_line).Value Then
                most_volume_ticker = ticker
                most_volume = ws.Range("L" & summary_line).Value
            End If
                
            'initialize variables after print
            ticker = ws.Cells(i + 1, 1).Value
            open_price = ws.Cells(i + 1, 3).Value
            summary_line = summary_line + 1
            total_volume = 0
            
        Else
            'accumulate total volume
            total_volume = total_volume + ws.Cells(i, 7).Value
        
        End If
    
    Next i
    
    'Challenges - Adding Greatest Values
    'Add titles
    ws.Range("O1").Value = "Ticker"
    ws.Range("P1").Value = "Value"
    
    'Print greatest values
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("O2").Value = most_increase_ticker
    ws.Range("P2").Value = most_increase
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("O3").Value = most_decrease_ticker
    ws.Range("P3").Value = most_decrease
    ws.Range("N4").Value = "Greatest Total Volume"
    ws.Range("O4").Value = most_volume_ticker
    ws.Range("P4").Value = most_volume
    
    'Format and reset cell width
    ws.Range("P2:P3").NumberFormat = "0.00%"
    ws.Columns("A:P").EntireColumn.AutoFit

Next ws

MsgBox ("All Done!")

End Sub
Sub Initialize_sheets()

'cleanup
For Each ws In Worksheets
    ws.Columns("I:P").Delete shift:=xlToLeft

Next ws

End Sub
