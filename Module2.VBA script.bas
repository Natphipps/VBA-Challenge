Attribute VB_Name = "Module1"
Sub stockmarket():


'loop through the worksheets
Dim ws As Worksheet

For Each ws In Worksheets


'create values for headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"



'define ticker variable
Dim ticker_name As String

'keep track of the ticker location in the table
Dim summary_table As Double
summary_table = 2

'define open price variable
Dim open_price As Double
Dim open_price_row As Double
open_price_row = 2


'define close price
Dim close_price As Double

'define yearly change
Dim yearly_change As Double

'define percent change
Dim percent_change As Double

'define total stock volume
Dim total_stock_volume As Double
total_stock_volume = 0

'count the number of rows
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'loop through the rows
For i = 2 To lastrow

    'check if next cell value is different from current cell value
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
    
    'set the ticker name
    ticker_name = ws.Cells(i, 1).Value
    
    'set total stock volume
    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    
    'print total stock volume in summary table
    ws.Range("L" & summary_table).Value = total_stock_volume
    
    'print ticker_name in summary table
    ws.Range("I" & summary_table).Value = ticker_name

    
    'open price ws.Range
    open_price = ws.Cells(open_price_row, 3).Value
    
    'close price ws.Range
    close_price = ws.Cells(i, 6).Value
    
    'calculate yearly change
    yearly_change = (close_price - open_price)
    
    
    'calculate percent change
    percent_change = ((close_price - open_price) / open_price)
    
    'print percent change
    ws.Range("K" & summary_table).Value = percent_change
    
    'change percent change into decimal format
    ws.Range("K" & summary_table).NumberFormat = "0.00%"
    
    'print yearly change
    ws.Range("J" & summary_table).Value = yearly_change
    
    
    'create an if statement for yearly change conditional formatting
     If ws.Range("J" & summary_table).Value >= 0 Then
    
        ws.Range("J" & summary_table).Interior.ColorIndex = 4
        
        ElseIf ws.Range("J" & summary_table).Value < 0 Then
        
        ws.Range("J" & summary_table).Interior.ColorIndex = 3

End If

   'create an if statement for percent change conditional formatting
    If ws.Range("K" & summary_table).Value >= 0 Then
    
     ws.Range("K" & summary_table).Interior.ColorIndex = 4
    
        ElseIf ws.Range("K" & summary_table).Value < 0 Then
        
            ws.Range("K" & summary_table).Interior.ColorIndex = 3
    
End If

    
    'reset total stock volume
    total_stock_volume = 0
    
    '
    summary_table = summary_table + 1
    open_price_row = i + 1
    
    
    
    Else
    
    total_stock_volume = total_stock_volume + ws.Cells(i, 7).Value
    
    
    
    
    
 End If
 
Next i


'create headers
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"
ws.Range("O2").Value = "greatest % increase"
ws.Range("O3").Value = "greatest % decrease"
ws.Range("O4").Value = "greatest total volume"

'grabs max/min from ranges
ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & lastrow))


maxpercentindex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
minpercentindex = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & lastrow)), ws.Range("K2:K" & lastrow), 0)
maxvolumeindex = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & lastrow)), ws.Range("L2:L" & lastrow), 0)

ws.Range("P2").Value = ws.Cells(maxpercentindex + 1, 9)
ws.Range("P3").Value = ws.Cells(minpercentindex + 1, 9)
ws.Range("P4").Value = ws.Cells(maxvolumeindex + 1, 9)




ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").NumberFormat = "0.00%"



Next ws


End Sub



