Attribute VB_Name = "Module1"
Sub StockMarketAnalysis()
Dim ws As Worksheet
'loop through each worksheet
For Each ws In ActiveWorkbook.Worksheets
ws.Activate

Dim ticker_symbol As String

Dim first_worksheet As String
Dim last_row As Long
Dim i As Long



Dim opening_price As Double
Dim closing_price As Double
Dim percentage_change As Double

opening_price = Cells(2, 3).Value
'figure out the last row in each worksheet
last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
Dim ticker_row As Long
'separate variable to increment for our output since i would get us out of cell alignment
ticker_row = 2
'start at zero or else get an overflow
Dim trading_volume As Double
trading_volume = 0



'Put in the additional Column Headings
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

     'loop through all the rows in each sheet to get the unique ticker symbols
    For i = 2 To last_row
    'Debug.Print last_row
    'if the row I am on and row below have a different symbol, then
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'we write the first entry to our ticker column at the last end of our ticker year
            ticker_symbol = Cells(i, 1).Value
            Cells(ticker_row, 9).Value = ticker_symbol
            'Debug.Print ticker_symbol
             closing_price = Cells(i, 6).Value
             yearly_change = closing_price - opening_price
             Cells(ticker_row, 10).Value = yearly_change
             'avoid divide by zero scenarios for percentage changes
                If (opening_price = 0 And closing_price = 0) Then
                    percentage_change = 0
                ElseIf (opening_price = 0 And closing_price <> 0) Then
                    percentage_change = 0
                Else
                    percentage_change = (closing_price - opening_price) / opening_price
                    Cells(ticker_row, 11).Value = percentage_change
                    Cells(ticker_row, 11).NumberFormat = "0.00%"
                End If
            'finish summing up the volume by adding in the last row in the series of ticker symbols
             trading_volume = trading_volume + Cells(i, 7).Value
             Cells(ticker_row, 12).Value = trading_volume
            'Debug.Print ticker_row
             ticker_row = ticker_row + 1
            'increment the opening price to get first row for next ticker
            opening_price = Cells(i + 1, 3)
            trading_volume = 0
        Else
            trading_volume = trading_volume + Cells(i, 7).Value
            'Debug.Print trading_volume
        End If
     Next i
     
'Only iterate through the calculated percentages rather than all rows in sheet
Dim last_percent_change As Long
last_percent_change = ws.Cells(Rows.Count, 9).End(xlUp).Row
Debug.Print last_percent_change
Dim j As Long
For j = 2 To last_percent_change
    If Cells(j, 11).Value >= 0 Then
    Cells(j, 11).Interior.ColorIndex = 4
    Else: Cells(j, 11).Interior.ColorIndex = 3
    End If
Next j

Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

Dim P As Long
Debug.Print last_percent_change
'Another loop to get the max values for each
For P = 2 To last_percent_change
    If Cells(P, 11).Value = Application.WorksheetFunction.Max(ws.Range("K2:K" & last_percent_change)) Then
        Cells(2, 16).Value = Cells(P, 9).Value
        Cells(2, 17).Value = Cells(P, 11).Value
        Cells(2, 17).NumberFormat = "0.00%"
    ElseIf Cells(P, 11).Value = Application.WorksheetFunction.Min(ws.Range("K2:K" & last_percent_change)) Then
        Cells(3, 16).Value = Cells(P, 9).Value
        Cells(3, 17).Value = Cells(P, 11).Value
        Cells(3, 17).NumberFormat = "0.00%"
    ElseIf Cells(P, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & last_percent_change)) Then
        Cells(4, 16).Value = Cells(P, 9).Value
        Cells(4, 17).Value = Cells(P, 12).Value
    End If
   Next P
          
 Next ws
End Sub


