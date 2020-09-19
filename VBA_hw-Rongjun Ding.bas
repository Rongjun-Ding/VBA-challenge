Attribute VB_Name = "Module1"
Sub The_VBA_of_Wall_Street()
'create varibales to analyze the stock

Range("I1").Value = "TICKER"
Range("J1").Value = "YEARLY CHANGE"
Range("K1").Value = "PERCENT CHANGE"
Range("L1").Value = "TOTAL STOCK VOLUME"
Range("P1").Value = "Ticker"
Range("Q1").Value = "VALUE"
Range("O2").Value = "Greatest % Increase"
Range("O3").Value = "Greatest % Decrease"
Range("O4").Value = "Greatest Total Volume"

'create varibles to help analyzing data
Dim ticker As String
Dim yearlychange As Double
Dim percentchange As Double
percentchange = 0#
Dim totalstock As Single
totalstock = 0
Dim startrow As Long
startrow = 2
Dim tickercount As Long
tickercount = 2
Dim O As Double
'assign the very first value of the first value to O
O = Cells(2, 3).Value
Dim C As Double

Cells(2, 17).Value = 0
Cells(3, 17).Value = 0
Cells(4, 17).Value = 0

'determine how many rows in total

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

'find all different types of tickers in stock market

For startrow = 2 To lastrow
'use if statement to find out all different tickers
If Cells(startrow, 1).Value <> Cells(startrow + 1, 1).Value Then
ticker = Cells(startrow, 1).Value
'add up the last stock vol for the previous ticker
totalstock = totalstock + Cells(startrow, 7).Value
'compare the open and close value
yearlychange = (Cells(startrow, 6) - O)

'assign all the value to the summary table
Range("I" & tickercount).Value = ticker
Range("L" & tickercount).Value = totalstock
Range("J" & tickercount).Value = yearlychange
'color index
If yearlychange >= 0 Then
Range("J" & tickercount).Interior.ColorIndex = 4
Else: Range("J" & tickercount).Interior.ColorIndex = 3
End If

'compare % change
percentchange = (Cells(startrow, 6) - O) / O
Range("K" & tickercount).Value = percentchange
If percentchange > Cells(2, 17).Value Then
Cells(2, 17).Value = Format(percentchange, "0%")
Cells(2, 16).Value = Range("I" & tickercount).Value
ElseIf percentchange < Cells(3, 17).Value Then
Cells(3, 17).Value = Format(percentchange, "0%")
Cells(3, 16).Value = Range("I" & tickercount).Value
End If

'format the cells' value into %
Cells(tickercount, 11) = Format(percentchange, "0%")
tickercount = tickercount + 1
'set all the counters into 0
totalstock = 0
percentchange = 0
yearlychange = 0
O = Cells(startrow + 1, 3).Value
Else
'add up the stock vol for the previous ticker
totalstock = totalstock + Cells(startrow, 7).Value

End If
Next startrow

'find that the greatest total vol
For x = 2 To lastrow
If Cells(x, 12) > Cells(4.17).Value Then
Cells(4.17).Value = Cells(x, 12).Value
Cells(4.16).Value = Cells(x, 9).Value
End If

Next x


End Sub
