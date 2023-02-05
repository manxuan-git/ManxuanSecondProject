Sub Alphabetical_testing()

'set the ticket name in column i to l(column 9 to 12)
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = " Greatest % Increase"
Cells(3, 15).Value = " Greatest % Decrease"
Cells(4, 15).Value = " Greatest Total Volume"

'populate Ticker in column i( column 9) and Total in column l ( column12)
LastRow = Cells(Rows.Count, 1).End(xlUp).row

Dim ccticker As String
Dim ccTotal As Double
ccTotal = 0
Dim ccOpen As Double
Dim ccClose As Double
Dim ccYearlyChange As Double
Dim ccPercentChange As Double
Dim greatestTotalValueTicker As String
Dim greatestTotalValue As Double
Dim greatestIncreaseTicker As String
Dim greatestIncreaseValue As Double
Dim greatestDecreaseTicker As String
Dim greatestDecreaseValue As Double
Dim ccRows As Integer
ccRows = 2
Dim row As Long

'get the first open
ccOpen = Cells(2, 3).Value
'start loop
For row = 2 To LastRow

  If Cells(row + 1, 1).Value <> Cells(row, 1).Value Then
 
     'get the first close and loop the rest open and close
     ccClose = Cells(row, 6).Value
     ccYearlyChange = ccClose - ccOpen
     Cells(ccRows, 10).Value = ccYearlyChange
     'calculate and print the percent change
     ccPercentChange = ccYearlyChange / ccOpen
     'print percent change with % sign
     Cells(ccRows, 11).Value = FormatPercent(ccPercentChange)
     'identify next open
     ccOpen = Cells(row + 1, 3).Value
     'identify and print ticker name and total amount
     ccName = Cells(row, 1).Value
     ccTotal = ccTotal + Cells(row, 7).Value
     Cells(ccRows, 9).Value = ccName
     Cells(ccRows, 12).Value = ccTotal
     ccRows = ccRows + 1
     ccTotal = 0
  Else
    ccTotal = ccTotal + Cells(row, 7).Value
  End If
  
 Next row
 
 'calculate greatest increase & decrease & total volume value in column O to Q
 LastR = Cells(Rows.Count, 9).End(xlUp).row
 Dim r As Integer
 For r = 2 To LastR
 
    If greatestIncreaseValue < Cells(r, 11).Value Then
        greatestIncreaseValue = Cells(r, 11).Value
        greatestIncreaseTicker = Cells(r, 9).Value
    End If
    
    If greatestDecreaseValue > Cells(r, 11).Value Then
        greatestDecreaseValue = Cells(r, 11).Value
        greatestDecreaseTicker = Cells(r, 9).Value
    End If
    
    If greatestTotalValue < Cells(r, 12).Value Then
        greatestTotalValue = Cells(r, 12).Value
        greatestTotalValueTicker = Cells(r, 9).Value
    End If
   
   'change color in column J ( yearly change)
    If Cells(r, 10).Value > 0 Then
        Cells(r, 10).Interior.ColorIndex = 4
    Else
         Cells(r, 10).Interior.ColorIndex = 3
    End If
   'change color in column k ( percent change)
    If Cells(r, 11).Value > 0 Then
        Cells(r, 11).Interior.ColorIndex = 4
    Else
        Cells(r, 11).Interior.ColorIndex = 3
    End If
           
Next r
'print greatest increase & decrease & total volume value in column O to Q
Cells(2, 16).Value = greatestIncreaseTicker
Cells(2, 17).Value = greatestIncreaseValue
Cells(3, 16).Value = greatestDecreaseTicker
Cells(3, 17).Value = greatestDecreaseValue
Cells(4, 16).Value = greatestTotalValueTicker
Cells(4, 17).Value = greatestTotalValue


End Sub



