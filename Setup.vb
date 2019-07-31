'//////INSTRUCTIONS/////////
'Create a script that will loop through all the stocks for one year for each run and take the following information.
'A. The ticker symbol.
'B. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
'C. The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
'D. The total stock volume of the stock.
'E. You should also have conditional formatting that will highlight positive change in green and negative change in red.
'Column Names and Setup
Sub ColumnNames()

 ' Setting Column Heading
  ' ---------------------------------------
ColumnName = Array("Ticker", "Yearly Change", "Percentage Change", "Total Stock Volume")
Range("I1:L1").Value = ColumnName

 ' Setting Column Width
  ' ---------------------------------------
Range("I1").ColumnWidth = 6.13
Range("J1").ColumnWidth = 11.56
Range("K1").ColumnWidth = 15.94
Range("L1").ColumnWidth = 16

 ' Setting cell color formatting
  ' ---------------------------------------
   'Establishing last row for loop
    Lastrow = Range("L1", Range("L1").End(xlDown)).Rows.Count
    Range("L1").Select
    
  ' Loop through all sybmols
  For i = 1 To Lastrow
    If Cells(i + 1, 10).Value > 0 Then
    Cells(i + 1, 10).Interior.ColorIndex = 4
    End If
  Next i
  For i = 1 To Lastrow
   If Cells(i + 1, 10).Value < 0 Then
    Cells(i + 1, 10).Interior.ColorIndex = 3
    End If
   Next i
   
End Sub