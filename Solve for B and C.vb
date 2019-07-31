'Solve (B&C) for the Yearly and % change of the stock.

Sub PriceChanges()

  ' Set an initial variable for holding the symbol
  Dim Symbol As String

  ' Set an initial variable for opening price per symbol
  Dim OPrice As Double
   OPrice = 0
   
  ' Set an initial variable for closing price per symbol
  Dim CPrice As Double
   CPrice = 0
   
   'Set array for Open Date
   Dim ODate As Double
    ODate = 20160101
   
   'Set array for Close Date
   Dim C_Date As Double
    C_Date = 20161230
   
  ' Set an initial variable for Yearly Change per symbol
  Dim Yearly_Change As Double
   Yearly_Change = 0
  
    ' Set an initial variable for Percentage Change per symbol
  Dim Perc_Change As Double
  Perc_Change = 0

  ' Keep track of the location for each symbol in the summary table
  Dim Vsum As Integer
  Vsum = 2

   'Establishing last row for loop
    Lastrow = Range("A1", Range("A1").End(xlDown)).Rows.Count
    'Lastrow = 1000
    Range("A1").Select
    
    ' debugging by write iutput to file
    n = FreeFile()
   Open "/Users/Drew's Surface/OneDrive/Documents/Drew's Documents/Columbia/Unit 2 Homework The VBA of Wall Street/d.txt" For Output As #n
    
    OPrice = Cells(2, 3).Value
  ' Loop through all sybmols
  For i = 2 To Lastrow

    
    ' Check if we are still within the same symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
   
   Print #n, ""
   Print #n, "found", i, Cells(i, 1).Value, Cells(i, 2).Value
   Print #n, "found", i + 1, Cells(i + 1, 1).Value, Cells(i + 1, 2).Value

       ' Nested If to set Open price
        'If Cells(i + 1, 2).Value = ODate Then
       ' Print #n, "Cells(i + 1, 2).Value = ODate", Cells(i + 1, 2).Value, ODate
               
        
         ' Nested If to set Closed price
        'If Cells(i + 1, 2).Value = C_Date Then
       'Print #n, "Cells(i + 1, 2).Value = C_Date", Cells(i + 1, 2).Value, C_Date
         ' Set the Closing Price
           CPrice = Cells(i, 6).Value
           Print #n, "CPrice = Cells(i, 6).Value", CPrice, Cells(i, 6).Value
        'End If
      
      
      ' Calc to the Yearly Change
      Yearly_Change = CPrice - OPrice
      Print #n, Yearly_Change
      
      ' Calc to the Percentage Change
      Perc_Change = (Yearly_Change / OPrice)
      Print #n, Yearly_Change, OPrice, Perc_Change

      ' Print the Symbol in the Summary Table
      Range("J" & Vsum).Value = Yearly_Change

      ' Print the Volume Total to the Summary Table
      Range("K" & Vsum).Value = Perc_Change

      ' Add one to the summary table row
      Vsum = Vsum + 1
      
      ' Reset the Yearly Change
      Yearly_Change = 0
      
      ' Reset the Percentage Change
      Perc_Change = 0

    ' If the cell immediately following a row is the same symbol...
        ' Set the Opening Price
        ' new open price 29.61
           OPrice = Cells(i + 1, 3).Value
           Print #n, i + 1, "new Oprice ", Cells(i + 1, 3).Value
        'End If
    Else

      ' Calc Changes
      'Yearly_Change = CPrice - OPrice
     

    End If

  Next i
Print #n, "done"
Close #n

End Sub


