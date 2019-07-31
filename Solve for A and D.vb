'Solve (A&D) for the Unique symbol and total stock volume of the stock.

Sub volumetotal()

  ' Set an initial variable for holding the symbol
  Dim Symbol As String

  ' Set an initial variable for holding the total volume per symbol
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for each symbol in the summary table
  Dim Vsum As Integer
  Vsum = 2

   'Establishing last row for loop
    Lastrow = Range("A1", Range("A1").End(xlDown)).Rows.Count
    Range("A1").Select
    
  ' Loop through all sybmols
  For i = 2 To Lastrow

    ' Check if we are still within the same symbol, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Symbol
      Symbol = Cells(i, 1).Value

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

      ' Print the Symbol in the Summary Table
      Range("I" & Vsum).Value = Symbol

      ' Print the Volume Total to the Summary Table
      Range("L" & Vsum).Value = Volume_Total

      ' Add one to the summary table row
      Vsum = Vsum + 1
      
      ' Reset the Volume Total
      Volume_Total = 0

    ' If the cell immediately following a row is the same symbol...
    Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If

  Next i

End Sub