Sub TickerVolume()

'to create headers for the results
Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"


  ' Set an initial variable for holding the brand name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total volume per stock ticker
  Dim Volume_Total As Double
  Volume_Total = 0

  ' Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock activity
  For i = 2 To 798000

    ' Check if we are still within the same ticker name, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

      ' Print the Ticker Name in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Brand Amount to the Summary Table
      Range("J" & Summary_Table_Row).Value = Volume_Total

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Volume total
      Volume_Total = 0

    ' If the cell immediately following a row is the same brand...
    Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(i, 7).Value

    End If

  Next i

End Sub
