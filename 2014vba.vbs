Sub stock_data()

    Cells(1, 9).Value = "Ticker Label"
    Cells(1, 10).Value = "Volume Total"
  ' Set an initial variable for holding the ticker_label
  Dim ticker_label As String

  ' Set an initial variable for holding the total volume
  Dim total_volume As Double
  total_volume = 0

  ' Keep track of the location for each ticker label in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Loop through all stock data
  For i = 2 To 705714

    ' Check if we are still within the same ticker label, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker label
      ticker_label = Cells(i, 1).Value

      ' Add to the total_volume
      total_volume = total_volume + Cells(i, 7).Value

      ' Print the Ticker Label in the Summary Table
      Range("I" & Summary_Table_Row).Value = ticker_label

      ' Print the volume total to the Summary Table
      Range("J" & Summary_Table_Row).Value = total_volume

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the volume total
      total_volume = 0

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the volume total
      total_volume = total_volume + Cells(i, 7).Value

    End If

  Next i

End Sub

