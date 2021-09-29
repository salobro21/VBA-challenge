Attribute VB_Name = "Module1"
Sub vba_challenge()

  ' Ticker name
  Dim ticker As String

  ' Set an initial variable for the total volume
  Dim volumetotal As Long
  
  ' Keep track of the location of the ticker in the summary table
  Dim summarytablerow As Integer
  summarytablerow = 2

  ' Loop through all rows
  For i = 2 To 70926

    ' checking for different tickers
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' ticker name
      ticker = Cells(i, 1).Value

      ' add to volume total
      ticker = volumetotal + Cells(i, 3).Value

      ' name the ticker in the table
      Range("I" & summarytablerow).Value = ticker
      
      ' put in the volume total
      Range("J" & summarytablerow).Value = volumetotal

      ' add a new row in the summary table
      summarytablerow = summarytablerow + 1
      
      ' Reset the Brand Total
      totalvolume = 0

    ' if its still the same ticker
    Else

      ' add to the volume total
      totalvolume = totalvolume + Cells(i, 3).Value

    End If

  Next i

End Sub

