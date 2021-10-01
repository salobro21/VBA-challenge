Sub vba_challenge()

' Define variables
    Dim ticker As String
    Dim totalvolume As Double
    Dim lastrow As Long
  
' Keep track of the location of the ticker in the summary table
    Dim summarytablerow As Integer
    summarytablerow = 2
    totalvolume = 0
    
' There's multiple worksheets so...
    For Each ws In Worksheets
    nextworksheet = ws.Name
  
' Loop through all rows
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow

' checking for different tickers until the last row of the worksheet
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' ticker name
    ticker = Cells(i, 1).Value

' add to volume total
    totalvolume = totalvolume + Cells(i, 7).Value

' name the ticker in the table
    Range("I" & summarytablerow).Value = ticker
             
' put in the volume total
    Range("J" & summarytablerow).Value = totalvolume
       
' Reset the Brand Total
    totalvolume = 0

' add a new row in the summary table
    summarytablerow = summarytablerow + 1
   
' if its still the same ticker

Else

' add to the volume total
    totalvolume = totalvolume + Cells(i, 7).Value
   
End If

Next i

Next ws

End Sub