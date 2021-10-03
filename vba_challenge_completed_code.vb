Sub vba_challenge()

Dim ws As Worksheet
For Each ws In Worksheets
ws.Activate

Range("I1:L1").Interior.ColorIndex = 6
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Price Change"
Cells(1, 11).Value = "Yearly Percent Change"
Cells(1, 12).Value = "Total Volume"

' Bonus
Range("N2:N4").Interior.ColorIndex = 6
Range("O1:P1").Interior.ColorIndex = 6
Cells(2, 14).Value = "Greatest Percent Increase"
Cells(3, 14).Value = "Greatest Percent Decrease"
Cells(4, 14).Value = "Greatest Total Volume"
Cells(1, 15).Value = "Ticker"
Cells(1, 16).Value = "Value"
      
' Define variables
Dim ticker As String
Dim totalvolume As Double
Dim lastrow As Long
Dim yearlychange As Double
Dim yearopen As Double
Dim yearclose As Double
Dim yearpercent As Double
Dim greatestincrease As Double
Dim greatestdecrease As Double
Dim greatestvolume As Double
  
Dim table As Long
dim opentable as long
table = 2
opentable = 2
totalvolume = 0
greatestincrease = 0
greatestdecrease = 0
greatestvolume = 0

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
totalvolume = totalvolume + ws.Cells(i, 7).Value
ws.Range("I" & table).Value = ticker             
ws.Range("L" & table).Value = totalvolume       
totalvolume = 0

yearopen = ws.Range("C" & opentable)
yearclose = ws.Range("F" & i)
yearlychange = yearclose - yearopen
ws.Range("J" & table).Value = yearlychange

If yearopen = 0 Then
yearpercent = 0
Else
yearopen = ws.Range("C" & opentable)
yearpercent = yearlychange / yearopen
End If

ws.Range("K" & table).NumberFormat = "0.00%"
ws.Range("K" & table).Value = yearpercent

If ws.Range("J" & table).Value >= 0 Then
ws.Range("J" & table).Interior.ColorIndex = 4
Else
ws.Range("J" & table).Interior.ColorIndex = 3
End If

table = table + 1
opentable = i + 1

Else
totalvolume = totalvolume + ws.Cells(i, 7).Value
End If

Next i

lastrowvalue = ws.Cells(Rows.Count, 11).End(xlUp).Row
For j = 2 To lastrowvalue

ws.Range("p2").NumberFormat = "0.00%"
ws.Range("p3").NumberFormat = "0.00%"
        
If ws.Range("K" & j).Value > greatestincrease Then
greatestincrease = ws.Range("K" & j).Value
ws.Range("P2").Value = greatestincrease
ws.Range("O2").Value = ws.Range("I" & j).Value
End If
            
If ws.Range("K" & j).Value < greatestdecrease Then
greatestdecrease = ws.Range("K" & j).Value
ws.Range("P3").Value = greatestdecrease
ws.Range("O3").Value = ws.Range("I" & j).Value
End If

If ws.Range("L" & j).Value > greatestvolume Then
greatestvolume = ws.Range("L" & j).Value
ws.Range("P4").Value = greatestvolume
ws.Range("O4").Value = ws.Range("I" & j).Value
End If

Next j
Next ws
End Sub
