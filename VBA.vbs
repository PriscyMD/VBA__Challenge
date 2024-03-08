Sub StockData():

'Loops iterates all worksheet
  
  For Each ws In Worksheets

'Value of list

    Dim counter As Integer
    
'Variables
   
    Dim ticketname As String
    Dim volume As Double
    Dim firstprice As Double
    Dim lastprice As Double
    Dim maxchange As Double
    Dim minchange As Double
    Dim maxvol As Double
    
'Reseting
    
    maxchange = 0
    minchange = 0
    maxvol = 0
    
'Find last value

    lastrow = Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastrowj = Cells(ws.Rows.Count, 10).End(xlUp).Row
    counter = 2
    volume = 0
    firstprice = 0
    lastprice = 0
    
'Title Column
    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O3").Value = "Greatest Total Volumne"
    
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
For i = 2 To lastrow

Next i

'Conditional Formatting

  For i = 2 To lastrowj
  If ws.Cells(i, 10).Value > 0 Then
    ws.Cells(i, 10).Interior.Color = RGB(0, 255, 0)
    
  ElseIf ws.Cells(i, 10).Value = 0 Then
    ws.Cells(i, 10).Interior.Color = RGB(125, 126, 133)
    
  Else
    ws.Cells(i, 10).Interior.Color = RGB(255, 0, 0)
  
  End If
'Conditions Loops Percentage
  
  If ws.Cells(i, 11).Value > maxchange Then
    maxchange = ws.Cells(i, 11).Value
    ws.Range("P2").Value = ws.Cells(i, 9).Value
    ws.Range("Q2").Value = ws.Cells(i, 11).Value
    ws.Range("Q2").NumberFormat = "0.00%"
    
  ElseIf ws.Cells(i, 11).Value < minchange Then
    minchange = ws.Cells(i, 11).Value
    ws.Range("P3").Value = ws.Cells(i, 9).Value
    ws.Range("Q3").Value = ws.Cells(i, 11).Value
    ws.Range("Q3").NumberFormat = "0.00%"
    
  ElseIf ws.Cells(i, 12).Value > maxvol Then
    maxchange = ws.Cells(i, 11).Value
    ws.Range("P4").Value = ws.Cells(i, 9).Value
    ws.Range("Q4").Value = ws.Cells(i, 12).Value
  End If
    
'Difference between actual value
    
  If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    ticketname = ws.Cells(i, 1).Value
    
'Sum Last ticker

    vol = vol + ws.Cells(i, 7).Value
    
'Value Last Price

    lastprice = ws.Cells(i, 6).Value
    
    ws.Cells(counter, 9).Value = ticketname
    ws.Cells(counter, 12).Value = volume
    ws.Cells(counter, 10).Value = lastprice - firstprice
    ws.Cells(counter, 11).Value = ((lastprice - firstprice) / firstprice)
    ws.Cells(counter, 11).NumberFormat = "0.00%"
    
    counter = counter + 1
    
'Clear variable

    volume = 0
    
  Else
    
'Sum Vol of stock
    
    volume = volume + ws.Cells(i, 7).Value
    
  End If
  
'Condition Value

  If ws.Cells(i, 2).Value > ws.Cells(i + 1, 2).Value Then
    firstprice = ws.Cells(i + 1, 3).Value
    
'Condition applies for Ticker Change

  ElseIf i = 2 Then
  
    firstprice = ws.Range("C2").Value
    
  End If
  Next i
  Next ws
End Sub


