# Homework_VBA
VBA codes 
# Alphabetical_testing & Multiple_year_stock_data

Sub Stock_Market()

Dim ws As Worksheet
Dim ticker As String
Dim i, j As Long
Dim open_price, closed_price, yearly_change, percent_change, Greatest_increase As Double
Dim count, count1, Lrow, Lrow2 As Long
Dim total_stock As Double

For Each ws In Worksheets

' definiendo valor inicial
    count = 2
    count1 = 2
    Lrow = ws.Cells(Rows.count, 1).End(xlUp).Row
    total_stock = 0
    
'definiendo headers
ws.Cells(1, 10).Value = "ticker"
ws.Cells(1, 11).Value = "yearly change"
ws.Cells(1, 12).Value = "% change"
ws.Cells(1, 13).Value = "Total Stock"

'Greatest headers
ws.Range("O2").Value = "Greatest % increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


'1st Loop

For i = 2 To Lrow
    total_stock = total_stock + ws.Cells(i, 7).Value
   
If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
    
  ticker = ws.Cells(i, 1).Value
 
      ws.Cells(count, 10).Value = ticker
      ws.Cells(count, 13).Value = total_stock
       total_stock = 0
        
    ' yearly change
      open_price = ws.Cells(count1, 3)
      closed_price = ws.Cells(i, 6).Value
      yearly_change = closed_price - open_price
    'para imprimir el valor y open price
      ws.Cells(count, 11).Value = yearly_change
      
     'fixing the error caused by dividing by 0
      If open_price = 0 Then
      ws.Cells(count, 12).Value = 0
      Else
      percent_change = yearly_change / open_price
      ws.Cells(count, 12).Value = percent_change
      ws.Cells(count, 12).NumberFormat = "#,##.00%"
      End If

    ' conditional- color
    If percent_change < 0 Then
     ws.Cells(count, 12).Interior.ColorIndex = 3
     
     Else
     ws.Cells(count, 12).Interior.ColorIndex = 4
     
     End If
       
    count = count + 1
    count1 = i + 1
    
     End If
    Next i
      
     'defining the greatest
     Lrow2 = ws.Cells(Rows.count, 12).End(xlUp).Row

     For j = 2 To Lrow2
     ' Greatest % increase
     If ws.Cells(j, 12).Value > ws.Cells(2, 17).Value Then
     ws.Cells(2, 17).Value = ws.Cells(j, 12).Value
     ws.Cells(2, 16).Value = ws.Cells(j, 10).Value
     ws.Cells(2, 17).NumberFormat = "#,##.00%"
     End If
     
     'Greatest % decrease
     If ws.Cells(j, 12).Value < ws.Cells(3, 17).Value Then
     ws.Cells(3, 17).Value = ws.Cells(j, 12).Value
     ws.Cells(3, 16).Value = ws.Cells(j, 10).Value
     ws.Cells(3, 17).NumberFormat = "#,##.00%"
     End If
     
     'Greatest Total volume
     If ws.Cells(j, 13).Value > ws.Cells(4, 17).Value Then
     ws.Cells(4, 17).Value = ws.Cells(j, 13).Value
     ws.Cells(4, 16).Value = ws.Cells(j, 10).Value
     ws.Cells(4, 17).NumberFormat = "#,###,###."
     
     End If
     
  
     
   

 
    Next j
   
 Next ws
End Sub

