
Sub ticker()
' Set a variable for specifying the column of interest
  Dim count, i, volume, total, column, rc As Long
  Dim openval, closeval, variance, percentchange As Double
  
 For Each ws In Worksheets ' This for loop repeats for each worksheet in the workbook
  openval = 1
  closeval = 0
  column = 1
  count = 1
  tem = 1
  volume = 0
  

 rc = ws.Cells(Rows.count, 1).End(xlUp).Row ' counts the number of columns in the worksheet
  volume = 0
  ws.Cells(1, 10).Value = "Ticker"            'Names the output columns in the worksheet
  ws.Cells(1, 11).Value = "Yearly change"     'Names the output columns in the worksheet
  ws.Cells(1, 12).Value = "Percentage change"  'Names the output columns in the worksheet
  
  ws.Cells(1, 13).Value = "Total stock volume"  'Names the output columns in the worksheet
  
 
  
  For i = 2 To rc                       ' This for loop repeats for all columns in each worksheet
    If ws.Cells(i + 1, column).Value = ws.Cells(i, column).Value Then 'Condition for checking consecutive stock names
       If (tem = 1) Then
        openval = ws.Cells(i, column + 2).Value ' Records the stock value beginning of the year
        volume = ws.Cells(count + 1, 13).Value
        tem = 0
       End If
       ws.Cells(count + 1, 13).Value = ws.Cells(count + 1, 13) + ws.Cells(i, 7).Value  ' counting the volume of each stock
     
           ' Searches for when the value of the next cell is different than that of the current cell
    ElseIf ws.Cells(i + 1, column).Value <> ws.Cells(i, column).Value Then  'Condition for checking consecutive stock names
     
      ws.Cells(count + 1, 10).Value = ws.Cells(i, column).Value
      ws.Cells(count + 1, 13).Value = ws.Cells(count + 1, 13) + ws.Cells(i, 7).Value
       
      closeval = ws.Cells(i, column + 5).Value   ' Records the stock  value end of the year
      variance = closeval - openval              ' Difference in closing to opening value for each stock
      
      If (variance < 0) Then
        ws.Cells(count + 1, 11).Interior.ColorIndex = 3    ' Stocks in negative coded with Red
      Else
        ws.Cells(count + 1, 11).Interior.ColorIndex = 4     ' Stocks in positive coded with Green
      End If
      
       ws.Cells(count + 1, 11).Value = variance
      
                
        If (openval <> 0) Then
         percentchange = (variance / openval) * 100           'Percentage change in stock value
        End If
      ws.Cells(count + 1, 12).Value = Str(Round(percentchange, 2)) + "%"
      
   
       percentchange = 1          ' Resetting the variable values for next work sheet
       openval = 1
       closeval = 0
       count = count + 1
       tem = 1
    End If

 Next i
 
 
  
Next ws     ' goes to the beginning to repeat the code for next worksheet in the same workbook
End Sub







