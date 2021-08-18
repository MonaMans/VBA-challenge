Attribute VB_Name = "Module1"
'Steps:
'--------------------------------------

'Create a script that will loop through all the stocks for one year and output the following:
'1. The ticker symbol
'2. Yearly change from opening price at the beginning of a given year to the closing price at the end of that year
'3. Conditional formatting highlighting positive change in green and negative change in red for Step 2
'4. The percent change from opening price at the beginning of a given year to the closing price at the end of that year
'5. The total stock volume of the stock

Sub VBA_of_Wall_Street()

For Each ws In Worksheets

' Set an initial variable for holding the ticker symbol
  Dim Ticker_Symbol As String
  
  'Set an initial variable for stock price
  Dim opening_price As Double
  Dim closing_price As Double
  Dim Yearly_change As Double
  Dim Percentage_change As Double
                  
                
  'Determine last row for a ticker symbol
  Dim LastRow As Long
  LastRow = ws.Range("A" & Rows.Count).End(xlUp).Row
   
  ' Set an initial variable for holding the total stock volume per ticker symbol
  Dim Stock_Volume_Total As Double
  Stock_Volume_Total = 0
  
  'Create summary table headers
  ws.Cells(1, lastcolumn + 9) = "Ticker"
  ws.Cells(1, lastcolumn + 10) = "Yearly Change"
  ws.Cells(1, lastcolumn + 11) = "Percent Change"
  ws.Cells(1, lastcolumn + 12) = "Total Stock Volume"

  ' Keep track of the location for each ticker symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  'Set initial opening price
  opening_price = ws.Cells(2, 3).Value

  ' Loop through all ticker symbols
  For I = 2 To LastRow
  
    ' Check if we are still within the same ticker symbol, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then

      ' Set the ticker symbol
      Ticker_Symbol = ws.Cells(I, 1).Value

      ' Add to the Total Stock volume
      Stock_Volume_Total = Stock_Volume_Total + ws.Cells(I, 7).Value
      
      'Set the closing price
      closing_price = ws.Cells(I, 6).Value
           
                  
      'Add the yearly change
      Yearly_change = closing_price - opening_price
      ws.Cells(2, 10).NumberFormat = "0.00"
      
      'Add Percent Change
       If opening_price <> 0 Then
      Percentage_change = Yearly_change / opening_price
      ws.Range("K:K").NumberFormat = "0.00%"
      End If
      
      'Set opening price
      opening_price = ws.Cells(I + 1, 3).Value
      
                                    
      ' Print the ticker symbol in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol

      ' Print the total stock volume amount to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Stock_Volume_Total
      
      ' Print the yearly change amount to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Yearly_change
      
      ' Print the percentage change amount to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Percentage_change
      
   
     'Apply conditional formatting to yearly change
     If ws.Range("J" & Summary_Table_Row).Value >= 0 Then
     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
     Else
     ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
     End If
     
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
    ' Reset the total stock volume
      Stock_Volume_Total = 0
      Else
    
      ' Add to the total stock volume
      Stock_Volume_Total = Stock_Volume_Total + ws.Cells(I, 7).Value
      
    End If
    
           
  Next I
  
  Next ws
    
End Sub



