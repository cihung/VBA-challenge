Sub StockAnalysis()
   'Declare and set worksheet
   Dim ws As Worksheet
   
   'Loop through all stocks for one year
   For Each ws In Worksheets
  
   'set headers for each worksheet
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"

  ' Set an initial variable for holding the brand name
  Dim Ticker As String

  'Set variables for Data
  Dim year_open As Double
  Dim year_close As Double
  Dim yearly_change As Double
  Dim percent_change As Double

  ' Set an initial variable for holding the total volume per ticker
  Dim VolPerTi As Double
  VolPerTi = 0

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

'Set value to 0 for Greatest%
ws.Range("Q2").Value = 0
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = 0
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = 0
'ws.Range("Q4").NumberFormat = "#,###,###,###,###"


  'counts the number of rows
  Dim lastrow As Long
     lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  Dim I As Long
  Dim j As Integer
  
  opnprc_indx = 2
  
  ' Loop through all tickers
  For I = 2 To lastrow

    ' Check if we are still within the same ticker, if it is not...
    If ws.Cells(I + 1, 1).Value <> ws.Cells(I, 1).Value Then
      
      ' Set the Ticker name
      Ticker = ws.Cells(I, 1).Value
      VolPerTi = VolPerTi + ws.Cells(I, 7).Value
      ' Print the Ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker
      
      ' Add to the total volume per ticker
      total_volume = VolPerTi
      ' Print the Total Stock Volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = total_volume
            
      
      'Calculate the Yearly change
      'Find open and close date
       year_close = ws.Cells(I, 6).Value
        year_open = ws.Cells(opnprc_indx, 3).Value
        'calculate yrly changb
        yearly_change = (year_close - year_open)
      'Print the Yrly change
      ws.Range("J" & Summary_Table_Row).Value = yearly_change
      opnprc_indx = I + 1
    
    'color code yrly change
  If ws.Cells(Summary_Table_Row, 10).Value > 0 Then
  'red
     ws.Cells(Summary_Table_Row, 10).Interior.Color = vbGreen
     'green
     Else
     ws.Cells(Summary_Table_Row, 10).Interior.Color = vbRed
  End If
    
    
     'check for non-divisibility conditions whe calculating the percent change
     If (year_open = 0) Then
        percent_change = 0
     Else
     percent_change = yearly_change / year_open
     End If
     
     'print yearly change for each ticker in the summary table
      ws.Range("K" & Summary_Table_Row).Value = percent_change
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

       'input greatest
       'ws.Range("Q2") = MaxP
       'increase
       If (percent_change > ws.Range("Q2").Value) Then
        ws.Range("P2").Value = Ticker
        ws.Range("Q2").Value = percent_change
       End If
       
       'decrease
       If (percent_change < ws.Range("Q3").Value) Then
       ws.Range("P3").Value = Ticker
       ws.Range("Q3").Value = percent_change
       End If
       
       'total volume
      If (VolPerTi > ws.Range("Q4").Value) Then
      ws.Range("P4").Value = Ticker
      ws.Range("Q4").Value = VolPerTi
      End If
      
      ' Reset the Ticker
      VolPerTi = 0
      'reset the opening price
      year_open = ws.Cells(I + 1, 3)
      yearly_change = 0
      
  
   ' If the cell immediately following a row is the same brand...
   ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
    Else

      ' Add to the Total Volume per Ticker
      VolPerTi = VolPerTi + ws.Cells(I, 7).Value
    
 End If
 
Next I
'Determine the color for conditional formatting for positive change in green
'find last row on summary table
  
     
Next

End Sub

