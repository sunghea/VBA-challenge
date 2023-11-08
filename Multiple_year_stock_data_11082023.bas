Attribute VB_Name = "Module1"
Sub year_stock()

 Dim ws As Worksheet
  ' Set an initial variable for holding the ticker_Name
  Dim ticker_Name As String
  ' Set the last row
  Dim last_row As Long
  ' Set the opening price at the beginning of a given year counter
  Dim op_counter As Integer
  ' Set the opening price
  Dim opening_price As Double
  ' Set the closing price
  Dim closing_price As Double
  ' Set the Yearly change
  Dim Yearly_changee As Double
  ' Set the Yearly change
  Dim Percentage_change As Double
  ' Set the Percentage change
  Dim ticker_Total As Double
  ticker_Total = 0
  'Keep track of the location for each ticker name in the summary table
  Dim Summary_Table_Row As Integer
  Dim totalSum As Integer

  
 ' Apply the same content by Worksheets
For Each ws In Worksheets
    
sheetname = ws.Name
MsgBox (sheetname)
   
    
 'Insert Summary_Table that into each worksheet
  ws.Cells(1, 9).Value = "Ticker"
  ws.Cells(1, 10).Value = "opening price"
  ws.Cells(1, 11).Value = "closing price"
  ws.Cells(1, 12).Value = "Yearly changee"
  ws.Cells(1, 13).Value = "Percent Change"
  ws.Cells(1, 14).Value = "Total Stock Volume"
     
  ' The last row check
  last_row = ws.Range("A1").End(xlDown).Row
  MsgBox (last_row)
  
  Summary_Table_Row = 2
  totalSum = 0

  
 ' Loop through all  ticker
  
  For i = 2 To last_row

    ' Check if we are still within the same  ticker name, if it is not...
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   
     ' Set the  ticker name name
      ticker_Name = ws.Cells(i, 1).Value
    ' Print the  ticker name in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = ticker_Name
      
    ' Set the closing_price
      closing_price = ws.Cells(i, 6).Value
    ' Print the  closing_price to the Summary Table
      ws.Range("k" & Summary_Table_Row).Value = closing_price
      
    ' Set the Yearly_changee
      Yearly_changee = closing_price - opening_price
    ' Print the  Yearly_changee to the Summary Table
      ws.Range("l" & Summary_Table_Row).Value = Yearly_changee
      
      ' if it is negative, it is light green
        If ws.Range("L" & Summary_Table_Row) < 0 Then
           ws.Range("L" & Summary_Table_Row).Interior.Color = RGB(255, 0, 0) ' red
      
      ' if it is positive, it is light green
        ElseIf ws.Range("L" & Summary_Table_Row) > 0 Then
           ws.Range("l" & Summary_Table_Row).Interior.Color = RGB(0, 255, 0) ' green
      
      ' if it is zero, it is default
        Else
           ws.Range("l" & Summary_Table_Row).Interior.Color = xlNone
        End If
      
          
     ' Set the Percentage_change
       Percentage_change = Yearly_changee / opening_price
     ' Print the  Percentage_change to the Summary Table
       ws.Range("M" & Summary_Table_Row).NumberFormat = "0.00%"
       ws.Range("M" & Summary_Table_Row).Value = Percentage_change
                  
      ' Add to the ticker_Total
       ticker_Total = ticker_Total + ws.Cells(i, 7).Value
      ' Print the ticker_Name Amount to the Summary Table
        ws.Range("N" & Summary_Table_Row).Value = ticker_Total
       'ws.Range("N" & Summary_Table_Row).NumberFormat = "[$CAD]#,##0.00"
      
      ' Add one to the summary table row
       Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the ticker_Total
        ticker_Total = 0
      ' Reset the ticker number
        totalSum = 0

    ' If the cell immediately following a row is the same brand...
    Else
      ' Add to the ticker_Total
      ticker_Total = ticker_Total + ws.Cells(i, 7).Value
      
      ' Add the ticker number
      totalSum = totalSum + 1
        
      ' Setting the initial value when the ticketter is newly started Set the opening_price
         If (totalSum = 1) Then
         opening_price = ws.Cells(i, 3).Value
      ' Print the  opening_price to the Summary Table
         ws.Range("j" & Summary_Table_Row).Value = opening_price
              
         End If
 

    End If
    

  Next i
 


Next ws
             

End Sub





