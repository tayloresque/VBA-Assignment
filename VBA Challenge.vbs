Sub StockData()
'Set Variable to loop through worksheets
Dim ws1 As Worksheet
For Each ws1 In Worksheets

'Headers
ws1.Range("I1") = "Stock Ticker"
ws1.Range("J1") = "Yearly Change"
ws1.Range("K1") = "Percent Change"
ws1.Range("L1") = "Total Stock Volume"
ws1.Range("P1") = "Stock Ticker"
ws1.Range("Q1") = "Value"
ws1.Range("O2") = "Greatest % Increase"
ws1.Range("O3") = "Greatest % Decrease"
ws1.Range("O4") = "Greatest Total Volume"

'set an initial variable for ticker
Dim t As Long

'Set open row as variable to use in yearly change calculation
Dim Open_Row As Long
Open_Row = 2

'Set Percent Change as variable
Dim Percent_Change As Double

'set volume as variable
Dim Total_Volume As LongLong
Total_Volume = 0

'Declare table
Dim Table As Integer
Table = 2

'keep track of the yearly change of each stock
Dim Yearly_Change As Double
Yearly_Change_Row = 2

'Set open and close prices as variables
Dim Open_Price As Double
Open_Price = 0
Dim Close_Price As Double
Close_Price = 0

'Declare a last row
Last_Row = ws1.Cells(Rows.Count, "A").End(xlUp).Row

'loop through all stocks
For t = 2 To Last_Row

'Check if we are still in the same stock
   If ws1.Cells(t + 1, 1).Value <> ws1.Cells(t, 1).Value Then
   
      ' Set the Stock name
      Stock_Name = ws1.Cells(t, 1).Value
      
      'Calculate Yearly change
      Close_Price = ws1.Cells(t, 6).Value
      Open_Price = ws1.Cells(Open_Row, 3).Value
      Yearly_Change = Close_Price - Open_Price

      'Calculate Percent Change
      Percent_Change = Yearly_Change / Open_Price
      
      'add to stock volume
      Total_Volume = Total_Volume + ws1.Cells(t, 7).Value
      
      'Print Stock Name in the column
      ws1.Range("I" & Table).Value = Stock_Name
      
      'Print Yearly Change
      ws1.Range("J" & Table).Value = Yearly_Change
      ws1.Range("J" & Table).NumberFormat = "0.00"
      
      'Print Percent Change
      ws1.Range("K" & Table).Value = Percent_Change
      ws1.Range("K" & Table).NumberFormat = "0.00%"
      
      
      'Conditional Formatting
      Select Case Yearly_Change
        Case Is > 0
            ws1.Range("J" & Table).Interior.ColorIndex = 4
            Case Is < 0
            ws1.Range("J" & Table).Interior.ColorIndex = 3
            Case Else
            ws1.Range("J" & Table).Interior.ColorIndex = 0
            
            End Select

      
      'Update Open Row Variable
      Open_Row = t + 1
      
      'Print Volume in column
     ws1.Range("L" & Table).Value = Total_Volume
     
     'Add one row for total volume
     Table = Table + 1
     
     'Reset volume total
      Total_Volume = 0
      
      'If the cell following the row is the same stock
      Else
      
      'Add to volume
      Total_Volume = Total_Volume + ws1.Cells(t, 7).Value
      
End If
Next t

'Print Greatest % increase + decrease + volume
      ws1.Range("Q2") = "%" & WorksheetFunction.Max(ws1.Range("K2:K" & Last_Row)) * 100
      ws1.Range("Q3") = "%" & WorksheetFunction.Min(ws1.Range("K2:K" & Last_Row)) * 100
      ws1.Range("Q4") = WorksheetFunction.Max(ws1.Range("L2:L" & Last_Row))
      
      'Find Greatest % increase + decrease + volume stock names
      Increase_Index = WorksheetFunction.Match(WorksheetFunction.Max(ws1.Range("K2:K" & Last_Row)), ws1.Range("K2:K" & Last_Row), 0)
      Decrease_Index = WorksheetFunction.Match(WorksheetFunction.Min(ws1.Range("K2:K" & Last_Row)), ws1.Range("K2:K" & Last_Row), 0)
      Greatest_Volume_Index = WorksheetFunction.Match(WorksheetFunction.Max(ws1.Range("L2:L" & Last_Row)), ws1.Range("L2:L" & Last_Row), 0)
      
      'Print Greatest % increase + decrease + volume stock names
      ws1.Range("P2").Value = ws1.Cells(Increase_Index + 1, 9)
      ws1.Range("P3").Value = ws1.Cells(Decrease_Index + 1, 9)
      ws1.Range("P4").Value = ws1.Cells(Greatest_Volume_Index + 1, 9)
      

Next ws1
End Sub
