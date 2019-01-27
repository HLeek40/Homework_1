Attribute VB_Name = "Module1"
Sub Stock_Ticker()

  ' Set an initial variable for holding stock ticker name
  Dim Ticker_Name As String

  ' Set an initial variable for holding the total per stock ticker
  Dim Stock_Volume_Total As Double
  Stock_Volume_Total = 0

  ' Keep track of the location for each stock ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  Cells(1, 9).Value = "Ticker"
  Cells(1, 10).Value = "Total Stock Volume"

  'Find the last row in the worksheet
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Loop through all stock data
  For i = 2 To LastRow

    'Check if we are still within the same stock ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the Ticker Name
      Ticker_Name = Cells(i, 1).Value

      ' Add to the Stock Volume Total
      Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value

      ' Print the Stock Ticker in the Table
      Range("I" & Summary_Table_Row).Value = Ticker_Name

      ' Print the Total Stock Volume to the Table
      Range("J" & Summary_Table_Row).Value = Stock_Volume_Total

      ' Add one to the table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock Volume Total
      Stock_Volume_Total = 0

    ' If the cell immediately following a row is the same ticker symbol
    Else

      ' Add to the Stock Volume Total
      Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value

    End If

  Next i

End Sub

