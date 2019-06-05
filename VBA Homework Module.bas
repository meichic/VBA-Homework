Attribute VB_Name = "Module1"
Sub stock()

Dim ws As Worksheet

For Each ws In ActiveWorkbook.Worksheets
    ws.Activate
    
 ' Set an initial variable for holding the Ticker_Symbol
  Dim Ticker_Symbol As String
  Cells(1, 10).Value = "Ticker_Symbol"

  ' Set an initial variable for holding the Total_Stock_Volume
  Dim Stock_Volume As LongLong
  Stock_Volume = 0
  Cells(1, 11) = "Total_Stock_Volume"

  ' Keep track of the location for each Ticker_Symbol in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
 ' Determine the Last Row
    Dim LastRow As LongLong
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all stock volume
    For I = 2 To LastRow
  
  ' Check if we are still within the same Ticker_Symbol, if it is not...
     If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then

      ' Set the Ticker_Symbol
        Ticker_Symbol = Cells(I, 1).Value

      ' Add to the Stock_Volume
        Stock_Volume = Stock_Volume + Cells(I, 7).Value

      ' Print the Ticker_Symbol in the Summary Table
        Range("J" & Summary_Table_Row).Value = Ticker_Symbol

      ' Print the Stock Volume to the Summary Table
        Range("K" & Summary_Table_Row).Value = Stock_Volume

      ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Stock_Volume
        Stock_Volume = 0

    ' If the cell immediately following a row is the same brand...
        Else

      ' Add to the Total_Stock_Volume
        Stock_Volume = Stock_Volume + Cells(I, 7).Value

   End If

    Next I
   
  Next ws
    
End Sub

