Sub stock_summary()

   'Loop through worksheets
   For Each ws In Worksheets

  ' Set a variable for holding ticker symbol
  Dim Ticker As String

  ' Set a variable for holding yearly change (beginning of year to end of year)
  Dim Year_Change As Double
  Year_Change = 0

  ' Set a variable for holding percent change (beginning of year to end of year)
  Dim Pct_Change As Double
  Pct_Change = 0

  ' Set a variable for holding total stock volume
  Dim Total_Vol As LongLong
  Total_Vol = 0

  ' Keep track of location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

   'Set summary table
   ws.Range("I1").Value = "Ticker"
   ws.Range("J1").Value = "Yearly Change"
   ws.Range("K1").Value = "Percent Change"
   ws.Range("L1").Value = "Total Stock Volume"

  ' Loop through all ticker values
  LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
  For i = 2 To LastRow

    ' Set a variable for holding beginning of year amount
    Dim Year_Start As Double
   Year_Start = ws.Cells(i, 3).Value

    ' Check if same ticker name, if NOT:
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      ' Set the ticker symbol
      Ticker = ws.Cells(i, 1).Value

      ' Determine year change
      Year_Change = ws.Cells(i, 6).Value - Year_Start 

      ' Determine percent change
      Pct_Change = Year_Change/Year_Start 

      ' Add to stock volume
      Total_Vol = Total_Vol + ws.Cells(i, 7).Value

      ' Print ticker in the Summary Table
      ws.Range("I" & Summary_Table_Row).Value = Ticker

      ' Print year change to the Summary Table
      ws.Range("J" & Summary_Table_Row).Value = Year_Change

      ' Print percent change to the Summary Table
      ws.Range("K" & Summary_Table_Row).Value = Pct_Change
      ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
      
      ' Print total stock volume to the Summary Table
      ws.Range("L" & Summary_Table_Row).Value = Total_Vol

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the year change, year start, percent change, and total stock volume
      Year_Change = 0
      Year_Start = ws.Cells(i, 3).Value
      Pct_Change = 0
      Total_Vol = 0

    ' If the cell immediately following a row is the same ticker name:
    Else

      ' Add to stock volume
      Total_Vol = Total_Vol + ws.Cells(i, 7).Value


    End If

  Next i

  '-------------------COLOR FORMATTING---------------------------

   ' Loop through all ticker yearly change values
  LastSumRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
  For j = 2 To LastSumRow

    ' Check if year change is positive, if YES:
    If ws.Cells(j, 10).Value >= 0 Then

        'set interior cell color to green
        ws.Cells(j, 10).Interior.ColorIndex = 4

    ' if year change is negative:
    Else

        'set interior cell color to red
        ws.Cells(j, 10).Interior.ColorIndex = 3

    End if

    Next j

  '-----------------COLUMN FIT FORMATTING-------------------------
    ws.Range("J1:L1").EntireColumn.AutoFit

    Next ws

End Sub
