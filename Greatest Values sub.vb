Sub greatest_calc ()

   'Loop through worksheets
   For each ws In Worksheets

  'Set a variable for holding ticker symbol for increase, decrease, and stock volume
  Dim Ticker_sum_inc As String
  Dim Ticker_sum_dec as String
  Dim Ticker_sum_vol as String

  ''Set a variable for holding increase, decrease, and stock volume
  Dim Pct_Inc as Double
  Dim Pct_Dec as Double
  Dim Great_Vol as LongLong

  Pct_Inc = 0
  Pct_Dec = 0
  Great_Vol = 0

   'Set greatest value summary table
   ws.Range("O2").Value = "Greatest % Increase"
   ws.Range("O3").Value = "Greatest % Decrease"
   ws.Range("O4").Value = "Greatest Total Volume"
   ws.Range("P1").Value = "Ticker"
   ws.Range("Q1").Value = "Value"

   ' Nested loop through all summary table values
  LastRow = ws.Cells(Rows.Count, 10).End(xlUp).Row
  For i = 2 To LastRow

  For j = 9 to 12

 ' Scan dataset for highest percent change value
    If ws.Cells(i, 11).Value > Pct_Inc Then

      ' Update the ticker symbol
      Ticker_sum_inc = ws.Cells(i, 9).Value

      ' Update percent increase value
      Pct_Inc = ws.Cells(i, 11).Value 

    'Scan for lowest percent change value
    Elseif ws.Cells(i, 11).Value < Pct_Dec Then

      ' Update the ticker symbol
      Ticker_sum_dec = ws.Cells(i, 9).Value

      ' Update percent decrease value
      Pct_Dec = ws.Cells(i, 11).Value 

    'Scan for greatest total volume
    Elseif ws.Cells(i, 12).Value > Great_Vol Then

      ' Update the ticker symbol
      Ticker_sum_vol = ws.Cells(i, 9).Value

      ' Update percent dec value
      Great_Vol = ws.Cells(i, 12).Value 

      End if

    Next j

    Next i

      ' Print greatest increase data in the Summary Table
      ws.Range("P2").Value = Ticker_sum_inc
      ws.Range("Q2").Value = Pct_Inc
      ws.Range("Q2").NumberFormat = "0.00%"

      ' Print greatest decrease data in the Summary Table
      ws.Range("P3").Value = Ticker_sum_dec
      ws.Range("Q3").Value = Pct_Dec
      ws.Range("Q3").NumberFormat = "0.00%"

      ' Print greatest volume data in the Summary Table
      ws.Range("P4").Value = Ticker_sum_vol
      ws.Range("Q4").Value = Great_Vol
      
  '-----------------COLUMN FIT FORMATTING-------------------------
    ws.Range("O1:Q1").EntireColumn.AutoFit

    Next ws

End Sub