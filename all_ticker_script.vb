Sub all_ticker_script()


' Listing All Tickers

' Adding Column Labels
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly_Change"
Range("K1").Value = "Percent_Change"
Range("L1").Value = "Total_Stock_Volume"

    ' Variable for holding ticker name
    Dim ticker As String

    ' Tracker for location for each ticker in the summary column (I)
    Dim summary_ticker_row As Integer
    summary_ticker_row = 2

    'Setting last row of tickers for end of loop
    Dim LastRow As Long, j As Long
    last_row = Cells(Rows.Count, "A").End(xlUp).Row

      ' Loop through all tickers
      For i = 2 To last_row

        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

          ' Set the ticker
          ticker = Cells(i, 1).Value

          ' Print the ticker in the summary column (I)
          Range("I" & summary_ticker_row).Value = ticker

          ' Add one to the summary ticker column row
          summary_ticker_row = summary_ticker_row + 1

        End If

      Next i

' Yearly Changes

Dim year_open As Double
    year_open = Cells(2, 3).Value
    Dim year_close As Double
    year_close = 0
    Dim sum_change_row As Integer
    sum_change_row = 2


    'Setting last row of tickers for end of loop
    last_row = Cells(Rows.Count, "A").End(xlUp).Row

      ' Loop through all tickers
      For i = 2 To last_row

        ' Check if we are still within the same ticker, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

          ' Set the year_close
          year_close = Cells(i, 6).Value

          '' Print the yearly change in the summary column (J)
          Range("J" & sum_change_row).Value = year_close - year_open
          Range("J" & sum_change_row).NumberFormat = "#,##0.00"
          
          ' Fill yearly change cell green if positive, red if negative
          If Range("J" & sum_change_row).Value < 0 Then
             Range("J" & sum_change_row).Interior.Color = vbRed
             Range("J" & sum_change_row).Font.Color = vbWhite
          Else
            Range("J" & sum_change_row).Interior.Color = vbGreen
            Range("J" & sum_change_row).Font.Color = vbBlack
          End If

          '' Print the percent change in the summary column (K)
          If year_open = 0 Then
            Range("K" & sum_change_row).Value = 0
          Else
            Range("K" & sum_change_row).Value = (year_close - year_open) / year_open
          End If
          Range("K" & sum_change_row).NumberFormat = "0.0%;(0.0%)"
          
          ' Fill percent change cell green if positive, red if negative
          If Range("K" & sum_change_row).Value < 0 Then
             Range("K" & sum_change_row).Interior.Color = vbRed
             Range("K" & sum_change_row).Font.Color = vbWhite
          Else
            Range("K" & sum_change_row).Interior.Color = vbGreen
            Range("K" & sum_change_row).Font.Color = vbBlack
          End If
          
          ' Add one to the summary change row
          sum_change_row = sum_change_row + 1

          ' Set the next ticker year open
          year_open = Cells(i + 1, 3).Value

        End If

      Next i

'Total Stock Volume

Dim adder As LongLong
adder = 0
sum_change_row = 2

' Formatting column
Columns("L").NumberFormat = "#,##0"

'Setting last row of tickers for end of loop
last_row = Cells(Rows.Count, "A").End(xlUp).Row

      ' Loop through all tickers
      For i = 2 To last_row

      ' Check if we are still within the same ticker
        If Cells(i + 1, 1).Value = Cells(i, 1).Value Then

        'Add daily volume to adder
        adder = adder + Cells(i, 7).Value

        Else
        
        'Add daily volume to adder
        adder = adder + Cells(i, 7).Value

        ' Print the yearly volume in the summary column (L)
        Range("L" & sum_change_row).Value = adder

        ' Add one to the summary change row
        sum_change_row = sum_change_row + 1

        ' Reset adder
        adder = 0

        End If

      Next i


' Calculating Greatest Summary

Dim greatest_increase As Double
Dim greatest_decreaase As Double
Dim greatest_volume As LongLong
Dim greatest_ticker As String

greatest_increase = Cells(2, "K").Value
greatest_decrease = Cells(2, "K").Value
greatest_volume = Cells(2, "L").Value
greatest_ticker = ""

' Adding Labels
Range("N2").Value = "Greatest Increase (%)"
Range("N3").Value = "Greatest Decrease (%)"
Range("N4").Value = "Greatest Total Volume"
Range("O1").Value = "Ticker"
Range("P1").Value = "Value"

'Formatting Cells
Cells(2, "P").NumberFormat = "0.0%;[Red](0.0%)"
Cells(3, "P").NumberFormat = "0.0%;[Red](0.0%)"
Cells(4, "P").NumberFormat = "#,##0"

'Setting last row of tickers for end of loop
last_row = Cells(Rows.Count, "I").End(xlUp).Row

     ' Loop through all tickers for greatest increase
          For i = 2 To last_row

               If Cells(i, "K").Value > greatest_increase Then
               greatest_increase = Cells(i, "K").Value
               greatest_ticker = Cells(i, "I").Value
               
               End If

               Cells(2, "P").Value = greatest_increase
               Cells(2, "O").Value = greatest_ticker

          Next i

     ' Loop through all tickers for greatest increase
           For i = 2 To last_row

               If Cells(i, "K").Value < greatest_decrease Then
               greatest_decrease = Cells(i, "K").Value
               greatest_ticker = Cells(i, "I").Value
               
               End If

               Cells(3, "P").Value = greatest_decrease
               Cells(3, "O").Value = greatest_ticker

          Next i

     ' Loop through all tickers for greatest total volume

          For i = 2 To last_row

               If Cells(i, "L").Value > greatest_volume Then
               greatest_volume = Cells(i, "L").Value
               greatest_ticker = Cells(i, "I").Value
               
               End If

               Cells(4, "P").Value = greatest_volume
               Cells(4, "O").Value = greatest_ticker

          Next i
          
    MsgBox ("Script complete.")

End Sub

