Sub stocks_test()

Dim i As Long
Dim ticker As String
Dim begin_Year As Currency
Dim end_Year As Currency
Dim year_Change As Currency
Dim total_Volume As LongLong
Dim percent_Change As Double
Dim ws As Worksheet
Dim summary_Row As Integer
Dim great_Increase As Currency
Dim great_Decrease As Currency
Dim large_Volume As LongLong
Dim ticker_percent_Increase As String
Dim ticker_percent_Decrease As String
Dim ticker_total_Volume As String

'being looking through each sheet

For Each ws In ActiveWorkbook.Worksheets

'you must activate each of the worksheets first

ws.Activate

'make headings for the results summary

Range("I" & 1).Value = "Ticker"
Range("J" & 1).Value = "Yearly Change"
Range("K" & 1).Value = "Percentage Change"
Range("L" & 1).Value = "Total Stock Volume"
Range("K:K").NumberFormat = "0.00%"

'initialize variables

i = 2
summary_Row = 2
great_Decrease = 0
great_Increase = 0
large_Volume = 0

'Look through the column of stock ticker name until there is no tickers

While Not (Range("A" & i).Value = "")
'  as the data is sequential by date then this pulls the begin_Year value
    If Range("A" & i - 1).Value <> Range("A" & i).Value Then
        begin_Year = Range("c" & i).Value
    End If
'if we know that the ticker changes to the next ticker on next time through then we are on the last date.  Print results.

    If Range("A" & i + 1).Value <> Range("A" & i).Value Then
        ticker = Cells(i, 1).Value
        end_Year = Range("F" & i).Value
        year_Change = end_Year - begin_Year

'this avoids a division by zero error
        If begin_Year <> 0 Then
            percent_Change = year_Change / begin_Year * 100
        Else
            percent_Change = 0
        End If

'  this outputs the results using the increment for the summary row to place the summary
        total_Volume = total_Volume + Range("G" & i).Value
        Range("I" & summary_Row).Value = ticker
        Range("J" & summary_Row).Value = year_Change
        Range("K" & summary_Row).Value = percent_Change & " %"
        Range("L" & summary_Row).Value = total_Volume

' set the color of the cells depending on their value
        If year_Change <= 0 Then
            Range("J" & summary_Row).Interior.ColorIndex = 3
        Else
            Range("J" & summary_Row).Interior.ColorIndex = 4
        End If

' Save the largest and smallest values from the data on each worksheet.
        If total_Volume > large_Volume Then
            large_Volume = total_Volume
            ticker_total_Volume = ticker
        End If

        If great_Increase < percent_Change Then
            great_Increase = percent_Change
            ticker_percent_Increase = ticker
        End If

        If great_Decrease > percent_Change Then
            great_Decrease = percent_Change
            ticker_percent_Decrease = ticker
        End If

' increment the new summary row so new data is below.  initialize total volume value
        summary_Row = summary_Row + 1
        total_Volume = 0

    Else

'we still need to total the volume for the ticker even though it isn't the end of the year
        total_Volume = total_Volume + Range("G" & i).Value
    End If
    i = i + 1
Wend

' Make highest level summary table for the data
Range("O1").Value = "ticker"
Range("P1").Value = "Value"

Range("N" & 2).Value = "Greatest % increase"
Range("N" & 3).Value = "Greatest % decrease"
Range("N" & 4).Value = "Greatest Total Volume"

Range("O" & 2).Value = ticker_percent_Increase
Range("O" & 3).Value = ticker_percent_Decrease
Range("O" & 4).Value = ticker_total_Volume

Range("P2").Value = great_Increase & " %"
Range("P2").NumberFormat = "0.00%"
Range("P3").Value = great_Decrease & " %"
Range("P3").NumberFormat = "0.00%"
Range("P" & 4).Value = large_Volume

Next ws

End Sub
