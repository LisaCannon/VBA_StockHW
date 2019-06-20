Attribute VB_Name = "Module1"
Sub stock_hard_challenge()
    'declare variables
    Dim ws As Worksheet
    Dim i, tickercount, vol_temp, vol_max As Long
    Dim close_temp, open_temp, difference, pct_change, _
    pct_max, pct_min As Double
    Dim summary_row As Integer
    Dim ticker_temp  As String
    'applies code within loop to all worksheets
    For Each ws In Worksheets
        ws.Activate
        'labels
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yearly Change"
        Cells(1, 11).Value = "Percent Change"
        Cells(1, 12).Value = "Total Stock Volume"
        Cells(2, 14).Value = "Greatest Percent Increase"
        Cells(3, 14).Value = "Greatest Percent Decrease"
        Cells(4, 14).Value = "Greatest Total Volume"
        Cells(1, 15).Value = "Ticker"
        Cells(1, 16).Value = "Value"
        'initialize value
        summary_row = 2
        vol_temp = 0
        open_temp = Cells(2, 3).Value
        Cells(4, 16).Value = 0
        pct_max = 0
        pct_min = 0
        'find number of rows
        tickercount = Cells(Rows.Count, 1).End(xlUp).Row
        'loop through each row of data
        For i = 2 To tickercount
            ticker_temp = Cells(i, 1).Value
            'checking to see if a new ticker is starting on the next row
            If Cells(i + 1, 1) <> ticker_temp Then
                'when the row is the last entry for a ticker,
                'the total is printed on the summary table
                vol_temp = vol_temp + CLng(Cells(i, 7).Value)
                Range("I" & summary_row).Value = ticker_temp
                Range("L" & summary_row).Value = vol_temp
                'check to see if the current ticker's volume is the max
                If vol_temp > Cells(4, 16).Value Then
                    Cells(4, 16).Value = vol_temp
                    Cells(4, 15).Value = ticker_temp
                End If
                'reset the temporary volume variable for the next ticker
                vol_temp = 0
                'calculate the annual change and percent change
                If i > 2 Then
                    close_temp = Cells(i, 6).Value
                    difference = close_temp - open_temp
                    'this is necessary if the opening value for the first
                    'date is zero
                    If open_temp <> 0 Then
                        pct_change = difference / open_temp
                    Else
                        pct_change = 0
                    End If
                    Range("J" & summary_row).Value = difference
                    Range("K" & summary_row).Value = pct_change
                    'check to see if the percent change is the largest
                    'positive or negative change
                    If pct_change > pct_max Then
                        pct_max = pct_change
                        Cells(2, 16).Value = pct_max
                        Cells(2, 15).Value = ticker_temp
                    ElseIf pct_change < pct_min Then
                        pct_min = pct_change
                        Cells(3, 16).Value = pct_min
                        Cells(3, 15).Value = ticker_temp
                    End If
                    'conditional formatting
                    If difference < 0 Then
                        Range("J" & summary_row).Interior.ColorIndex = 3
                    Else
                        Range("J" & summary_row).Interior.ColorIndex = 4
                    End If
                    'Because the next row starts a new ticker, you grab the
                    'opening value to keep until you run through the full year
                    'of the net ticker to get the year's closing value
                    open_temp = Cells(i + 1, 3).Value
                End If
                'because a new ticker row was added to the summary table,
                'we add 1 to the summary row
                summary_row = summary_row + 1
            'when you aren't at the last row for a ticker,
            'you add to the running volume
            Else
                vol_temp = vol_temp + CLng(Cells(i, 7).Value)
            End If
        Next i
        'formatting percent fields
        Range("K:K").NumberFormat = "0.00%"
        Range("P2:P3").NumberFormat = "0.00%"
        Range("I:P").EntireColumn.AutoFit
    Next
End Sub







