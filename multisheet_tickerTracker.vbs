Sub buttonClick()
'function to run though worksheets
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        'calls function that assesses stock tickers
        Call tickerTracker
    Next
    Application.ScreenUpdating = True
End Sub

Sub tickerTracker()

    'set variable for line count
    Dim lineCount As Double
    'set variable to hold each credit card brand
    Dim ticker As String
    'set variable and initialize to hold ticker total volume
    Dim totalVol As Double
    totalVol = 0
    'set variable to track ticker summary location
    Dim summary_table_row As Integer
    summary_table_row = 2
    'set variable for ticker opening value
    Dim tickerOpen As Double
    tickerOpen = Range("C2").Value
    'set variable for ticker closing value
    Dim tickerClose As Double
    tickerClose = 0
    'set variables for "greatest" values
    Dim gpiTicker As String
    Dim gpdTicker As String
    Dim gtvTicker As String
    Dim gpiAmt As Double
    Dim gpdAmt As Double
    Dim gtvAmt As Double
    'initialize "greatest" variables
    gpiTicker = "Not Found Yet"
    gpdTicker = "Not Found Yet"
    gtvTicker = "Not Found Yet"
    gpiAmt = 0
    gpdAmt = 0
    gtvAmt = 0
    
    'set column headers for summary table
    Range("I1") = "Ticker Symbol"
    Range("J1") = "Total Stock Volume"
    Range("K1") = "Yearly Change"
    Range("L1") = "Percent Change"
    'set table column and row descriptors for "greatest"
    Range("N2") = "Greatest % Increase"
    Range("N3") = "Greatest % Decrease"
    Range("N4") = "Greatest Total Vol"
    Range("O1") = "Ticker"
    Range("P1") = "Value"
    
    'see how many lines we are analyzing per sheet #Comment out msgbox for production
    lineCount = ActiveSheet.UsedRange.Rows.Count
    'MsgBox "no. of entries is " & lineCount
    
    'loop thourgh all ticker symbols
        For i = 2 To lineCount
            'runs ONLY if next line ticker symbol does not match
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            'sets ticker symbol for summary table
            ticker = Cells(i, 1).Value
            'adds final vol to total vol
            totalVol = totalVol + Cells(i, 7).Value
            'puts ticker and tickerVol in summary columns
            Range("I" & summary_table_row).Value = ticker
            Range("J" & summary_table_row).Value = tickerVol
            'compares totalVol to gtvAmt
                If tickerVol > gtvAmt Then
                    gtvTicker = ticker
                    gtvAmt = tickerVol
                End If
            'sets ticker close value for summary columns comparisons
            tickerClose = Cells(i, 6).Value
            'sets "yearly change" value
            Range("K" & summary_table_row).Value = (tickerClose - tickerOpen)
            'format "yearly change" cell background
                If Range("K" & summary_table_row).Value < 0 Then
                    Range("K" & summary_table_row).Interior.ColorIndex = 3
                ElseIf Range("K" & summary_table_row).Value > 0 Then
                    Range("K" & summary_table_row).Interior.ColorIndex = 4
                End If
            'sets "percent change" and formats as percent and error checks for 0 values
                If tickerClose = 0 Or tickerOpen = 0 Then
                    Range("L" & summary_table_row).Value = "Invalid Data"
                ElseIf tickerClose <> 0 Or tickerOpen <> 0 Then
                    Range("L" & summary_table_row).Value = -(1 - (tickerClose / tickerOpen))
                    Range("L" & summary_table_row).NumberFormat = "0.00%"
                End If
            'Compare percent change to "greatest" values
                If tickerClose <> 0 Or tickerOpen <> 0 Then
                    If -(1 - (tickerClose / tickerOpen)) > gpiAmt Then
                        gpiTicker = ticker
                        gpiAmt = -(1 - (tickerClose / tickerOpen))
                    ElseIf -(1 - (tickerClose / tickerOpen)) < gpdAmt Then
                        gpdTicker = ticker
                        gpdAmt = -(1 - (tickerClose / tickerOpen))
                    End If
                End If
            'Imcrements summary table
            summary_table_row = summary_table_row + 1
            'resets ticker volume tracker
            tickerVol = 0
            'sets new value for next ticker open
            tickerOpen = Cells(i + 1, 3).Value
            Else
            'For all lines where the next line has the same ticker value, just add to total
            tickerVol = tickerVol + Cells(i, 7).Value
            End If
        Next i
    'display "greatest" values
    Range("O2") = gpiTicker
    Range("O3") = gpdTicker
    Range("O4") = gtvTicker
    Range("P2") = gpiAmt
    Range("P2").NumberFormat = "0.00%"
    Range("P3") = gpdAmt
    Range("P3").NumberFormat = "0.00%"
    Range("P4") = gtvAmt

End Sub