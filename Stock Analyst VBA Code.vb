
Sub stock_analyst()
    'Create a script that will loop through all the stocks for one year and output the following information.
    
        '* The ticker symbol.
        Dim ticker As String
    
        '* YEARLY CHANGE from OPENING PRICE at the beginning of a given year to the CLOSING PRICE at the end of the year.
        Dim yearlyChange As Double
        Dim openPrice As Double
        'First opening price to appear in the sheet is cell C3, so
        openPrice = Cells(2, 3).Value
        Dim closePrice As Double
    
        '*The PERCENT CHANGE from opening price at the beginning of a given year to the closing price at the end of that year.
        Dim percentChange As Double
    
        'The total stock VOLUME of the stock.
        Dim totalVolume As Double
        totalVolume = 0

        'A row tracker for loops, omits row 1.
        Dim tickerRowCounter As Integer
        tickerRowCounter = 2
    
    'The "last row" code for finding the end of a sheet
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'This will loop through rows to gather ticker data, and calculate percent/yearly changes
    For i = 2 To lastRow
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticker = Cells(i, 1).Value
            totalVolume = totalVolume + Cells(i, 7).Value
            closePrice = Cells(i, 6).Value
            yearlyChange = closePrice - openPrice
            'To avoid errors in percent change when dividing by zero:
            If openPrice = 0 Then
                percentChange = 0
            Else
                percentChange = yearlyChange / openPrice
            End If
            'Print items to summary table
            Range("I" & tickerRowCounter).Value = ticker
            Range("J" & tickerRowCounter).Value = yearlyChange
            Range("K" & tickerRowCounter).Value = percentChange
            Range("L" & tickerRowCounter).Value = totalVolume
            'Reset row counter, volume and opening price
            tickerRowCounter = tickerRowCounter + 1
            totalVolume = 0
            openPrice = Cells(i + 1, 3)
        Else
            totalVolume = totalVolume + Cells(i, 7).Value
        End If
    Next i
    '*You should also have CONDITIONAL FORMATTING that will highlight POSITIVE change in GREEN and NEGATIVE change in RED.
    summaryLastRow = Cells(Rows.Count, 9).End(xlUp).Row
    For i = 2 To summaryLastRow
        If Cells(i, 10).Value > 0 Then
            Cells(i, 10).Interior.Color = vbGreen
        Else
            Cells(i, 10).Interior.Color = vbRed
        End If
    Next i
    
    'Summary table headers:
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Stock Volume"
End Sub