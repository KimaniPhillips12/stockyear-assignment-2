Attribute VB_Name = "Module1"
Sub ticker()
    ' Set dimensions
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim rowCount As Long
    ' Set title row
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("Q4").Value = "Greatest Total Volume"
    Range("Q3").Value = "Greatest % Decrease"
    Range("Q2").Value = "Greatest % Increase"
    ' Set initial values
    increase = 0
    decrease = 0
    Volume = 0
    increase_ticker = ""
    decrease_ticker = ""
    ticker_volume = ""
    j = 2
    total = 0
    change = 0
    Start = 2
    ' get the row number of the last row with data
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    For i = 2 To rowCount
        ' If ticker changes then print results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ' Handle zero total volume
            If total = 0 Then
                ' Print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
            Else
                ' Find first non-zero starting value
                If Cells(Start, 3) = 0 Then
                    For Find_value = Start To i
                        If Cells(Find_value, 3).Value <> 0 Then
                            Start = Find_value
                            Exit For
                        End If
                    Next Find_value
                End If
                ' Calculate yearly change and percent change
                change = Cells(i, 6).Value - Cells(Start, 6).Value
                porc_change = change / Cells(Start, 6).Value
                ' Start of the next stock ticker
                Start = i + 1
                ' Print the results
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("K" & 2 + j).Value = porc_change
                End If
            ' Reset variables for a new stock ticker
            total = 0
            change = 0
            j = j + 1
        Else
            ' Accumulate total volume
            total = total + Cells(i, 7).Value
        End If
    Next i
    ' Reset variables for the last stock ticker
    Range("P2").Value = increase_ticker
    Range("P3").Value = decrease_ticker
    Range("P4").Value = ticker_volume
    Cells(4, 17).Value = Volume
    Cells(3, 17).Value = porc_decrease
    Cells(2, 17).Value = porc_increase
End Sub
