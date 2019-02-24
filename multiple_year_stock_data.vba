Sub total_volume()
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

''loop through each sheet
For Each ws In ThisWorkbook.Worksheets
    ws.Activate
    Dim ticker As New Collection
    Dim sum As New Collection
    Dim open_year As New Collection
    Dim close_year As New Collection
    Index = 0
    old_ticker = Q
    ''loop through cells in ticker column to determine if new or old ticker
    For I = 2 To Cells(Rows.Count, 1).End(xlUp).Row
        tick = Cells(I, 1).Value
        ''if old ticker, add volume as well as open and close to current index in these lists
        If tick = old_ticker Then
            new_volume = sum(Index) + Cells(I, 7).Value
            sum.Remove (Index)
            sum.Add new_volume
            new_open = open_year(Index) + Cells(I, 3).Value
            open_year.Remove (Index)
            open_year.Add new_open
            new_close = close_year(Index) + Cells(I, 6).Value
            close_year.Remove (Index)
            close_year.Add new_close
        ''if new ticker, move to next index in list of tickers and their corresponding total volumes, open and close year values
        Else
            Index = Index + 1
            ticker.Add tick
            sum.Add Cells(I, 7).Value
            open_year.Add Cells(I, 3).Value
            close_year.Add Cells(I, 6).Value
            old_ticker = tick
        End If
    Next I
    ''make new columns with total volume for each tickers, as well as change from open of year to close year both raw and percent change
    For I = 1 To ticker.Count
        Cells((I + 1), 9).Value = ticker(I)
        Cells((I + 1), 12).Value = sum(I)
        Cells((I + 1), 10).Value = close_year(I) - open_year(I)
        Percent_Chng = ((close_year(I) - open_year(I)) / open_year(I)) * 100
        Cells((I + 1), 11).Value = Percent_Chng
    Next I
    
    Range("I1") = "Ticker"
    Range("L1") = "Total Stock Volume"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    
    ''conditional formatting of positive change in green and negative change in red
    For I = 2 To ticker.Count
        If Cells(I, 10).Value >= 0 Then
        Cells(I, 10).Interior.ColorIndex = 4
        ElseIf Cells(I, 10).Value < 0 Then
        Cells(I, 10).Interior.ColorIndex = 3
        End If
    Next I
    
    Dim max As LongLong
    max = 0
    Dim max_increase As Double
    max_increase = 0
    Dim max_decrease As Double
    max_decrease = 0
    
    Dim max_ticker As String
    max_ticker = ""
    Dim max_increase_ticker As String
    max_increase_ticker = ""
    Dim max_decrease_ticker As String
    max_decrease_ticker = ""
    
    ''loop through values in new total volume and percent change columns to determine which values are the lowest and highest percent change, as well as highest total volume
    For I = 2 To ticker.Count
        If Cells(I, 12).Value > max Then
        max = Cells(I, 12).Value
        max_ticker = Cells(I, 9).Value
        End If
        If Cells(I, 11).Value > max_increase Then
        max_increase = Cells(I, 11).Value
        max_increase_ticker = Cells(I, 9).Value
        End If
        If Cells(I, 11).Value < max_decrease Then
        max_decrease = Cells(I, 11).Value
        max_decrease_ticker = Cells(I, 9).Value
        End If
    Next I
    
    ''print values in desired cells
    Range("O2") = "Greatest % Increase"
    Range("P2") = max_increase_ticker
    Range("O3") = "Greatest % Decrease"
    Range("P3") = max_decrease_ticker
    Range("O4") = "Greatest Total Volume"
    Range("P4") = max_ticker
    Range("P1") = "Ticker"
    Range("Q1") = "Value"
    Range("Q2") = max_increase
    Range("Q3") = max_decrease
    Range("Q4") = max
Next

starting_ws.Activate ''activate the worksheet that was originally active

End Sub
