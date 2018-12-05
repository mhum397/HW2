Sub easy_solution()

    ' Set an initial variable for holding the ticker name
    Dim Ticker_Name As String

    ' Set an initial variable for holding the total volume per ticker symbol
    Dim Ticker_Total As Double
    Ticker_Total = 0

    ' Keep track of the location for each ticket symbol and volume in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    Dim ws As Worksheet
    Dim Lastrow As Long
    Set ws = ActiveSheet
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through ticker symbols
    For i = 2 To Lastrow

        ' Check if we are still within the same ticker symbol, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

        ' Set the Ticker symbol
            Ticker_Name = Cells(i, 1).Value

        ' Add to the volume total
            Ticker_Total = Ticker_Total + Cells(i, 7).Value

        ' Print the Ticker Symbol and Volume in the Summary Table
            Range("K1").Value = "Ticker Symbol"
            Range("L1").Value = "Volume"
      
            Range("K" & Summary_Table_Row).Value = Ticker_Name

        ' Print the Volume Amount to the Summary Table
            Range("L" & Summary_Table_Row).Value = Ticker_Total

        ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
        ' Reset the volume Total
            Ticker_Total = 0

        ' If the cell immediately following a row is the same brand...
        Else

        ' Add to the volume Total
            Ticker_Total = Ticker_Total + Cells(i, 7).Value

        End If

    Next i

End Sub

