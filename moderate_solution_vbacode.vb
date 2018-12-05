Sub moderate_solution()

    ' Set an initial variable for holding the ticker name
    Dim Ticker_Name As String

    ' Set an initial variable for holding the total volume per ticker symbol
    Dim Ticker_Total As Double
    Ticker_Total = 0
    
    ' Set an initial variables to complete yearly change calculation
    Dim Open_Price As Double
    Dim Close_Price As Double
    Dim Change_Price As Double
    
    ' Set initial variables to complete the percent change calculation
    Dim Percent_Change As Double

    ' Keep track of the location for each ticker symbol and other data in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    
    ' Create the worksheet variable to find the last row of each worksheet
    Dim ws As Worksheet
    Dim Lastrow As Long
    Set ws = ActiveSheet
    Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Loop through ticker symbols
    For i = 2 To Lastrow

        ' Check if we are still within the same ticket symbol, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            ' Set the Ticker symbol
            Ticker_Name = Cells(i, 1).Value

            ' Add to the volume total
            Ticker_Total = Ticker_Total + Cells(i, 7).Value
            
            ' Store the Close Price
            Close_Price = Cells(i, 6).Value
            
            ' Calculate the Change between the open and close price
            Change_Price = Close_Price - Open_Price
            
            ' Do not allow for division of 0 if the open price is 0
            If Open_Price > 0 Then
            
                Percent_Change = (Close_Price - Open_Price) / Open_Price
            
            Else
            
                Percent_Change = 1
            
            End If
            
            ' Print the Ticker Symbol and Volume in the Summary Table
            Columns("J").ColumnWidth = 13
            Columns("K").ColumnWidth = 13
            Columns("L").ColumnWidth = 13
            Columns("M").ColumnWidth = 16
            Range("J1").Value = "Ticker Symbol"
            Range("K1").Value = "Yearly Change"
            Range("L1").Value = "Percent Change"
            Range("M1").Value = "Total Stock Volume"
            
            ' Print the ticker name to the Summary Table
            Range("J" & Summary_Table_Row).Value = Ticker_Name
            
            ' Print the yearly change to the Summary Table
            Range("K" & Summary_Table_Row).Value = Change_Price
            
            ' Print the percent change to the Summary Table and format to percentage
            Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
            Range("L" & Summary_Table_Row).Value = Percent_Change
            
            ' Print the volume to the Summary Table
            Range("M" & Summary_Table_Row).Value = Ticker_Total

            ' Format the color of yearly change based on positive or negative change
            If Change_Price >= 0 Then
                
                Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            
            Else
            
                Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            
            End If

            ' Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
      
            ' Reset the volume Total
            Ticker_Total = 0
            
            ' Reset the yearly change
            Change_Price = 0
            
            ' Reset the percentage change
            Percent_Change = 0
            

        ' If the cell immediately following a row is the same ticker
        Else

            ' Add to the volume Total
            Ticker_Total = Ticker_Total + Cells(i, 7).Value
            
            ' Gather the Open Price for the stock
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
            
                Open_Price = Cells(i, 3).Value
            
            End If
        
            
        End If

    Next i

End Sub

