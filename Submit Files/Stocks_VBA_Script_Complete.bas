Attribute VB_Name = "Module1"
Sub Ticker()
  For Each ws In Worksheets
    
    'Sets an initial variable for holding the Ticker ID
    Dim Ticker As String

    'Determine the Last Row
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row


    'Sets the Variable to hold the Yearly Change
    Dim Open_num As Double
    Dim Close_num As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double

    Dim Greatest_Percent_Increase As Double
    Dim Greatest_Percent_Decrease As Double
    Dim Greatest_Total_Value As Double


    Open_num = 1
    Close_num = 0
    Yearly_Change = 0
    Percent_Change = 0

    Greatest_Percent_Increase = 0
    Greatest_Percent_Decrease = 0
    Greatest_Total_Value = 0

    'Sets the variable to hold the total Stock Volume
    Dim Stock_Volume_Total As Double
    Stock_Volume_Total = 0


    'Keep track of the location for each tickers information in the summary page
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    'Stating the Header Values
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"



    'Loop through all tickers and info
    For i = 2 To LastRow
        'Sets the Opening number of the stock of the year once the previous ticker cells does not match
        If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Then

            Open_num = ws.Cells(i, 3).Value

        End If

        'Check if we are still within the same ticker, if it is not we collect the info
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Set the Ticker ID
            Ticker = ws.Cells(i, 1).Value

            'Set the Close number
            Close_num = ws.Cells(i, 6).Value

            'Add to the Stock Total Volume
            Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value

            'Calculate the Yearly Change
            Yearly_Change = Close_num - Open_num

            'Calculate the Percent change based on the yearly change
            Percent_Change = (Yearly_Change) / Open_num

            'Print the Ticker ID for the Summer Table
            ws.Range("I" & Summary_Table_Row).Value = Ticker

            'Print the Yearly Change for the Summer Table
            ws.Range("J" & Summary_Table_Row).Value = Yearly_Change

            'Print the Percent Change for the Summer Table
            ws.Range("K" & Summary_Table_Row).Value = FormatPercent(Percent_Change)

            'Print the Stock Volume total for the Summary Total
            ws.Range("L" & Summary_Table_Row).Value = Stock_Volume_Total

            'Add ont to the Summart Table Row
            Summary_Table_Row = Summary_Table_Row + 1

            'Resest the Stock Volume Total
            Stock_Volume_Total = 0

            'Reset the Open number
            Open_num = 1

            'Reset the Closing
            Close_num = 0

            'Reset the Yearly Change
            Yearly_Change = 0

            'Resest the Percent Change
            Percent_Change = 0

        'If the cell immediatly following a row is the same Ticker ID...
        Else

            'Add on to the Stock Volume Total
            Stock_Volume_Total = Stock_Volume_Total + ws.Cells(i, 7).Value

        End If


    Next i

    For i = 2 To LastRow

      ' Sets the conditions for the interior color of the cell
        If ws.Cells(i, 10).Value > 0 Then

            'Sets the Cell to the color green
            ws.Cells(i, 10).Interior.ColorIndex = 4

        Else

            'Sets the Cell to the color red
            ws.Cells(i, 10).Interior.ColorIndex = 3

        End If

    Next i

    For i = 2 To LastRow

        'Resets Summary to the top of the sheet
        Summary_Table_Row = 2

        'Use the highest Percent increase for the searching value
        If ws.Cells(i, 11).Value > Greatest_Percent_Increase Then
            Summary_Table_Row = 2

            Greatest_Percent_Increase = ws.Cells(i, 11).Value

            Ticker = ws.Cells(i, 9).Value

            ws.Range("P" & Summary_Table_Row).Value = Ticker

            ws.Range("Q" & Summary_Table_Row).Value = FormatPercent(Greatest_Percent_Increase)

        End If

        If ws.Cells(i, 11).Value < Greatest_Percent_Decrease Then
            Summary_Table_Row = 3

            Greatest_Percent_Decrease = ws.Cells(i, 11).Value

            Ticker = ws.Cells(i, 9).Value

            ws.Range("P" & Summary_Table_Row).Value = Ticker

            ws.Range("Q" & Summary_Table_Row).Value = FormatPercent(Greatest_Percent_Decrease)

        End If

        If ws.Cells(i, 12).Value > Greatest_Total_Value Then
            Summary_Table_Row = 4

            Greatest_Total_Value = ws.Cells(i, 12).Value

            Ticker = ws.Cells(i, 9).Value

            ws.Range("P" & Summary_Table_Row).Value = Ticker

            ws.Range("Q" & Summary_Table_Row).Value = Greatest_Total_Value

        End If
    Next i
  Next ws

End Sub

