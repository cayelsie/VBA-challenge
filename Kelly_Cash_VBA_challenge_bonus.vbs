Sub stocks()
'Declare variables for ticker, opening & closing prices, yearly change, percent change & total stock volume
Dim Ticker As String
Dim Opening As Double
Dim Closing As Double
Dim Yearly As Double
Dim Percent As Double
Dim Total As Double

'Declare variable for keeping track of each ticker by row for the summary table
Dim SummaryRow As Integer

'Declare the last row to be able to use it in the loop
Dim LastRow As Long
LastRow = Cells(Rows.Count, 1).End(xlUp).Row

'Declare a variable to hold position of the first row of each ticker for obtaining the year's opening value
Dim Firstday As Long

'Declare a variable for the active worksheet
Dim ws As Worksheet

'Create a loop through the worksheets
For Each ws In Worksheets

'Puts heading text in each worksheet
ws.Range("J1").Value = "Ticker"
ws.Range("K1").Value = "Yearly Change"
ws.Range("L1").Value = "Percent Change"
ws.Range("M1").Value = "Total Stock Volume"

'Start Total stock volume value at zero
Total = 0

'Start opening date row at row 2 to correspond with first ticker
Firstday = 2

'Start summary table at row 2
SummaryRow = 2

    'Create a loop through the stock data
    For i = 2 To LastRow

        'Declare variable and assign a value to the stock volume to ensure it is set to Double throughout the calculations (big numbers!)
        Dim Stockvolume As Double
        Stockvolume = ws.Range("G" & i).Value


        'Check to see if still looping through the same ticker display by comparing each cell in the first column to the next cell down. If the ticker in the next cell has changed, then:
        If ws.Range("A" & i).Value <> ws.Range("A" & (i + 1)).Value Then


            'Assign a value to Ticker
            Ticker = ws.Range("A" & i).Value
    
            'Assign a value to the open price, starting with the first row of data for the first ticker
            Opening = ws.Range("C" & Firstday).Value

            'Pull the last closing value for the new ticker
            Closing = ws.Range("F" & i).Value
    
            'Yearly change calculation
            Yearly = Closing - Opening
    
            'Account for opening values of zero
            If Opening = 0 Then
            Percent = 0
    
           'Percent change calculation occurs as long as opening value is above zero. Keep it at two decimal places.
            Else
            Percent = Round(Yearly / Opening * 100, 2)
    
            End If

            'Ensure that the last stock volume is summed with the total for the ticker data being calculated
            Total = Total + Stockvolume

            'Declare the ticker symbol in the summary table
            ws.Range("J" & SummaryRow).Value = Ticker

            'Declare the total stock volume in the summary table
            ws.Range("M" & SummaryRow).Value = Total
    
            'Declare Yearly change in summary table
             ws.Range("K" & SummaryRow).Value = Yearly
    
            'Declare Percent change in summary table and add the percent symbol
            ws.Range("L" & SummaryRow).Value = "%" & Percent
    
            'If the yearly change is positive then turn the cell green
             If Yearly > 0 Then
             ws.Range("K" & SummaryRow).Interior.ColorIndex = 4
        
            'If the yearly change is not positive then turns the cell red
            Else
            ws.Range("K" & SummaryRow).Interior.ColorIndex = 3
    
            End If
        
            'Ensure that the total gets set back to zero for the new ticker
            Total = 0

            'Ensure that the summary data for the next ticker will be deposited into the next row
             SummaryRow = SummaryRow + 1
    
             'Ensure that the row placeholder for the opening day is starting at the first row of the new ticker
             Firstday = i + 1

        'If the data is still running through the same ticker, ensure that the stock volume is being summed
         Else
        Total = Total + Stockvolume
        
        End If
    Next i
    
    'Declare variables for second summary table
    Dim MaxPercent As Double
    Dim MinPercent As Double
    Dim MaxVolume As Double
    Dim MaxPercentTicker As String
    Dim MaxTotalTicker As String
    Dim MinPercentTicker As String
   
    'Display headings for second summary table in each worksheet
    ws.Range("P2").Value = "Greatest % Increase"
    ws.Range("P3").Value = "Greatest % Decrease"
    ws.Range("P4").Value = "Greatest Total Volume"
    ws.Range("Q1").Value = "Ticker"
    ws.Range("R1").Value = "Value"

    'Start each numeric variable at zero
    MaxPercent = 0
    MinPercent = 0
    MaxVolume = 0
    
        'Loop through the percentage data in the first summary table.
        For x = 2 To SummaryRow
     
            'If the percent in a cell is greater than the last value, hold it as the MaxPercent and hold the corresponding ticker
            If ws.Range("L" & x).Value > MaxPercent Then
            MaxPercent = ws.Range("L" & x).Value
            MaxPercentTicker = ws.Range("J" & x).Value
        
                'Display the values in the second summary table in each worksheet
                ws.Range("R2").Value = Round(MaxPercent * 100, 2) & "%"
                ws.Range("Q2").Value = MaxPercentTicker

            End If
        Next x

        'Loop through the percentage data in the first summary table.
        For y = 2 To SummaryRow
        
            'If the percent in a cell is less than the last value, hold it as the MinPercent and hold the corresponding ticker
            If ws.Range("L" & y).Value < MinPercent Then
            MinPercent = ws.Range("L" & y).Value
            MinPercentTicker = ws.Range("J" & y).Value
        
                'Display the values in the second summary table in each worksheet
                ws.Range("R3").Value = Round(MinPercent * 100, 2) & "%"
                ws.Range("Q3").Value = MinPercentTicker
            
            End If
         Next y
    
        'Loop through the stock volume data in the first summary table.
        For Z = 2 To SummaryRow
   
            'If the percent in a cell is greater than the last value, hold it as the MaxVolume and hold the corresponding ticker
            If ws.Range("M" & Z).Value > MaxVolume Then
            MaxVolume = ws.Range("M" & Z).Value
            MaxTotalTicker = ws.Range("J" & Z).Value

                'Display the values in the second summary table in each worksheet
                ws.Range("R4").Value = MaxVolume
                ws.Range("Q4").Value = MaxTotalTicker
    
            End If
         Next Z
 
    Next ws
End Sub