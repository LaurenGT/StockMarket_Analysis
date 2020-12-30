Attribute VB_Name = "FinalAnalysisCode"
Sub yearlyStockDataAnalysis()

'reduce run time
Application.ScreenUpdating = False

'assign variable
Dim ticker As String
Dim volume As Double
Dim cummulativeVolume As Integer
Dim lastRow As Long
Dim currentWS As Worksheet
Dim summaryTableRow As Integer
Dim openPrice As Double
Dim closePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double

'assign varibales to greatest % increase, dcrease and greatest volume
Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As LongLong
Dim lastSummaryRow As Long

'ensure that the loop starts on the first raw data tab after creating summary sheet
Worksheets("2016").Select

'interate through all worksheets
For Each currentWS In Worksheets

'add summary table headers to each data tab
currentWS.Range("I1") = "Ticker"
currentWS.Range("J1") = "Yearly Change"
currentWS.Range("K1") = "Percent Change"
currentWS.Range("L1") = "Total Stock Volume"

currentWS.Range("N2").Value = "Greatest % Increase"
currentWS.Range("N3").Value = "Greatest % decrease"
currentWS.Range("N4").Value = "Greatest Total Volume"
currentWS.Range("O1").Value = "Associated Ticker"
currentWS.Range("P1").Value = "Associated Value"

volume = 0
yearlyChange = 0
summaryTableRow = 2
openPrice = 2

'identify last row of source data on each tabs
lastRow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row

'iterative loop to find cummulative volume, yearly change and percent change
    For i = 2 To lastRow
        If currentWS.Cells(i + 1, 1).Value <> currentWS.Cells(i, 1).Value Then
            ticker = currentWS.Cells(i, 1).Value 'identify change in tickers
            volume = volume + currentWS.Cells(i, 7).Value 'track cummulative volume as iterations progress down rows
            currentWS.Range("I" & summaryTableRow).Value = ticker 'output for unique ticker value location
            currentWS.Range("L" & summaryTableRow).Value = volume 'output for cummulative volume location
            initialOpen = currentWS.Cells(openPrice, 3).Value 'identify open price at start of new ticker
            closePrice = currentWS.Cells(i, 6).Value 'identify close price at end of current ticker
            yearlyChange = closePrice - initialOpen 'calculate difference in value
            currentWS.Range("J" & summaryTableRow).Value = yearlyChange 'output for yearlyChange location
                If initialOpen = "0" Then
                    currentWS.Range("K" & summaryTableRow).Value = "Undefined, open price=0"
                Else
                    percentChange = (yearlyChange / initialOpen) 'calculates percent change over year
                    currentWS.Range("K" & summaryTableRow).Value = percentChange 'output for percent change location
                End If
                
                'after each value is entered on the summary table, conditionally format
                'yearly change increase, green
                'yearly change decrease, red
                'anything else, white
                If yearlyChange > 0 Then
                    currentWS.Range("J" & summaryTableRow).Interior.Color = RGB(0, 255, 0)
                ElseIf yearlyChange < 0 Then
                    currentWS.Range("J" & summaryTableRow).Interior.Color = RGB(255, 0, 0)
                Else
                    currentWS.Range("J" & summaryTableRow).Interior.Color = RGB(255, 255, 255)
                End If
            
            summaryTableRow = summaryTableRow + 1 'acknowledges where next loop should start in summary table
            volume = 0 'reset the cummulative volume calculation
            openPrice = i + 1 'resets the identifier for opening price
        Else
            volume = volume + currentWS.Cells(i, 7).Value 'cummulative volume while ticker is consistent
        End If

    Next i 'restart loop with new i values incremented by one
Next 'iterate through next worksheet

'loop through summary tables on each tab
'populate the greatest %increase, %decrease and volume
Worksheets("2016").Select

For Each currentWS In Worksheets

'identify last row of summary tables on each tab
lastSummaryRow = currentWS.Cells(Rows.Count, 9).End(xlUp).Row

'identify greatest % increase, dcrease and greatest volume with max/min functions
greatestIncrease = WorksheetFunction.Max(currentWS.Range("K2:K" & lastSummaryRow))
greatestDecrease = WorksheetFunction.Min(currentWS.Range("K2:K" & lastSummaryRow))
greatestVolume = WorksheetFunction.Max(currentWS.Range("L2:L" & lastSummaryRow))
    
    'iterate through all rows in summary table to find values and tickers associated with above calculations
    For i = 2 To lastSummaryRow

        'insert identified greatest % increase and associated ticker into summary table
        If currentWS.Cells(i, 11).Value = greatestIncrease Then
            currentWS.Range("O2").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P2").Value = greatestIncrease
        End If
        
        'insert identified greatest % decrease and associated ticker into summary table
        If currentWS.Cells(i, 11).Value = greatestDecrease Then
            currentWS.Range("O3").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P3").Value = greatestDecrease
        End If
        
        'insert greatest volume and associated ticker into summary table
        If currentWS.Cells(i, 12).Value = greatestVolume Then
            currentWS.Range("O4").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P4").Value = greatestVolume
        End If
    
    Next i
    
    'format all new tables for readability
    currentWS.Range("I:P").Columns.AutoFit
    currentWS.Columns("K").NumberFormat = "0.00%"
    currentWS.Range("P2:P3").NumberFormat = "0.00%"
        
Next 'iterate through next worksheet

'paste all final summary tables on a new summary tab for better comparison of years
Worksheets("2016").Select

'add new summary tab
Sheets.Add(before:=Sheets(1)).Name = "summaryTab"

'assign summary tab as variable
Dim summaryTab As Worksheet

'add identifying headers and formatting for each year
Worksheets("summaryTab").Range("A1").Value = "2016 Stock Analysis"
Worksheets("summaryTab").Range("J1").Value = "2015 Stock Analysis"
Worksheets("summaryTab").Range("S1").Value = "2014 Stock Analysis"
Worksheets("summaryTab").Range("A1:H1, J1:Q1, S1:Z1").Merge
Worksheets("summaryTab").Range("A1:H1, J1:Q1, S1:Z1").HorizontalAlignment = xlCenter
Worksheets("summaryTab").Range("A1:H1, J1:Q1, S1:Z1").Interior.Color = RGB(224, 224, 224)

'copy and paste summary tables from source data tabs onto new summary tab

Worksheets("2016").Range("I1:P3169").Copy Destination:=Worksheets("summaryTab").Range("A2")
Worksheets("2015").Range("I1:P3005").Copy Destination:=Worksheets("summaryTab").Range("J2")
Worksheets("2014").Range("I1:P2836").Copy Destination:=Worksheets("summaryTab").Range("S2")

'format all column widths to see full content of all cells
Worksheets("summaryTab").Range("A:Z").Columns.AutoFit

MsgBox ("Analysis Complete. See summaryTab for all years or look through yearly tabs for individual summary tables.")

End Sub
