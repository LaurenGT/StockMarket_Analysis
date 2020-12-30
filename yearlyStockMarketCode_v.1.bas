Attribute VB_Name = "Module2"
'functioning to add summary tab to each raw data tab
Sub yearlyStockData()

'reduce run time
Application.ScreenUpdating = False

'start stock analysis
'gathering unique ticker, yearly change in value, percent change in value and total stock volume

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

volume = 0
yearlyChange = 0

'ensure that the loop starts on the first raw data tab after creating summary sheet
Worksheets("2016").Select

'interate through all worksheets
For Each currentWS In Worksheets

'add summary table headers
currentWS.Range("I1") = "Ticker"
currentWS.Range("J1") = "Yearly Change"
currentWS.Range("K1") = "Percent Change"
currentWS.Range("L1") = "Total Stock Volume"

'identify ticker with greatest % increase and corresponding value
'identify ticker with greatest % decrease and corresponding value
'identify ticker with greatest total volumen and corresponding value
currentWS.Range("N2").Value = "Greatest % Increase"
currentWS.Range("N3").Value = "Greatest % decrease"
currentWS.Range("N4").Value = "Greatest Total Volume"
currentWS.Range("O1").Value = "Associated Ticker"
currentWS.Range("P1").Value = "Associated Value"

Debug.Print (currentWS.Name)

summaryTableRow = 2
openPrice = 2
lastRow = currentWS.Cells(Rows.Count, 1).End(xlUp).Row
'    MsgBox (lastRow) last row=70926 on tab A

'iterative loop to find cummulative volume
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
                'yearly change>0, green
                'yearly change<0 red
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

Worksheets("2016").Select

For Each currentWS In Worksheets

lastSummaryRow = currentWS.Cells(Rows.Count, 9).End(xlUp).Row

'identify greatest % increase, dcrease and greatest volume with max function
greatestIncrease = WorksheetFunction.Max(currentWS.Range("K2:K" & lastSummaryRow))
greatestDecrease = WorksheetFunction.Min(currentWS.Range("K2:K" & lastSummaryRow))
greatestVolume = WorksheetFunction.Max(currentWS.Range("L2:L" & lastSummaryRow))

    For i = 2 To lastSummaryRow

        'insert identified greatest % increase into summary table
        If currentWS.Cells(i, 11).Value = greatestIncrease Then
            currentWS.Range("O2").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P2").Value = greatestIncrease
        End If
        
        'insert identified greatest % decrease into summary table
        If currentWS.Cells(i, 11).Value = greatestDecrease Then
            currentWS.Range("O3").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P3").Value = greatestDecrease
        End If
        
        'insert greatest volume into summary table
        If currentWS.Cells(i, 12).Value = greatestVolume Then
            currentWS.Range("O4").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P4").Value = greatestVolume
        End If
    
    Next i
    'format cells for readability
    currentWS.Range("I:P").Columns.AutoFit
    currentWS.Columns("K").NumberFormat = "0.00%"
    currentWS.Range("P2:P3").NumberFormat = "0.00%"
        
Next 'iterate through next worksheet

Debug.Print ("complete")
End Sub
'greatest % increase, decrease and volume of the yearly data
'testing
Sub yearlyDataChallenge()

'assign variables
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

For Each currentWS In Worksheets

lastSummaryRow = currentWS.Cells(Rows.Count, 9).End(xlUp).Row

'identify ticker with greatest % increase and corresponding value
'identify ticker with greatest % decrease and corresponding value
'identify ticker with greatest total volumen and corresponding value
currentWS.Range("N2").Value = "Greatest % Increase"
currentWS.Range("N3").Value = "Greatest % decrease"
currentWS.Range("N4").Value = "Greatest Total Volume"
currentWS.Range("O1").Value = "Associated Ticker"
currentWS.Range("P1").Value = "Associated Value"

'identify greatest % increase, dcrease and greatest volume with max function
greatestIncrease = WorksheetFunction.Max(currentWS.Range("K2:K" & lastSummaryRow))
greatestDecrease = WorksheetFunction.Min(currentWS.Range("K2:K" & lastSummaryRow))
greatestVolume = WorksheetFunction.Max(currentWS.Range("L2:L" & lastSummaryRow))

    For i = 2 To lastSummaryRow

        'insert identified greatest % increase into summary table
        If currentWS.Cells(i, 11).Value = greatestIncrease Then
            currentWS.Range("O2").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P2").Value = greatestIncrease
        End If
        
        'insert identified greatest % decrease into summary table
        If currentWS.Cells(i, 11).Value = greatestDecrease Then
            currentWS.Range("O3").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P3").Value = greatestDecrease
        End If
        
        'insert greatest volume into summary table
        If currentWS.Cells(i, 12).Value = greatestVolume Then
            currentWS.Range("O4").Value = currentWS.Cells(i, 9).Value
            currentWS.Range("P4").Value = greatestVolume
        End If
    
    Next i
    'format cells for readability
    currentWS.Range("I:P").Columns.AutoFit
    currentWS.Columns("K").NumberFormat = "0.00%"
    currentWS.Range("P2:P3").NumberFormat = "0.00%"
        
Next 'iterate through next worksheet

End Sub


Sub challenge()
Worksheets("summaryTab").Select

'identify ticker with greatest % increase and corresponding value
'identify ticker with greatest % decrease and corresponding value
'identify ticker with greatest total volumen and corresponding value
Worksheets("summaryTab").Range("F2").Value = "Greatest % Increase"
Worksheets("summaryTab").Range("F3").Value = "Greatest % decrease"
Worksheets("summaryTab").Range("F4").Value = "Greatest Total Volume"
Worksheets("summaryTab").Range("G1").Value = "Associated Ticker"
Worksheets("summaryTab").Range("H1").Value = "Associated Value"

Dim greatestIncrease As Double
Dim greatestDecrease As Double
Dim greatestVolume As LongLong

'identify greatest % increase, dcrease and greatest volume with max function
greatestIncrease = WorksheetFunction.Max(Range("C2:C" & lastRow))
greatestDecrease = WorksheetFunction.Min(Range("C2:C" & lastRow))
greatestVolume = WorksheetFunction.Max(Range("D2:D" & lastRow))

For i = 2 To lastRow
    'insert identified greatest % increase into summary table
    If Cells(i, 3).Value = greatestIncrease Then
        Worksheets("summaryTab").Range("H2").Value = greatestIncrease
        Worksheets("summaryTab").Range("G2").Value = Cells(i, 3).Offset(, -2).Value
    End If
    
    'insert identified greatest % decrease into summary table
    If Cells(i, 3).Value = greatestDecrease Then
        Worksheets("summaryTab").Range("H3").Value = greatestDecrease
        Worksheets("summaryTab").Range("G3").Value = Cells(i, 3).Offset(, -2).Value
    End If
    
    'insert greatest volume into summary table
    If Cells(i, 4).Value = greatestVolume Then
        Worksheets("summaryTab").Range("H4").Value = greatestVolume
        Worksheets("summaryTab").Range("G4").Value = Cells(i, 4).Offset(, -3).Value
    End If
    
Next i

'format cells to be more readable

Worksheets("summaryTab").Columns("A:H").AutoFit
Worksheets("summaryTab").Rows(1).Font.Bold = True
Worksheets("summaryTab").Range("F2:F4").Font.Bold = True
Worksheets("summaryTab").Range("H2:H3").NumberFormat = "0.00%"
Worksheets("summaryTab").Range("C2:c" & lastRow).NumberFormat = "0.00%"

MsgBox ("Analysis Complete. See summary tab for results.")


End Sub
