# VBA-Challenge

Sub StockAnalysis()

Dim Ticker As String
Dim YearlyChange As Double
Dim TotalVolume As Double
Dim SummaryTable_RowCounter As Integer

'set up ws for loop
Dim ws As Worksheet
For Each ws In Worksheets

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    SummaryTable_RowCounter = 2
    YearOpen = ws.Cells(2, 3).Value
    
    For r = 2 To LastRow
    
        'set up conditional
        If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
            Ticker = ws.Cells(r, 1).Value
            ws.Cells(SummaryTable_RowCounter, 9).Value = Ticker
            ' calculate and pull yearly change from opening price to closing price
            YearClose = ws.Cells(r, 6).Value
            YearlyChange = YearClose - YearOpen
            ws.Cells(SummaryTable_RowCounter, 10) = YearlyChange
            'Calculate and pull percentage change from opening price to closing price
            PercentChange = (YearClose - YearOpen) / YearClose
            ws.Cells(SummaryTable_RowCounter, 11).Value = PercentChange
            'calculate and pull stock volume for each ticker
            volume = volume + ws.Cells(r, 7).Value
            ws.Cells(SummaryTable_RowCounter, 12) = volume
        
            SummaryTable_RowCounter = SummaryTable_RowCounter + 1
            volume = 0
            YearOpen = ws.Cells(r + 1, 3).Value
            
        Else
            volume = volume + ws.Cells(r, 7).Value
        End If
        
       
        ' format positive and negative cells
        If ws.Cells(r, 10).Value > "0.00" Then
            ws.Cells(r, 10).Interior.Color = vbGreen
        ElseIf ws.Cells(r, 10).Value < "0.00" Then
            ws.Cells(r, 10).Interior.Color = vbRed
        ElseIf ws.Cells(r, 10).Value = "0.00" Then
           ws.Cells(r, 10).Interior.Color = vbWhite
        End If
        
        ws.Cells(r, 11).NumberFormat = "0.00%"
        
    Next r
    

    
Next ws

End Sub




