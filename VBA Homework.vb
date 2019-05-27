Sub Homework_Code_ALL_SHEETS()

Dim ws As Worksheet
For Each ws In ThisWorkbook.Worksheets

'Name Columns and Other Desciptors

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
'Making the sheet look nice
    ws.Columns(10).AutoFit
    ws.Columns(11).AutoFit
    ws.Columns(12).AutoFit
    ws.Columns(15).AutoFit
    ws.Columns(17).AutoFit
    
'Define Variables
    
    Dim i As Long
    Dim last_row As Long
    
    'Easy
    Dim TotalStockVolume As Double
    Dim Volume As Double
    Dim Ticker As String
    Dim NextTicker As String
    Dim TickerCounter As Long
    
    'Moderate
    Dim PrevTicker As String
    Dim YearStartOpenPrice As Double
    Dim YearEndClosePrice As Double
    Dim OpenPrice As Double
    Dim ClosePrice As Double
    Dim YearlyChange As Double
    Dim PercentChange As Double



    'Scaled numbers for loops
    last_row_1 = Cells(ws.Rows.Count, 1).End(xlUp).Row
    last_row_11 = Cells(ws.Rows.Count, 11).End(xlUp).Row
    TickerCounter = 2
    
    
For i = 2 To last_row_1
    'Assigning Values
    Volume = ws.Cells(i, 7).Value
    Ticker = ws.Cells(i, 1).Value
    NextTicker = ws.Cells(i + 1, 1).Value
    PrevTicker = ws.Cells(i - 1, 1).Value
    OpenPrice = ws.Cells(i, 3).Value
    ClosePrice = ws.Cells(i, 6).Value
    
    TotalStockVolume = TotalStockVolume + Volume
    
    'Loop to find Yearly Change
    If Ticker <> PrevTicker Then
        YearStartOpenPrice = OpenPrice
    End If
       
    'Determine if next cell is different
    If Ticker <> NextTicker Then
        ws.Cells(TickerCounter, 9).Value = Ticker
        ws.Cells(TickerCounter, 12).Value = TotalStockVolume
        ws.Cells(TickerCounter, 12).NumberFormat = "#,##0_);(#,##0)"
        TotalStockVolume = 0
        
        'Yearly Change and Percent Change Code
        YearEndClosePrice = ClosePrice
        YearlyChange = YearEndClosePrice - YearStartOpenPrice
        
        If YearStartOpenPrice > 0 Then
            PercentChange = (YearEndClosePrice / YearStartOpenPrice) - 1
        End If
        
        If YearStartOpenPrice <= 0 Then
        PercentChange = 0
        End If
        
        ws.Cells(TickerCounter, 10).Value = YearlyChange
        ws.Cells(TickerCounter, 10).NumberFormat = "General"
        ws.Cells(TickerCounter, 11).Value = PercentChange
        ws.Cells(TickerCounter, 11).NumberFormat = "0.00%"
            If ws.Cells(TickerCounter, 11).Value > 0 Then
                ws.Cells(TickerCounter, 11).Interior.ColorIndex = 4
            End If
            If ws.Cells(TickerCounter, 11).Value < 0 Then
                ws.Cells(TickerCounter, 11).Interior.ColorIndex = 3
            End If
        
        'Reseting Variables
        YearStartOpenPrice = 0
        YearEndClosePrice = 0
        YearlyChange = 0
        PercentChange = 0
        
        'Adding to TickerCounter
        TickerCounter = TickerCounter + 1
    End If
Next i
    
'Loop #2 for Greatest Increase, Decrease, and Volume

'Define Variables
    Dim GreatIncValue As Double
    Dim GreatDecValue As Double
    Dim GreatVolValue As Double
    Dim GreatIncTicker As String
    Dim GreatDecTicker As String
    Dim GreatVolTicker As String
    Dim PercentChangeValue As Double
    Dim TickerValue As String
    
    GreatIncValue = 0
    GreatDecValue = 0
    GreatVolValue = 0

For i = 2 To last_row_11
   
    If ws.Cells(i, 11).Value >= GreatIncValue Then
        GreatIncValue = ws.Cells(i, 11).Value
        GreatIncTicker = ws.Cells(i, 9).Value
    End If
    
    If ws.Cells(i, 11).Value <= GreatDecValue Then
        GreatDecValue = ws.Cells(i, 11).Value
        GreatDecTicker = ws.Cells(i, 9).Value
    End If
    
    If ws.Cells(i, 12).Value >= GreatVolValue Then
        GreatVolValue = ws.Cells(i, 12).Value
        GreatVolTicker = ws.Cells(i, 9).Value
    End If
    
Next i

ws.Cells(2, 16).Value = GreatIncTicker
ws.Cells(2, 17).Value = GreatIncValue
    ws.Cells(2, 17).NumberFormat = "0.00%"
ws.Cells(3, 16).Value = GreatDecTicker
ws.Cells(3, 17).Value = GreatDecValue
    ws.Cells(3, 17).NumberFormat = "0.00%"
ws.Cells(4, 16).Value = GreatVolTicker
ws.Cells(4, 17).Value = GreatVolValue
    ws.Cells(4, 17).NumberFormat = "#,##0_);(#,##0)"
    
Next ws

End Sub
