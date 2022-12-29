Attribute VB_Name = "Module1"
'Create a script that loops through all the stocks for one year and outputs the following information:
 ' * The ticker symbol.
  '* Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
  '* The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
  '* The total stock volume of the stock.
'**Note:** Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

'go through all sheets
'for each sheet
'set variables up
'for each row
'if ticker is sameas before
'add properly
'if not
'new cell
'do some calc


Sub doStuffToASheet(wSheet As Worksheet)
    
    Dim tickerSymbol As String
    Dim stockYear As String
    
    Dim yearlyChange As Double
    Dim percentYearlyChange As Double
    Dim yearStart As Double
    Dim yearEnd As Double
    
    Dim curTotalStockVolume As Double
    Dim totalNumberOfRows As Long
    
    Dim targetYear As Integer
    Dim stockIndex As Integer
    Dim GreatIncStockIndex As Integer
    Dim GreatDecStockIndex As Integer
    Dim GreatVolStockIndex As Integer
    

    totalNumberOfRows = WorksheetFunction.CountA(Columns("A:A"))
    targetYear = 2020
    stockIndex = 1
    tickerSymbol = wSheet.Cells(2, 1).Value
    yearStart = wSheet.Cells(2, 3).Value
    
    greatestIncreaseStockIndex = 1
    greatestDecreaseStockIndex = 1
    greatestVolumeStockIndex = 1
    
    Dim i As Integer
    i = 2
    For i = 2 To totalNumberOfRows
        stockYear = Left(Str(wSheet.Cells(i, 2)), 5)
        stockYear = Right(stockYear, 4)
        If (CInt(stockYear) = targetYear) Then
            If (tickerSymbol = wSheet.Cells(i, 1).Value) Then
                curTotalStockVolume = curTotalStockVolume + wSheet.Cells(i, 7).Value
            Else
                'calculate stuff
                yearEnd = wSheet.Cells(i - 1, 6).Value
                yearlyChange = yearStart - yearEnd
                percentYearlyChange = yearStart / yearEnd * 100
                totalStockVolume = curTotalStockVolume
                'update the stock list
                Call updateStockList(wSheet, stockIndex, tickerSymbol, yearlyChange, percentYearlyChange, totalStockVolume)
                'conditional highlighting
                Call conditionalHighlighting(wSheet, yearlyChange, stockIndex)
                'establish greatest
                If (percentYearlyChange > wSheet.Cells(greatestIncreaseStockIndex, 11).Value) Then
                    greatestIncreaseStockIndex = stockIndex
                ElseIf (percentYearlyChange < wSheet.Cells(greatestDecreaseStockIndex, 11).Value) Then
                    greatestDecreaseStockIndex = stockIndex
                End If
                
                If (totalStockVolume > wSheet.Cells(greatestVolumeStockIndex, 12).Value) Then
                    greatestVolumeStockIndex = stockIndex
                End If
                ''set up next stock
                stockIndex = stockIndex + 1
                tickerSymbol = wSheet.Cells(i, 1).Value
                curTotalStockVolume = 0
                yearStart = wSheet.Cells(i, 3).Value
            End If
        End If
    Next i
    
    'store great inc dec and vol in three cells
    wSheet.Cells(2, 14).Value = "Greatest % Increase"
    wSheet.Cells(2, 15).Value = wSheet.Cells(greatestIncreaseStockIndex, 9).Value
    wSheet.Cells(2, 16).Value = wSheet.Cells(greatestIncreaseStockIndex, 11).Value
    
    wSheet.Cells(3, 14).Value = "Greatest % Decrease"
    wSheet.Cells(3, 15).Value = wSheet.Cells(greatestDecreaseStockIndex, 9).Value
    wSheet.Cells(3, 16).Value = wSheet.Cells(greatestDecreaseStockIndex, 11).Value
    
    wSheet.Cells(4, 14).Value = "Greatest Volume"
    wSheet.Cells(4, 15).Value = wSheet.Cells(greatestVolumeStockIndex, 9).Value
    wSheet.Cells(4, 16).Value = wSheet.Cells(greatestVolumeStockIndex, 12).Value
    
End Sub

Sub updateStockList(wSheet, stockIndex, tickerSymbol, yearlyChange, percentYearlyChange, totalStockVolume)
    wSheet.Cells(stockIndex, 9).Value = tickerSymbol
    wSheet.Cells(stockIndex, 10).Value = yearlyChange
    wSheet.Cells(stockIndex, 11).Value = percentYearlyChange
    wSheet.Cells(stockIndex, 12).Value = totalStockVolume
                
End Sub
Sub conditionalHighlighting(wSheet, yearlyChange, stockIndex)
    If (yearlyChange > 0) Then
                    wSheet.Cells(stockIndex, 10).Interior.Color = RGB(0, 255, 0)
                    wSheet.Cells(stockIndex, 11).Interior.Color = RGB(0, 255, 0)
                Else
                    wSheet.Cells(stockIndex, 10).Interior.Color = RGB(255, 0, 0)
                    wSheet.Cells(stockIndex, 11).Interior.Color = RGB(255, 0, 0)
                End If
End Sub



Sub loopThroughSheets()
Dim ws As Worksheet
For Each ws In ActiveWorkbook.Worksheets
Call doStuffToASheet(ws)
Next ws

End Sub

Sub clr(ws)
ws.Range("I:L").Clear
End Sub

Sub doOtherStuffToASheet(ws)

End Sub
