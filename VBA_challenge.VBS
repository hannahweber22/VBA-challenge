Attribute VB_Name = "Module1"
Sub vba_challenge()
'Create variable to loop through all worksheets

Dim currentSheet As Worksheet
For Each currentSheet In Worksheets

'Calculate and output stock ticker symbol, yearly change, percent change and total volume

    'Display output titles
    currentSheet.Range("I1").Value = "Ticker"
    currentSheet.Range("J1").Value = "Yearly Change"
    currentSheet.Range("K1").Value = "Percent Change"
    
    'Variables
    Dim ticker As String
    Dim openprice As Double
    Dim closeprice As Double
    Dim yearchange As Double
    Dim percentchange As Double
    Dim totalvolume As Double
    Dim rowcounter As Integer
    Dim rowshift As Integer
    totalvolume = 0
    rowcounter = 0 'keeps track of the rows between when a ticker changes to the next ticker
    rowshift = 2 'shifts the output rows
    
    'Loop through all rows in the current worksheet
    For i = 2 To 753001
        'If the current ticker does not equal the next ticker output the current ticker and calculate the yearly change, percent change, and total volume
        If currentSheet.Cells(i, 1).Value <> currentSheet.Cells(i + 1, 1).Value Then
    
        openprice = currentSheet.Cells(i - rowcounter, 3).Value
        closeprice = currentSheet.Cells(i, 6).Value
    
        'Ticker
        ticker = currentSheet.Cells(i, 1).Value
        currentSheet.Cells(rowshift, 9).Value = ticker
        'Yearly Change (beginning open price - end closing price)
        yearchange = closeprice - openprice
            If yearchange > 0 Then
            currentSheet.Cells(rowshift, 10).Value = yearchange
            currentSheet.Cells(rowshift, 10).Interior.ColorIndex = 4
            Else
            currentSheet.Cells(rowshift, 10).Value = yearchange
            currentSheet.Cells(rowshift, 10).Interior.ColorIndex = 3
            End If
        'Percent Change (beginning open price - ending close price/ beginning open price)
        percentchange = (closeprice - openprice) / openprice
        currentSheet.Cells(rowshift, 11).Value = percentchange
        currentSheet.Cells(rowshift, 11).NumberFormat = "0.00%"
        
        'Total Volume
        totalvolume = totalvolume + currentSheet.Cells(i, 7).Value
        currentSheet.Cells(rowshift, 12).Value = totalvolume
        'Shift down row for outputs
        rowshift = rowshift + 1
        
        'Reset rowcounter
        rowcounter = 0
        
        'Rest total volume
        totalvolume = 0
        
        'If the current ticker is equal to the next ticker, add the current tickers total volume and add to the row counter
        Else
        totalvolume = totalvolume + currentSheet.Cells(i, 7).Value
        rowcounter = rowcounter + 1
        End If
    
    Next i
    
    ' Calculate and output greatest percent increase, greatest percent decrease, and greatest total volume
    
    'Display ouptput titles
    currentSheet.Range("N2") = "Greatest % Increase"
    currentSheet.Range("N3") = "Greatest % Decrease"
    currentSheet.Range("N4") = "Greatest Total Volume"
    currentSheet.Range("L1").Value = "Total Stock Volume"
    currentSheet.Range("O1").Value = "Ticker"
    currentSheet.Range("P1").Value = "Value"
    
    'Variables
    Dim tickercount As Integer
    Dim greatincrease As Double
    Dim greatdecrease As Double
    Dim greatvolume As Double
    tickercount = currentSheet.Cells(Rows.Count, 11).End(xlUp).Row
    greatincrease = currentSheet.Cells(2, 11).Value
    greatdecrease = currentSheet.Cells(2, 11).Value
    greatvolume = currentSheet.Cells(2, 12).Value
    
    For j = 2 To tickercount
        'Greatest Increase
        If currentSheet.Cells(j, 11).Value > greatincrease Then
        greatincrease = currentSheet.Cells(j, 11).Value
        currentSheet.Cells(2, 15).Value = currentSheet.Cells(j, 9).Value
        currentSheet.Cells(2, 16).Value = greatincrease
        currentSheet.Cells(2, 16).NumberFormat = "0.00%"
        End If
        
        'Greatest Decrease
        If currentSheet.Cells(j, 11).Value < greatdecrease Then
        greatdecrease = currentSheet.Cells(j, 11).Value
        currentSheet.Cells(3, 15).Value = currentSheet.Cells(j, 9).Value
        currentSheet.Cells(3, 16).Value = greatdecrease
        currentSheet.Cells(3, 16).NumberFormat = "0.00%"
        End If
        
        'Greatest Total Volume
        If currentSheet.Cells(j, 12).Value > greatvolume Then
        greatvolume = currentSheet.Cells(j, 12).Value
        currentSheet.Cells(4, 15).Value = currentSheet.Cells(j, 9).Value
        currentSheet.Cells(4, 16).Value = greatvolume
        End If
    Next j

Next currentSheet
    
End Sub



