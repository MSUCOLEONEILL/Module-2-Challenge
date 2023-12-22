Attribute VB_Name = "Module1"
Sub stock()

    ' Define variables
    Dim lastrow As Long
    Dim Ticker As String
    Dim openingprice As Double
    Dim closingprice As Double
    Dim yearlychange As Double
    Dim percentchange As Double
    Dim totalvolume As Double
    Dim summaryrow As Integer
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    
    For Each ws In Worksheets
    
    ' find last row in Ws.
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' sum table headers
    ws.Cells(1, 9).Value = "Ticker"
   ws.Cells(1, 10).Value = "Yearly Change"
   ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

    ' Identify output for summary table
    summaryrow = 2

    ' Initialize variables for greatest values
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0

    ' Loop through each row of data
    For i = 2 To lastrow

        ' check if the current row contains a new ticker
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            ' assign the ticker symbol
            Ticker = ws.Cells(i, 1).Value
            ' assign the opening price
            openingprice = ws.Cells(i, 3).Value
            ' reset total volume for the new ticker
            totalvolume = 0
        End If

        ' calculate the total volume
        totalvolume = totalvolume + ws.Cells(i, 7).Value

        ' verify the specified row is the last of that ticker symbol
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
            ' identify the closing price
            closingprice = ws.Cells(i, 6).Value

            ' calculate yearly change
            yearlychange = closingprice - openingprice

            ' avoid division by zero error
            If openingprice <> 0 Then
                ' calculate percent change
                percentchange = (yearlychange / openingprice) * 100
            Else
                percentchange = 0
            End If

            ' output data to summary table
           ws.Cells(summaryrow, 9).Value = Ticker
           ws.Cells(summaryrow, 10).Value = yearlychange
            ws.Cells(summaryrow, 11).Value = percentchange
           ws.Cells(summaryrow, 12).Value = totalvolume

            ' update greatest values
            If percentchange > greatestIncrease Then
                greatestIncrease = percentchange
                greatestIncreaseTicker = Ticker
            End If

            If percentchange < greatestDecrease Then
                greatestDecrease = percentchange
                greatestDecreaseTicker = Ticker
            End If

            If totalvolume > greatestVolume Then
                greatestVolume = totalvolume
                greatestVolumeTicker = Ticker
            End If

            ' move to the next row in the summary table
            summaryrow = summaryrow + 1
        End If
    Next i

    ' Output greatest values
   ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    ws.Cells(2, 16).Value = greatestIncreaseTicker
    ws.Cells(3, 16).Value = greatestDecreaseTicker
   ws.Cells(4, 16).Value = greatestVolumeTicker

    ws.Cells(2, 17).Value = greatestIncrease
    ws.Cells(3, 17).Value = greatestDecrease
    ws.Cells(4, 17).Value = greatestVolume
    
Next ws

End Sub
