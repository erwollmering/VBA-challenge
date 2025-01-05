Attribute VB_Name = "Module1"
Sub Challenge2()

    Dim ws As Worksheet
    For Each ws In Worksheets
        Dim i As Long
        Dim lastrow As Long

        ' Variables for PART 2
        Dim earliestRow As Long
        Dim latestRow As Long
        Dim openingPrice As Double
        Dim closingPrice As Double
        Dim priceChange As Double
        
        'Variables for PART 3
        Dim pricechangepercent As Double

        ' Set an initial variable for holding the ticker name
        Dim Ticker As String

        ' Set an initial variable for holding the total per ticker type
        Dim Ticker_Total As Double
        Ticker_Total = 0

        ' Keep track of the location for each ticker type in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
        
        ' Add headers to the current worksheet
        ws.Cells(1, 9).Value = "Ticker Type" ' Column I
        ws.Cells(1, 12).Value = "Total Stock Volume" ' Column L
        ' Column J PART 2
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        'PART 5
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        
        'PART 6
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
                
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Loop through all ticker rows
        For i = 2 To lastrow

            ' Check if we are still within the same ticker type, if it is not...
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the Ticker type name
                Ticker = ws.Cells(i, 1).Value

                ' Add to the Ticker Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 3).Value

                ' Calculate earliest and latest rows for the ticker type (Injected logic)
                earliestRow = i - (i - 2) ' Earliest date is first occurrence
                latestRow = i ' Latest date is the last occurrence in this group
                
                ' Extract opening and closing prices (Injected logic)
                openingPrice = ws.Cells(earliestRow, 3).Value ' Column C
                closingPrice = ws.Cells(latestRow, 6).Value ' Column F

                ' Calculate the price change (Injected logic)
                priceChange = closingPrice - openingPrice

                
                'Calculate the percentage change (injected logic) PART 3
                pricechangepercent = (priceChange / openingPrice) * 100


                ' Print the Ticker type in the Summary Table
                ws.Range("I" & Summary_Table_Row).Value = Ticker

                ' Print the Ticker type Amount to the Summary Table
                ws.Range("L" & Summary_Table_Row).Value = Ticker_Total

                ' Print the price change to the Summary Table (Injected logic)
                ws.Range("J" & Summary_Table_Row).Value = priceChange
                
                 'Print the percentage change to the Summary Table (injected logic) PART 3
                ws.Range("K" & Summary_Table_Row).Value = pricechangepercent
                
                      'PART 4 format red and green
            
                          ' Find the last row in Column K
                            lastrow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row ' Column K = 11

                         ' Loop through each cell in Column K
                            For Each cell In ws.Range("K2:K" & lastrow)
                            If IsNumeric(cell.Value) Then ' Check if the cell contains a number
                            If cell.Value > 0 Then
                            cell.Interior.ColorIndex = 4 ' Green for positive
                            ElseIf cell.Value < 0 Then
                            cell.Interior.ColorIndex = 3 ' Red for negative
                            Else
                            cell.Interior.ColorIndex = xlNone ' Clear formatting for zero
                            End If
                             Else
                            cell.Interior.ColorIndex = xlNone ' Clear formatting for non-numeric cells
                             End If
                        Next cell
        
                ' Add one to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1

                ' Reset the Ticker Total
                Ticker_Total = 0

            ' If the cell immediately following a row is the same ticker type...
            Else

                ' Add to the Brand Total
                Ticker_Total = Ticker_Total + ws.Cells(i, 6).Value

            End If

        Next i

 ' PART 6: Greatest increase, decrease, and total volume
      Dim maxvalue As Double
        Dim minvalue As Double
        Dim maxvolume As Double
        Dim maxIncreaseRow As Long
        Dim maxDecreaseRow As Long
        Dim maxVolumeRow As Long
        Dim maxIncreaseTicker As String
        Dim maxDecreaseTicker As String
        Dim maxVolumeTicker As String

        ' Find the last row in Column K (Percent Change)
        lastrow = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row ' Column K = 11

        ' PART 6: Find greatest % increase
        maxvalue = Application.WorksheetFunction.Max(ws.Range("K2:K" & lastrow))
        For i = 2 To lastrow
            If ws.Cells(i, 11).Value = maxvalue Then ' Find the exact row of max value
                maxIncreaseRow = i
                Exit For
            End If
        Next i
        maxIncreaseTicker = ws.Cells(maxIncreaseRow, 9).Value ' Get Ticker Type from Column I

        ' PART 6: Find greatest % decrease
        minvalue = Application.WorksheetFunction.Min(ws.Range("K2:K" & lastrow))
        For i = 2 To lastrow
            If ws.Cells(i, 11).Value = minvalue Then ' Find the exact row of min value
                maxDecreaseRow = i
                Exit For
            End If
        Next i
        maxDecreaseTicker = ws.Cells(maxDecreaseRow, 9).Value ' Get Ticker Type from Column I

        ' PART 6: Find greatest total volume
        lastrow = ws.Cells(ws.Rows.Count, 12).End(xlUp).Row ' Column L = 12
        maxvolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & lastrow))
        For i = 2 To lastrow
            If ws.Cells(i, 12).Value = maxvolume Then ' Find the exact row of max volume
                maxVolumeRow = i
                Exit For
            End If
        Next i
        maxVolumeTicker = ws.Cells(maxVolumeRow, 9).Value ' Get Ticker Type from Column I

        ' Output the results
        ws.Range("P2").Value = maxIncreaseTicker ' Greatest % Increase Ticker
        ws.Range("Q2").Value = maxvalue         ' Greatest % Increase Value

        ws.Range("P3").Value = maxDecreaseTicker ' Greatest % Decrease Ticker
        ws.Range("Q3").Value = minvalue          ' Greatest % Decrease Value

        ws.Range("P4").Value = maxVolumeTicker   ' Greatest Total Volume Ticker
        ws.Range("Q4").Value = maxvolume         ' Greatest Total Volume Value

        
    Next ws

End Sub


