'Script needs to loop through all the stocks for one year
'for each run, (run = year?) need to take the following info:
'The ticker symbol
'yearly change from opening price at the beginning of a given year to the closing price at the end of that year
'the percent change from opening price at the beginning of a given year to the closing price at the end of that year
'the total stock volume of the stock

'need conditional formatting that will highlight positive changes in green, and negative changes in red

Sub testrun()

    For Each ws In Worksheets
    
        Dim Worksheetname As String
        Worksheetname = ws.Name
        Dim ticker As String
        Dim yearlychange As Double
        yearlychange = 0
        Dim percentchange As Double
        percentchange = 0
        Dim volume As String
        volume = 0
        Dim TableRow As Integer
        TableRow = 2
        Dim yearopen As Double
        yearopen = Cells(2, 3).Value
        Dim yearclose As Double
        yearclose = 0

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"

            For i = 2 To LastRow

                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

                    ticker = Cells(i, 1).Value
                    volume = volume + Cells(i, 7).Value
                    yearclose = Cells(i, 6).Value

                    yearlychange = (yearopen - yearclose) * (-1)
                    percentchange = (((yearclose - yearopen) - 1) / 100)
                        
                    Range("I" & TableRow).Value = ticker
                    Range("J" & TableRow).Value = yearlychange
                    Range("K" & TableRow).Value = percentchange
                    Range("L" & TableRow).Value = volume
                        
                    TableRow = TableRow + 1
                    volume = 0
                    yearclose = 0
                    yearopen = Cells(i + 1, 3).Value
                
                Else

                    volume = CDbl(Trim(volume + Cells(i, 7).Value))
                    
                End If
            
            Next i

            'Conditional Formatting after all cell values are filled
            Dim negcondition As FormatCondition, poscondition As FormatCondition
            Dim rng As Range

                Set rng = Range("J2" & LastRow)
                Set negcondition = rng.FormatConditions.Add(xlCellValue, xlLess, "=0")
                Set poscondition = rng.FormatConditions.Add(xlCellValue, xlGreater, "=0")

                    With negcondition
                        .Interior.ColorIndex = 3
                    End With
                        
                    With poscondition
                        .Interior.ColorIndex = 4
                    End With

            Range("K:K").NumberFormat = "0.00%"

            'Cycle through all sheets in workbook
            If ActiveSheet.Index = Worksheets.Count Then
                Worksheets(1).Select
            Else
                ActiveSheet.Next.Select
            End If

    Next ws

End Sub