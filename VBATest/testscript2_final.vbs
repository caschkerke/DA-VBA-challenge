Sub testrun()

    For Each ws In Worksheets

        'Dim variables (dim format variables later)
        Dim Worksheetname As String
        Worksheetname = ws.Name
        Dim ticker As String
        Dim yearlychange As Double
        yearlychange = 0
        Dim percentchange As Double
        percentchange = 0
        Dim volume As Double
        volume = 0
        Dim TableRow As Integer
        TableRow = 2
        Dim yearopen As Double
        yearopen = Cells(2, 3).Value
        Dim yearclose As Double
        yearclose = 0

        Dim Max As Double
        Max = 0
        Dim Min As Double
        Min = 0
        Dim VolMax As Double
        VolMax = 0
        Dim MaxName As String
        Dim MinName As String
        Dim VolName As String

        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        LastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column

        'Create the new table headings and rows
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"

            For i = 2 To LastRow
                
                'Populate new table with appropriate values from data
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

                For x = 2 To LastRow
    
                    If Cells(x, 11).Value > Max Then
    
                        Max = Cells(x, 11).Value
                        MaxName = Cells(x, 9).Value
                    
                    End If
                
                Next x
                
                    For c = 2 To LastRow
                                    
                        If Cells(c, 11).Value < Min Then
        
                            Min = Cells(c, 11).Value
                            MinName = Cells(c, 9).Value
                            
                        End If
                        
                    Next c
    
                        For v = 2 To LastRow
            
                            If Cells(v, 12).Value > VolMax Then
            
                                VolMax = (Cells(v, 12).Value)
                                VolName = Cells(v, 9).Value
                            
                            End If
                            
                        Next v

        ws.Cells(2, 16).Value = MaxName
        ws.Cells(3, 16).Value = MinName
        ws.Cells(4, 16).Value = VolName
        ws.Cells(2, 17).Value = Max
        ws.Cells(3, 17).Value = Min
        ws.Cells(4, 17).Value = VolMax

            'Conditional and General Formatting after all cell values are filled
            With ws.Range("J2", Range("J2").End(xlDown)).FormatConditions.Add(xlCellValue, xlLess, "=0")
                .Interior.ColorIndex = 3
            End With
            With ws.Range("J2", Range("J2").End(xlDown)).FormatConditions.Add(xlCellValue, xlGreater, "=0")
                .Interior.ColorIndex = 4
            End With
        
        ws.Range("K:K").NumberFormat = "0.00%"
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        ws.Columns("J:R").AutoFit

            'Cycle through all sheets in workbook
            If ActiveSheet.Index = Worksheets.Count Then
                Worksheets(1).Select
            Else
                ActiveSheet.Next.Select
            End If

    Next ws

End Sub