
Sub Homework:

    For Each ws In Worksheets

        Dim WorksheetName as String
        WorksheetName = ws.Name

        Dim LastRow as Long
            LastRow = ws.Cells(Rows.Count,1).End(xlUp).Row

            Dim Ticker as String 

            Dim YearOpening as Double
            YearOpening = ws.Cells(2,3).Value 

            Dim YearlyChange as Double

            Dim PercentChange as Double

            Dim TotalStock as Double
            TotalStock = 0

            Dim TableRow as Integer
            TableRow = 2


            For i = 2 to  LastRow

                If ws.Cells(i+1,1).Value<>ws.Cells(i,1).Value Then

                    Ticker = ws.Cells(i,1).Value

                    YearlyChange = ws.Cells(i,6).Value - YearOpening

                        If YearOpening = 0 Then

                        PercentChange = 0

                        Else PercentChange = YearlyChange / YearOpening

                        End If

                    TotalStock = TotalStock + ws.Cells(i,7).Value

                    ws.Range("J" & TableRow).Value = Ticker

                    ws.Range("K" & TableRow).Value = YearlyChange

                        If YearlyChange > 0 Then

                            ws.Range("K" & TableRow).Interior.Color = vbGreen

                        Else

                            ws.Range("K" & TableRow).Interior. Color = vbRed

                        End If

                    ws.Range("L" & TableRow).Value = PercentChange
                    ws.Range("L" & TableRow).NumberFormat = "0.00%"

                    ws.Range("M" & TableRow).Value = TotalStock

                    TableRow = TableRow + 1

                    YearOpening = ws.Cells(i+1,3).Value

                    TotalStock = 0

                Else

                    TotalStock = TotalStock + ws.Cells(i,7).Value
                    
                End If   

            Next i

    Next ws

End Sub