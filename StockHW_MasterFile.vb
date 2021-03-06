Sub stockchecker()

For Each ws In Worksheets

    ' Create a variables to hold the TickerName, YearlyChange, SummaryTableRow, LastCloseValue, and FirstOpenValue.
    ' Will repeatedly use these.
    Dim TickerName As String
    Dim YearlyChange As Double
    Dim PercentChange As Double
    Dim SummaryTableRow As Double
    Dim LastCloseValue As Long
    Dim FirstOpenValue As Long
    Dim TickerVolume As Double
    Dim i As Long
    Dim j as long

    ' Counts the number of rows
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    ' Initially set the SummaryTableRow & FirstOpenValue to be 2 for each row
    SummaryTableRow = 2
    FirstOpenValue = 2
    TickerVolume = 0

        ' Loop through each row using lastrow
        For i = 2 To lastrow
            
            ' Searches for when the value of the next cell is different than that of the current cell
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

                ' Set the TickerName
                TickerName = ws.Cells(i, 1).Value

                ' Print the TickerName to the Summary Table
                ws.Cells(SummaryTableRow, 10).Value = TickerName

                    ' Set LastCloseValue (row) equal to i
                    LastCloseValue = i

                ' Determine value of Yearly Change
                YearlyChange = ws.Cells(LastCloseValue, 6).Value - ws.Cells(FirstOpenValue, 3).Value

                ' Print Yearly Change to the Summary Table
                ws.Cells(SummaryTableRow, 11).Value = YearlyChange

                    ' Determine value of PercentChange when YearlyChange not equal to zero
                        If ws.Cells(FirstOpenValue, 3).Value <> 0 Then

                        PercentChange = YearlyChange / ws.Cells(FirstOpenValue, 3).Value
                        
                        End If
                    
                    ' Determine value of PercentChange when YearlyChange or cells(FirstOpenValue.3).value equal to zero
                    Else
                    
                    PercentChange = 0
                    
                    End If
                 
                ' Print PercentChange (in % format) to the Summary Table
                ws.Cells(SummaryTableRow, 12).Value = FormatPercent(PercentChange)

                ' Add to the TickerVolume
                TickerVolume = TickerVolume + ws.Cells(i, 7).Value

                ' Print the TickerVolume Amount to the Summary Table
                ws.Cells(SummaryTableRow, 13).Value = TickerVolume
                
                    ' Move FirstOpenValue to i and add 1 row
                    FirstOpenValue = i + 1

                    ' Add one to the summary table row
                    SummaryTableRow = SummaryTableRow + 1

                    ' Reset the TickerVolume
                    TickerVolume = 0
                                                                            
            'Else - If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            Else
        
                ' Add to the TickerVolume
                TickerVolume = TickerVolume + ws.Cells(i, 7).Value
                              
            End If
            
        Next i
            
        For j = 2 To lastrow
            
                'Add conditional highlight formatting
                If ws.Cells(j, 12).Value > 0 Then

                    'If value > 0 = green
                    ws.Cells(j, 12).Interior.ColorIndex = 4

                Else

                    'If value < 0 = Red
                    ws.Cells(j, 12).Interior.ColorIndex = 3

                End If
        Next j

Next ws
        
End Sub