Sub stock_analysis()

    ' Declare variables
    Dim total As Double
    Dim i As Long
    Dim change As Double
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim percentChange As Double
    Dim days As Integer
    Dim dailyChange As Double
    Dim averageChange As Double
    Dim ws As Worksheet
    For Each ws In Worksheets
    
        ' Set values for each worksheet
        j = 0
        total = 0
        change = 0
        start = 2
        dailyChange = 0
        
        ' Set the titles for the summary tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        ' Find the row number of the last row with data
        rowCount = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        For i = 2 To rowCount
        
            ' Print the results when a ticker changes
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                ' Stores results in variables
                total = total + ws.Cells(i, 7).Value
                
                ' If there is zero total volume
                If total = 0 Then
                
                    ' Print the results
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = 0
                    ws.Range("K" & 2 + j).Value = "%" & 0
                    ws.Range("L" & 2 + j).Value = 0
                    
                Else
                    ' Get First non zero starting value
                    If ws.Cells(start, 3) = 0 Then
                        For find_value = start To i
                            If ws.Cells(find_value, 3).Value <> 0 Then
                                start = find_value
                                
                                Exit For
                                
                            End If
                            
                        Next find_value
                        
                    End If
                    
                    ' Calculate Yearly and Percent Change
                    change = (ws.Cells(i, 6) - ws.Cells(start, 3))
                    percentChange = change / ws.Cells(start, 3)
                    
                    ' start of the next stock ticker
                    start = i + 1
                    
                    ' print the results
                    ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                    ws.Range("J" & 2 + j).Value = change
                    ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    ws.Range("K" & 2 + j).Value = percentChange
                    ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    ws.Range("L" & 2 + j).Value = total
                    
                    ' Make the cells with positive values green and negatives red
                    Select Case change
                        Case Is > 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        Case Is < 0
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        Case Else
                            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    End Select
                    
                End If
                ' reset variables for new stock ticker
                total = 0
                change = 0
                j = j + 1
                days = 0
                dailyChange = 0
                
            ' If ticker is still the same add results
            Else
                total = total + ws.Cells(i, 7).Value
            End If
        Next i
        
        ' Get the row number for the greatest % increase and decrease
        Dim maxIncreaseRow As Long
        Dim maxDecreaseRow As Long
        Dim maxVolumeRow As Long
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        
        maxIncreaseRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0) + 1
        maxDecreaseRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0) + 1
        maxVolumeRow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0) + 1
        
        ' Get the corresponding tickers for the greatest % increase, decrease, and total volume
        Dim maxIncreaseTicker As String
        Dim maxDecreaseTicker As String
        Dim maxVolumeTicker As String
        
        maxIncreaseTicker = ws.Cells(maxIncreaseRow, 9).Value
        maxDecreaseTicker = ws.Cells(maxDecreaseRow, 9).Value
        maxVolumeTicker = ws.Cells(maxVolumeRow, 9).Value
        
        ' Print the results
        ws.Range("P2").Value = maxIncreaseTicker
        ws.Range("P3").Value = maxDecreaseTicker
        ws.Range("P4").Value = maxVolumeTicker
        ws.Range("Q2").Value = (ws.Cells(maxIncreaseRow, 11).Value) * 100 & "%"
        ws.Range("Q3").Value = (ws.Cells(maxDecreaseRow, 11).Value) * 100 & "%"
        ws.Range("Q4").Value = ws.Cells(maxVolumeRow, 12).Value
    Next ws
End Sub
