Sub stock_analysis_all_worksheets()

    ' Set up dimensions
    Dim ws As Worksheet
    Dim total As Double
    Dim i As Long
    Dim j As Integer
    Dim start As Long
    Dim rowCount As Long
    Dim change As Double
    Dim percentChange As Double
    Dim volume As Variant
    Dim increase As Variant
    Dim decrease As Variant
    Dim find_value As Long

    ' Loop through all worksheets
    For Each ws In ThisWorkbook.Worksheets
        
        ' Skip Summary sheet
        If ws.Name <> "Summary" Then
            
            ' Assign initial values
            total = 0
            change = 0
            j = 0
            start = 2

            ' Create column headers
            ws.Range("I1").Value = "Ticker"
            ws.Range("J1").Value = "Quarterly Change"
            ws.Range("K1").Value = "% Change"
            ws.Range("L1").Value = "Total Stock Volume"
    
            ws.Range("O2").Value = "Greatest % Increase"
            ws.Range("O3").Value = "Greatest % Decrease"
            ws.Range("O4").Value = "Greatest Total Volume"
    
            ws.Range("P1").Value = "Ticker"
            ws.Range("Q1").Value = "Value"
    
            ' Get the last row of data
            rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row
    
            ' Loop through rows
            For i = 2 To rowCount

                ' Skip blank rows in column A to prevent errors
                If IsEmpty(ws.Cells(i, 1).Value) Then
                    Debug.Print "Skipping empty row: " & i
                    GoTo NextRow
                End If
    
                ' If ticker changes, print results
                If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                    ' Store results
                    total = total + ws.Cells(i, 7).Value
            
                    If total = 0 Then
                        ws.Range("I" & (2 + j)).Value = ws.Cells(i, 1).Value
                        ws.Range("J" & (2 + j)).Value = 0
                        ws.Range("K" & (2 + j)).Value = "0%"
                        ws.Range("L" & (2 + j)).Value = 0
                    Else
                        ' Find first non-zero starting value
                        If ws.Cells(start, 3) = 0 Then
                            For find_value = start To i
                                If ws.Cells(find_value, 3).Value <> 0 Then
                                    start = find_value
                                    Exit For
                                End If
                            Next find_value
                        End If
                
                        ' Calculate Change
                        If ws.Cells(start, 3).Value <> 0 Then
                            change = (ws.Cells(i, 6).Value) - ws.Cells(start, 3).Value
                            percentChange = change / ws.Cells(start, 3).Value
                        Else
                            change = 0
                            percentChange = 0
                        End If
                
                        start = i + 1
                
                        ' Print results
                        ws.Range("I" & (2 + j)).Value = ws.Cells(i, 1).Value
                        ws.Range("J" & (2 + j)).Value = change
                        ws.Range("J" & (2 + j)).NumberFormat = "0.00"
                        ws.Range("K" & (2 + j)).Value = percentChange
                        ws.Range("K" & (2 + j)).NumberFormat = "0.00%"
                        ws.Range("L" & (2 + j)).Value = total

                        ' Color formatting
                        Select Case change
                            Case Is > 0
                                ws.Range("J" & (2 + j)).Interior.ColorIndex = 4 ' Green
                            Case Is < 0
                                ws.Range("J" & (2 + j)).Interior.ColorIndex = 3 ' Red
                            Case Else
                                ws.Range("J" & (2 + j)).Interior.ColorIndex = 6 ' Yellow
                        End Select
                    End If
            
                    ' Reset variables for new stock ticker
                    total = 0
                    change = 0
                    j = j + 1
                Else
                    ' If ticker is still the same, add results
                    total = total + ws.Cells(i, 7).Value
                End If

NextRow:
            Next i
        
            ' Find and display the greatest increase, decrease, and total volume
            ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & rowCount))
            ws.Range("Q2").NumberFormat = "0.00%"

            ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & rowCount))
            ws.Range("Q3").NumberFormat = "0.00%"

            ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

            ' Handle potential errors in Match function
            On Error Resume Next
            increase = Application.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
            decrease = Application.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
            volume = Application.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)
            On Error GoTo 0

            ' Assign ticker symbols for the highest values, ensuring they exist
            If Not IsError(increase) Then ws.Range("P2").Value = ws.Cells(increase + 1, "I").Value
            If Not IsError(decrease) Then ws.Range("P3").Value = ws.Cells(decrease + 1, "I").Value
            If Not IsError(volume) Then ws.Range("P4").Value = ws.Cells(volume + 1, "I").Value
        End If
    
    Next ws

End Sub



