Attribute VB_Name = "Module1"
Sub stock_analysis():

    'Set up dimensions
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
    
    'Create new column headers
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Quarterly Change"
    Range("K1").Value = "% Change"
    Range("L1").Value = "Total Stock Volume"
    
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    
    'Assign initial values to the variables
    total = 0
    change = 0
    j = 0
    start = 2
    
    'Get the row number of the last row with data
    rowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'For loop
    For i = 2 To rowCount
    
        'If ticker changes, then print the results
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            'Stores the results into a variable
            total = total + Cells(i, 7).Value
            
            If total = 0 Then
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = 0
                Range("K" & 2 + j).Value = 0 & "%"
                Range("L" & 2 + j).Value = 0
    
            Else
                If Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
                
                change = (Cells(i, 6) - Cells(start, 3))
                percentChange = change / Cells(start, 3)
                
                start = i + 1
                
                Range("I" & 2 + j).Value = Cells(i, 1).Value
                Range("J" & 2 + j).Value = change
                Range("K" & 2 + j).Value = percentChange
                Range("L" & 2 + j).Value = total

                Select Case change
                    Case Is > 0
                        Range("J" & 2 + j).Interior.ColorIndex = 4
                    Case Is < 0
                        Range("J" & 2 + j).Interior.ColorIndex = 3
                    Case Else
                        Range("J" & 2 + j).Interior.ColorIndex = 6
                End Select

            End If
            
            total = 0
            change = 0
            j = j + 1
            days = 0

        ' If ticker is still the same add results
            Else
                total = total + Cells(i, 7).Value

            End If
            
        Next i
        
        Range("Q2") = "%" & WorksheetFunction.Max(Range("K2:K" & rowCount)) * 100
        Range("Q3") = "%" & WorksheetFunction.Min(Range("K2:K" & rowCount)) * 100
        Range("Q4") = WorksheetFunction.Max(Range("L2:L" & rowCount))

        increase = WorksheetFunction.Match(WorksheetFunction.Max(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
        decrease = WorksheetFunction.Match(WorksheetFunction.Min(Range("K2:K" & rowCount)), Range("K2:K" & rowCount), 0)
        volume = WorksheetFunction.Match(WorksheetFunction.Max(Range("L2:L" & rowCount)), Range("L2:L" & rowCount), 0)

        ' final ticker symbol for  total, greatest % of increase and decrease, and average
        Range("P2") = Cells(increase + 1, 9)
        Range("P3") = Cells(decrease + 1, 9)
        Range("P4") = Cells(volume + 1, 9)

End Sub



