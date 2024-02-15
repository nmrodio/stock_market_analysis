Attribute VB_Name = "Module1"
Sub stock_analysis()



'Start of LOOP to go through each worksheet/"Year"
Dim ws As Worksheet
For Each ws In Worksheets



'Inserting Column Headers "Ticker","Yearly Change", "Percent Change", "Total Stock Volume"
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
'Inserting Column & Row Headers for MIN & MAX "Analysis Table"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    

'Dimming for "Ticker Table" with Summarized Results
    Dim total As Double
    Dim i As Long
    Dim change As Single
    Dim j As Integer
    Dim start As Long
    Dim FinalRowA As Long
    Dim percentageChange As Single
    
'Dimming for MIN & MAX "Analysis Table"
    Dim pc_range As Range
    Dim v_range As Range
    Dim Max_pc As Double
    Dim Max_pc_Row As Integer
    Dim Min_pc As Double
    Dim Min_pc_Row As Integer
    Dim Max_v As LongLong
    Dim Max_v_Row As Integer
    

    'start/"assigning" intial values
    j = 0
    total = 0
    change = 0
    start = 2


   'Loop through each row
    FinalRow_A = ws.Cells(Rows.Count, "A").End(xlUp).row  'Finding last populated row in "Column A"
    For i = 2 To FinalRow_A

        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            'stores the results into variables
            total = total + ws.Cells(i, 7).Value
            
            If total = 0 Then
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = 0
                ws.Range("K" & 2 + j).Value = "%" & 0
                ws.Range("L" & 2 + j).Value = 0
                
            Else
            'Find the first "non-zero" starting value
                If ws.Cells(start, 3) = 0 Then
                    For find_value = start To i
                        If ws.Cells(find_value, 3).Value <> 0 Then
                            start = find_value
                            Exit For
                        End If
                    Next find_value
                End If
            'Calculating change for YearlyChange & PercentChange
            change = (ws.Cells(i, 6) - ws.Cells(start, 3))
            percentageChange = change / ws.Cells(start, 3)
            
            'Start of the next ticker
            start = i + 1
            
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
            ws.Range("J" & 2 + j).Value = change
            ws.Range("J" & 2 + j).NumberFormat = "0.00"
            ws.Range("K" & 2 + j).Value = percentageChange
            ws.Range("K" & 2 + j).NumberFormat = "0.00%"
            ws.Range("L" & 2 + j).Value = total
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value

            'Conditional Formatting for colors based on Positive or Negative Returns:
            Select Case change
                Case Is > 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    ws.Range("J" & 2 + j).Interior.ColorIndex = 0
            End Select
            
            Select Case change
                Case Is > 0
                    ws.Range("K" & 2 + j).Interior.ColorIndex = 4
                Case Is < 0
                    ws.Range("K" & 2 + j).Interior.ColorIndex = 3
                Case Else
                    ws.Range("K" & 2 + j).Interior.ColorIndex = 0
            End Select
        End If
        
        'Reset variables
        total = 0
        change = 0
        j = j + 1
        
        
        'If ticker is still the same
    Else
        total = total + ws.Cells(i, 7).Value
    End If
    
    Next i
    
     'Inserting Summary Table for MAX "Percent Change" per Ticker
    Set pc_range = ws.Range("K:K")

        Max_pc = Application.WorksheetFunction.Max(pc_range)
        Max_pc_Row = Application.WorksheetFunction.Match(Max_pc, pc_range, 0)  'Finding the row with the MAX PERCENT CHANGE to match with TICKER NAME
        
            'Outputting results - MAX PERCENT CHANGE with matching TICKER NAME
            ws.Cells(2, 17).Value = Format(Max_pc, "0.00%")
            ws.Cells(2, 16).Value = ws.Cells(Max_pc_Row, 9).Value
            
        ' Inserting Summary Table for MIN "Percent Change" per Ticker
        Min_pc = Application.WorksheetFunction.Min(pc_range)
        Min_pc_Row = Application.WorksheetFunction.Match(Min_pc, pc_range, 0) 'Finding the row with the MIN PERCENT CHANGE to match with TICKER NAME
        
            'Outputting results - MIN PERCENT CHANGE with matching TICKER NAME
            ws.Cells(3, 17).Value = Format(Min_pc, "0.00%")
            ws.Cells(3, 16).Value = ws.Cells(Min_pc_Row, 9).Value
            
    'Inserting Summary Table for Max "Trading Volume" per Ticker
    Set v_range = ws.Range("L:L")
    
        Max_v = Application.WorksheetFunction.Max(v_range)
        Max_v_Row = Application.WorksheetFunction.Match(Max_v, v_range, 0) 'FInding the row with the MAX VOLUME to match with TICKER NAME
        
        
            'Outputting results - MAX VOLUME with matching TICKER NAME
            ws.Cells(4, 17).Value = Max_v
            ws.Cells(4, 16).Value = ws.Cells(Max_v_Row, 9).Value
            
    'Loops to next worksheet/"year"
    Next
    
End Sub

