Attribute VB_Name = "Module1"
Sub stock_analysis():

'Identify and define the variables
Dim total As Double
Dim i As Long
Dim change As Double
Dim j As Integer
Dim start As Long
Dim rowCount As Long
Dim percentChange As Double
Dim ws As Worksheet
    
'Loop through all worksheets
For Each ws In Worksheets
    j = 0
    total = 0
    change = 0
    start = 2
        
    'Set title row - label columns with new names
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"

    'Find the row number of the last row with data
    rowCount = ws.Cells(Rows.Count, "A").End(xlUp).Row

    'Go through the whole data set starting at row 2 until the last row
    For i = 2 To rowCount

        'If ticker changes then print results
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

            'Store results in variables
            total = total + ws.Cells(i, 7).Value
            ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value

            'Handle zero total volume- so that we don't divide by 0 in future code
            If total = 0 Then
                percentChange = 0

                'Print the results
                ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
                ws.Range("J" & 2 + j).Value = change
                ws.Range("K" & 2 + j).Value = percentChange
                ws.Range("L" & 2 + j).Value = total

            Else
                percentChange = change / total
            End If

            'Find first non zero starting value
            If ws.Cells(start, 3).Value = 0 Then
                For find_value = start To i
                    If ws.Cells(Find - Value, 3).Value <> 0 Then
                        start = Find - Value
                'Exit the whole for loop all together
                Exit For
            End If
            Next find_value
        End If
     

        'Calculate Change and Percent Change
        change = ws.Cells(i, 6).Value - ws.Cells(start, 3).Value
        percentChange = change / ws.Cells(start, 3).Value

        'Start of the next stock ticker
        start = i + 1

        'Print and format the results,
        'Start with Ticker and Yearly Change
        ws.Range("I" & 2 + j).Value = ws.Cells(i, 1).Value
        ws.Range("J" & 2 + j).Value = change
        
        'Format Yearly Change as 0.00
        ws.Range("J" & 2 + j).NumberFormat = "0.00"
                    
        'Print Percent Change
        ws.Range("K" & 2 + j).Value = percentChange
        
        'Format Precent Change as 0.00%
        ws.Range("K" & 2 + j).NumberFormat = "0.00%"
                    
        'Print Total Stock Volume
        ws.Range("L" & 2 + j).Value = total

        'Color positive Yearly Change green and negative Yearly Change red
        Select Case change
            Case Is > 0
            ws.Range("J" & 2 + j).Interior.ColorIndex = 4
                        
            Case Is < 0
            ws.Range("J" & 2 + j).Interior.ColorIndex = 3
                        
            Case Else
            ws.Range("J" & 2 + j).Interior.ColorIndex = 0
                    
        End Select
                
        'Reset variables for new stock ticker, more values need to be set to 0
        total = 0
        change = 0
        j = j + 1
        start = i + 1
                
        'If ticker is still the same then add results
        Else
            total = total + ws.Cells(i, 7).Value
        End If
        
    Next i

    'Calculate the maximum and minimum Percent Change, and maximum Total Stock Volume, and place them in a separate part in the worksheet
    ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & rowCount)) * 100
    ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & rowCount))

    'Returns one less because header row is not a factor
    'Use Match Function
    increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & rowCount)), ws.Range("K2:K" & rowCount), 0)
    volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & rowCount)), ws.Range("L2:L" & rowCount), 0)

    'Place the final ticker symbol for greatest % of increase and decrease in Percent Change, and Total Stock Volume; respectively.
    ws.Range("P2") = ws.Cells(increase_number + 1, 9)
    ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
    ws.Range("P4") = ws.Cells(volume_number + 1, 9)

    'Label the Row Titles for Summary Table
    ws.Range("O2") = "Greatest % increase"
    ws.Range("O3") = "Greatest % decrease"
    ws.Range("O4") = "Greatest Total Stock Volume"

    'Label the Column Headers for Summary Table
    ws.Range("P1") = "Ticker"
    ws.Range("Q1") = "Value"

Next ws
End Sub
