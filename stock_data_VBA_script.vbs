Sub Tickerstocks()
    Dim ws As Worksheet

For Each ws In Worksheets
    
    Dim lastrow As Long
    Dim lastsummary As Long
    Dim i As Long
    Dim total As Double
    Dim yearchange As Double
    Dim percent As Double
    Dim openprice As Double
    Dim lines As Long
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim biggestvolume As Double
    
    'initialize
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    lastsummary = ws.Cells(Rows.Count, 9).End(xlUp).Row
    lines = 2
    total = 0
    openprice = 2
    greatestIncrease = 0
    greatestDecrease = 0
    biggestvolume = 0
    
    For i = 2 To lastrow
        total = total + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            yearchange = ws.Cells(i, 6).Value - ws.Cells(openprice, 3).Value
            percent = yearchange / ws.Cells(openprice, 3).Value
            
            'display summaries for each ticker
            ws.Cells(lines, 9).Value = ws.Cells(i, 1)
            ws.Cells(lines, 10).Value = yearchange
            ws.Cells(lines, 11).Value = percent
            ws.Cells(lines, 12).Value = total
            
            'set color to year change cell for positive or negative
            If yearchange < 0 Then
                ws.Cells(lines, 10).Interior.ColorIndex = 3
            ElseIf yearchange > 0 Then
                ws.Cells(lines, 10).Interior.ColorIndex = 4
            End If
            
            'reset variables for next ticker
            lines = lines + 1
            openprice = i + 1
            total = 0
            yearchange = 0
        End If
        
        'headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
    Next i
    
    'this for loop determines the greatest increase, decrease, and total volume
    For i = 2 To lastsummary
        'determine greatest % increase
        If ws.Cells(i, 11).Value > greatestIncrease Then
           greatestIncrease = ws.Cells(i, 11).Value
            ws.Cells(2, 17).Value = greatestIncrease
            ws.Cells(2, 16).Value = ws.Cells(i, 1).Value
            ws.Cells(2, 15).Value = "Greatest % Increase"
        End If
        
        'determine greatest % decrease
        If ws.Cells(i, 11).Value < greatestDecrease Then
            greatestDecrease = ws.Cells(i, 11).Value
            ws.Cells(3, 17).Value = greatestDecrease
            ws.Cells(3, 16).Value = ws.Cells(i, 1).Value
            ws.Cells(3, 15).Value = "Greatest % Decrease"
        End If
        
        'determine greatest total stock volume
        If ws.Cells(i, 12).Value > biggestvolume Then
            biggestvolume = ws.Cells(i, 12).Value
            ws.Cells(4, 17).Value = biggestvolume
            ws.Cells(4, 16).Value = ws.Cells(i, 1).Value
            ws.Cells(4, 15).Value = "Greatest Total Volume"
        End If
    Next i
        
    'autofit and percent style columns
    ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
    ws.Range("Q2:Q3").NumberFormat = "0.00%"
    ws.Columns("A:Q").AutoFit
Next ws

End Sub

