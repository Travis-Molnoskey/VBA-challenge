Sub WallStreet()
    Dim volume As Double
    Dim ticker As String
    Dim year_start As Double
    Dim year_end As Double
    Dim lastrow As Double
    Dim delta As Double
    Dim change As Double
    Dim i As Double
    Dim j As Double
    Dim ws As Worksheet
    
    Dim bigincreaseticker As String
    Dim bigdecreaseticker As String
    Dim bigvolumeticker As String
    Dim bigincrease As Double
    Dim bigdecrease As Double
    Dim bigvolume As Double

For Each ws In Worksheets

    'find last row in data
    lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row + 1
    
    
    i = 2                                   'initialize i
    j = 2                                   'initialize j
    year_start = ws.Cells(i, 6).Value          'initialize first ticker close
    
    'reset largest increase/decrease counters to 0 at each ws
    bigincrease = 0
    bigdecrease = 0
    bigvolume = 0
    
    'title the summary columns
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    
    For i = 2 To lastrow                    'loop through all rows of data
        
        'determine if ticker symbol changes
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            ticker = ws.Cells(i, 1).Value      'set ticker symbol to last row before change
            volume = volume + ws.Cells(i, 7).Value 'add final volume number
            year_end = ws.Cells(i, 6).Value    'assign last close value to end of year close
            
            delta = year_end - year_start   'calculate change in stock value
            change = (year_end - year_start) / year_start 'calculate percentage change in stock value
            
            
            ws.Cells(j, 9).Value = ticker      'print ticker value
            
            'print yearly change with formatting for green if >0 red if else
            ws.Cells(j, 10).Value = Format(delta, "#,##0.00")
            If delta > 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(j, 10).Interior.ColorIndex = 3
            End If
            
            'print percent change in percent format
            ws.Cells(j, 11).Value = Format(change, "Percent")
            
            'print volume
            ws.Cells(j, 12).Value = volume
            
            If volume > bigvolume Then
                bigvolume = volume
                bigvolumeticker = ws.Cells(i, 1).Value
            End If
            
            If change > bigincrease Then
                bigincrease = change
                bigincreaseticker = ws.Cells(i, 1).Value
            End If
            
            If change < bigdecrease Then
                bigdecrease = change
                bigdecreaseticker = ws.Cells(i, 1).Value
            End If
            
            j = j + 1
            volume = 0
        Else
            'if ticker does not change, add volume to volume running total
            volume = volume + ws.Cells(i, 7).Value
            
        End If
        
        
        
        
    Next i
    
    'print out headers for largest changes per year
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"
    
    'print out values for largest changes per year
    ws.Cells(2, 15).Value = bigincreaseticker
    ws.Cells(3, 15).Value = bigdecreaseticker
    ws.Cells(4, 15).Value = bigvolumeticker
    ws.Cells(2, 16).Value = Format(bigincrease, "percent")
    ws.Cells(3, 16).Value = Format(bigdecrease, "Percent")
    ws.Cells(4, 16).Value = bigvolume
    
    ws.Columns("I:P").AutoFit
    
Next ws
    
    
    
End Sub