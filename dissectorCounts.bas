Sub ColoRworkSpace()
'Select the dissector working area and paint it light yellow

ActiveSheet.range(Cells(4, 2), Cells(22, 20)).Select
Selection.Interior.ColorIndex = 36
ActiveSheet.range(Cells(4, 25), Cells(22, 43)).Select
Selection.Interior.ColorIndex = 36
    
'paint all dissectors based on their value
    For i = 5 To 21
        For j = 3 To 19
            Value = Cells(i, j).Value   'get value of each cell in range
            If Not (IsNumeric(Value)) Then
                If ((Value = "A" Or Value = "B" Or Value = "C")) Then
                    Cells(i, j).Interior.ColorIndex = 45    'paint discardable dissectors as light orange
                Else
                    Cells(i, j).Interior.ColorIndex = 40    'paint semi-usable dissectors as tan
                End If
            End If
        Next j
    Next i
    
    For i = 5 To 21
        For j = 26 To 42
            Value = Cells(i, j).Value   'get value of each cell in range
            If Not (IsNumeric(Value)) Then
                If ((Value = "A" Or Value = "B" Or Value = "C")) Then
                    Cells(i, j).Interior.ColorIndex = 45    'paint discardable dissectors as light orange
                Else
                    Cells(i, j).Interior.ColorIndex = 40    'paint discardable dissectors as tan
                End If
            End If
        Next j
    Next i
    
    
End Sub
Sub FormatCalculationArea()
'format the calculation area

    range("T4:T4,AQ4:AQ4,B22:B22,Y22:Y22").Select
    Selection.Value = "rawSum"
    range("U4:U4,AR4:AR4,B23:B23,Y23:Y23").Select
    Selection.Value = "Sum"
    range("V4:V4,AS4:AS4,B24:B24,Y24:Y24").Select
    Selection.Value = "rawDen"
    range("W4:W4,AT4:AT4,B25:B25,Y25:Y25").Select
    Selection.Value = "Den"
    
    range("T4:W25,B22:S25,AQ4:AT25,Y22:AP25").Select    'select calculation area
    With Selection                      'format cell
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    With Selection.Font                 'format font
        .Name = "Arial"
        .Size = 3
        .Bold = True
    End With
    
    range("T5:U23,C22:S23,AQ5:AR23,Z22:AP23").Select
    Selection.NumberFormat = "0"
    range("V5:W25,C24:U25,AS5:AT25,Z24:AR25").Select
    Selection.NumberFormat = "0.00"

    
End Sub
Sub GetDensityMeasurement()

'density measurement for each row and column of whole working area
''Function GetMeasurement(workingRange As range)
 
 Dim WorkingArea As range   'workinig area operator
 Dim Rw As range    'row operator
 Dim Cm As range    'column operator
 Dim Cl As range    'cell operator
 Dim lastRw As Integer  'last row number of working area
 Dim lastCm As Integer  'last column number of working area
 
 'Set WorkingArea = range("C5:S21")
 Set WorkingArea = range("Z5:AP21")
 
 lastRw = WorkingArea.Rows(0).row + WorkingArea.Rows.Count
 lastCm = WorkingArea.Columns(0).column + WorkingArea.Columns.Count
 
 
    'Get Row Measurement
    For Each Rw In WorkingArea.Rows 'row-wise measurement
        rawData = 0
        Data = 0
        rawSum = 0
        Sum = 0
        For Each Cl In Rw.Cells
            Value = Cl.Value   'get value of each cell in current row
            If (IsNumeric(Value)) And Len(Value) > 0 Then  'get semi-usable dissector value with raw parameters
                rawData = rawData + 1
                Data = Data + 1
                rawSum = rawSum + Value
                Sum = Sum + Value
            Else
                If Cl.Interior.ColorIndex = 40 Then
                    rawData = rawData + 1
                    rawSum = rawSum + Val(Right(Value, 1))
                End If
            End If
        Next Cl
        Cells(Rw.row, lastCm + 1).Value = rawSum
        Cells(Rw.row, lastCm + 2).Value = Sum
        Cells(Rw.row, lastCm + 3).Value = rawSum / rawData
        Cells(Rw.row, lastCm + 4).Value = Sum / Data
    Next Rw
    
    'Get Column Measurement
    For Each Cm In WorkingArea.Columns 'column-wise measurement
        rawData = 0
        Data = 0
        rawSum = 0
        Sum = 0
        For Each Cl In Cm.Cells
            Value = Cl.Value   'get value of each cell in current row
            If (IsNumeric(Value)) And Len(Value) > 0 Then  'get semi-usable dissector value with raw parameters
                rawData = rawData + 1
                Data = Data + 1
                rawSum = rawSum + Value
                Sum = Sum + Value
            Else
                If Cl.Interior.ColorIndex = 40 Then
                    rawData = rawData + 1
                    rawSum = rawSum + Val(Right(Value, 1))
                End If
            End If
        Next Cl
        Cells(lastRw + 1, Cm.column).Value = rawSum
        Cells(lastRw + 2, Cm.column).Value = Sum
        Cells(lastRw + 3, Cm.column).Value = rawSum / rawData
        Cells(lastRw + 4, Cm.column).Value = Sum / Data
    Next Cm
    
    'Get Area Measurement
        rawData = 0
        Data = 0
        rawSum = 0
        Sum = 0
        
    For Each Cl In WorkingArea.Cells 'area-wise measurement
        Value = Cl.Value   'get value of each cell in area
            If (IsNumeric(Value)) And Len(Value) > 0 Then  'get semi-usable dissector value with raw parameters
                rawData = rawData + 1
                Data = Data + 1
                rawSum = rawSum + Value
                Sum = Sum + Value
            Else
                If Cl.Interior.ColorIndex = 40 Then
                    rawData = rawData + 1
                    rawSum = rawSum + Val(Right(Value, 1))
                End If
            End If
    Next Cl
        Cells(lastRw + 1, lastCm + 1).Value = rawSum
        Cells(lastRw + 2, lastCm + 2).Value = Sum
        Cells(lastRw + 3, lastCm + 3).Value = rawSum / rawData
        Cells(lastRw + 4, lastCm + 4).Value = Sum / Data
    
''End Function

''GetMeasurement (range("C5:Q19"))

End Sub
