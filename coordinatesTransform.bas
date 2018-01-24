Sub GetiLastikStackCoordinates()
'Calculate new coordinates of the cropped stack for iLastik analysis

    'Get the data area
    Dim LastRw As Integer
    Dim LastCm As Integer
    LastRw = Columns("A:A").End(xlDown).Row
    LastCm = Rows("1:1").End(xlToRight).Column
    'Get the data area
    
       
    'Ask user for X,Y,Z offset
    Dim xOnset As Long
    Dim yOnset As Long
    Dim zOnset As Long
    
    On Error Resume Next
        xOnset = Application.InputBox("Enter X onset:")
    On Error GoTo 0
    
    On Error Resume Next
        yOnset = Application.InputBox("Enter Y onset:")
    On Error GoTo 0
    
    On Error Resume Next
        zOnset = Application.InputBox("Enter Z onset:")
    On Error GoTo 0
        
'    xOnset = -1276
'   yOnset = -1306
'    zOnset = 0
    

'    On Error Resume Next
'        xOffset = Application.InputBox("Enter X-axis offset")
'        yOffset = Application.InputBox("Enter Y-axis offset")
'        zOffset = Application.InputBox("Enter Z-axis offset")
'    On Error GoTo 0
    'Ask user for X,Y,Z offset
    
    
    'Create new coordinates based on offsets of cropped stack
    For i = 1 To LastRw
        Cells(i, LastCm + 2).Value = Cells(i, LastCm - 5)   'original id
    Next i
    
    For i = 1 To LastRw
        Cells(i, LastCm + 3).Value = Cells(i, LastCm - 4) - xOnset 'new x
    Next i

    For i = 1 To LastRw
        Cells(i, LastCm + 4).Value = Cells(i, LastCm - 3) - yOnset 'new y
    Next i
    
    For i = 1 To LastRw
        Cells(i, LastCm + 5).Value = Cells(i, LastCm - 2) - zOnset 'new z
    Next i
    'Create new coordinates based on offsets of cropped stack


    'Remove objects outside the cropped stack
    Dim Cl As Range
    For Each Cl In Range(Cells(1, LastCm + 2), Cells(LastRw, LastCm + 5)).Cells
        If Not (Cl.Value) < 0 Then
            Cells(Cl.Row, Cl.Column + 5).Value = Cl.Value
        End If
     Next Cl
    'Remove objects outside the cropped stack
    
    
    'Rearrange the new coordinates to display
    Dim newRw As Integer
    newRw = 1
    For i = 1 To LastRw
    ''Each Cl In Range(Cells(1, LastCm + 6), Cells(LastRw, LastCm + 8)).Cells
        If Not IsEmpty(Cells(i, LastCm + 8).Value) Then
            If Not IsEmpty(Cells(i, LastCm + 9).Value) Then
                If Not IsEmpty(Cells(i, LastCm + 10).Value) Then
                    Range(Cells(newRw, LastCm + 12), Cells(newRw, LastCm + 15)).Value = Range(Cells(i, LastCm + 7), Cells(i, LastCm + 10)).Value
                    newRw = newRw + 1
                End If
            End If
        End If
    Next i
    'Rearrange the new coordinates to display
    
    
    'Copy the new coordinates in a new work sheet and clean up the current worksheet
    odsheet = ActiveSheet.Name  'store orignal worksheet name
    Worksheets.Add(After:=ActiveSheet).Name = ActiveSheet.Name + "_NewCoordinates"  'create new work sheet to store new coordinates
    nwsheet = ActiveSheet.Name  'store newly created worksheet name
    
    Cells(1, 1).Value = "Original ID"
    Cells(1, 2).Value = "X"
    Cells(1, 3).Value = "Y"
    Cells(1, 4).Value = "Z"
    Rows("1:1").HorizontalAlignment = xlCenter

    Sheets(odsheet).Select
        Range(Cells(1, LastCm + 12), Cells(newRw, LastCm + 15)).Copy
    Sheets(nwsheet).Select
        Range(Cells(2, 1), Cells(newRw + 1, 4)).Select
        ActiveSheet.Paste
    
    Sheets(odsheet).Select
        Range(Cells(1, LastCm + 2), Cells(LastRw, LastCm + 15)).ClearContents
    'Copy the new coordinates in a new work sheet and clean up the current worksheet
    
    
End Sub

Sub Trim()

    'Get the data area
    Dim LastRw As Integer
    Dim LastCm As Integer
    LastRw = Columns("A:A").End(xlDown).Row
    LastCm = Rows("1:1").End(xlToRight).Column
    'Get the data area

    Dim Cl As Range
    
    For Each Cl In Range(Cells(2, 2), Cells(LastRw, LastCm)).Cells
        If (Cl.Value) > 1024 Then
            Cl.EntireRow.Select
            Selection.ClearContents
        End If
    Next Cl
    
    Range(Cells(1, 1), Cells(LastRw, LastCm)).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp
    
    LastRw = Columns("A:A").End(xlDown).Row
    For Each Cl In Range(Cells(2, 4), Cells(LastRw, LastCm)).Cells
        If (Cl.Value) > 512 Then
            Cl.EntireRow.Select
            Selection.ClearContents
        End If
    Next Cl

    Range(Cells(1, 1), Cells(LastRw, LastCm)).SpecialCells(xlCellTypeBlanks).Delete Shift:=xlUp

End Sub
