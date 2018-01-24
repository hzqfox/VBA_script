Sub GetCroppedStackCoordinates()
'Calculate new coordinates of the cropped stack for iLastik analysis

' DataTransfer Macro
' Macro recorded 8/19/2014 by hzq.fox@gmail.com
 oldSheet = ActiveSheet.Name  'store orignal worksheet name
 Worksheets.Add(After:=ActiveSheet).Name = ActiveSheet.Name + "_intermediateSheet"  'create new work sheet to store new coordinates
 NewSheet = ActiveSheet.Name  'store newly created worksheet name
 Sheets(NewSheet).Select
 cRow = 1
 crRow = cRow
 Cells(cRow, 1).Value = "sample" + oldSheet + "L"
 Cells(cRow, 2).Value = "sample" + oldSheet + "R"
 Cells(cRow, 3).Value = Cells(cRow, 1).Value + "_rawData"
 Cells(cRow, 4).Value = Cells(cRow, 2).Value + "_rawData"
 Cells(cRow, 5).Value = "sample" + oldSheet + "_1st_Quadrant"
 Cells(cRow, 6).Value = "sample" + oldSheet + "_3rd_Quadrant"
 
 
 Sheets(oldSheet).Select

 Dim WorkingArea As Range   'workinig area operator
 Dim Rw As Range    'row operator
 Dim Cm As Range    'column operator
 Dim Cl As Range    'cell operator
 Dim firstRw As Integer  'first row number of working area
 Dim firstCm As Integer  'first column number of working area
 Dim lastRw As Integer  'last row number of working area
 Dim lastCm As Integer  'last column number of working area

 'Ask user for dissector range
 On Error Resume Next
    dissectorRange = Application.InputBox("Enter dissector range:")
 On Error GoTo 0
    
 firstRw = 5
 firstCm = 3
 lastRw = firstRw + dissectorRange - 3
 lastCm = firstCm + dissectorRange - 3
 
 
    
 'get dissector data on the left panel
 Set WorkingArea = Range(Cells(firstRw, firstCm), Cells(lastRw, lastCm))
 For Each Rw In WorkingArea.Rows    'iterate each row in dissector range
    For Each Cl In Rw.Cells     'iterate each cell in row
        cValue = Cl.Value        'get value of each cell
        If (IsNumeric(cValue)) And Len(cValue) > 0 Then
            cRow = cRow + 1
            crRow = crRow + 1
            Sheets(NewSheet).Select    'select the newly created worksheet
            Cells(cRow, 1).Value = cValue 'Copy the new coordinates in a new work sheet and clean up the current worksheet
            Cells(crRow, 3).Value = cValue 'Copy the new coordinates in a new work sheet and clean up the current worksheet
        Else
            If Cl.Interior.ColorIndex = 40 Then
                crRow = crRow + 1
                Cells(crRow, 3).Value = Val(Right(cValue, 1))
            End If
        End If
    Next Cl
 Next Rw
 
 
'get dissector data of 1st and 3rd quadrant
 cRow = 1
 crRow = 1
 Sheets(oldSheet).Select
 Set WorkingArea = Range(Cells(firstRw, firstCm), Cells(lastRw, lastCm))
 For Each Rw In WorkingArea.Rows 'iterate each row in dissector range
    If Rw.Row Mod 2 = 1 Then       'if the row is an odd number
        For Each Cl In Rw.Cells     'iterate each cell in row
            cValue = Cl.Value           'get value of each cell
            If (IsNumeric(cValue)) And Len(cValue) > 0 Then
                cRow = cRow + 1
                Sheets(NewSheet).Select    'select the newly created worksheet
                Cells(cRow, 5).Value = cValue 'Copy the new coordinates in a new work sheet and clean up the current worksheet
            End If
        Next Cl
    Else
        For Each Cl In Rw.Cells     'iterate each cell in row
            cValue = Cl.Value           'get value of each cell
            If (IsNumeric(cValue)) And Len(cValue) > 0 Then
                crRow = crRow + 1
                Sheets(NewSheet).Select    'select the newly created worksheet
                Cells(crRow, 6).Value = cValue 'Copy the new coordinates in a new work sheet and clean up the current worksheet
            End If
        Next Cl
    End If
 Next Rw
 
 
 
 
 'get dissector data on the right panel
 firstCm = lastCm + 7
 lastCm = firstCm + dissectorRange - 3
 cRow = 1
 crRow = 1
 Sheets(oldSheet).Select
 Set WorkingArea = Range(Cells(firstRw, firstCm), Cells(lastRw, lastCm))
 For Each Rw In WorkingArea.Rows    'iterate each row in dissector range
    For Each Cl In Rw.Cells     'iterate each cell in row
        cValue = Cl.Value        'get value of each cell
        If (IsNumeric(cValue)) And Len(cValue) > 0 Then
            cRow = cRow + 1
            crRow = crRow + 1
            Sheets(NewSheet).Select    'select the newly created worksheet
            Cells(cRow, 2).Value = cValue 'Copy the new coordinates in a new work sheet and clean up the current worksheet
            Cells(crRow, 4).Value = cValue 'Copy the new coordinates in a new work sheet and clean up the current worksheet
        Else
            If Cl.Interior.ColorIndex = 40 Then
                crRow = crRow + 1
                Cells(crRow, 4).Value = Val(Right(cValue, 1))
            End If
        End If
    Next Cl
 Next Rw
 
 'Dim Counter As Integer
              
 
              
    
End Sub
