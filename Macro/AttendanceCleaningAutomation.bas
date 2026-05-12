Sub Employee_Attendance_Cleaning()

    'Set active worksheet
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    'Insert Header Row
    ws.Rows(1).Insert
    ws.Cells(1, 1).Value = "Employee Code"
    ws.Cells(1, 2).Value = "Employee Name"
    ws.Cells(1, 3).Value = "Date Present"
    
    'Find Last Row
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    'Find Columns by Header Name
    Dim codeCol As Long
    Dim nameCol As Long
    Dim dateCol As Long
    codeCol = ws.Rows(1).Find("Employee Code", LookAt:=xlWhole).Column
    nameCol = ws.Rows(1).Find("Employee Name", LookAt:=xlWhole).Column
    dateCol = ws.Rows(1).Find("Date Present", LookAt:=xlWhole).Column
    
    'Validate Columns Were Found
    If codeCol = 0 Then
        MsgBox "Column 'Employee Code' not found. Check headers and re-run."
        Exit Sub
    End If
    If nameCol = 0 Then
        MsgBox "Column 'Employee Name' not found. Check headers and re-run."
        Exit Sub
    End If
    If dateCol = 0 Then
        MsgBox "Column 'Date Present' not found. Check headers and re-run."
        Exit Sub
    End If
    
    'Loop Through All Rows — Strip Time, Trim Spaces
    Dim i As Long
    For i = 2 To lastRow
    
        'Strip Time from DateTime — Keep Date Only
        ws.Cells(i, dateCol).Value = Int(ws.Cells(i, dateCol).Value)
        
        'Trim Leading, Trailing and Middle Spaces from Employee Code
        ws.Cells(i, codeCol).Value = Application.WorksheetFunction.Trim(ws.Cells(i, codeCol).Value)
        
        'Trim Leading, Trailing and Middle Spaces from Employee Name
        ws.Cells(i, nameCol).Value = Application.WorksheetFunction.Trim(ws.Cells(i, nameCol).Value)
        
    Next i
    
    'Format Date Column as DD/MM/YYYY
    ws.Range(ws.Cells(2, dateCol), ws.Cells(lastRow, dateCol)).NumberFormat = "DD/MM/YYYY"
    
    'Confirm Completion
    MsgBox "Cleaning complete. " & lastRow - 1 & " records processed."

End Sub
