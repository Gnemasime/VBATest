Sub ValidateTimesheet()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim errorsCount As Integer
    Dim employeeID As String
    Dim timesheetDate As Date
    Dim hoursWorked As Double
    Dim projectCode As String
    
    ' Initialize error count
    errorsCount = 0

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("Timesheet") ' Change the sheet name if needed

    ' Find the last row with data in column A (Employee ID column)
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Loop through each row starting from row 2 (assuming row 1 is headers)
    For i = 2 To lastRow
        ' Get the values from each column
        employeeID = ws.Cells(i, 1).Value        ' Employee ID
        timesheetDate = ws.Cells(i, 2).Value     ' Date
        hoursWorked = ws.Cells(i, 3).Value       ' Hours Worked
        projectCode = ws.Cells(i, 4).Value      ' Project Code

        ' Validate Employee ID (6 digits long)
        If Len(employeeID) <> 6 Or Not IsNumeric(employeeID) Then
            ws.Cells(i, 1).Interior.Color = RGB(255, 0, 0) ' Red color
            errorsCount = errorsCount + 1
        Else
            ws.Cells(i, 1).Interior.Color = RGB(255, 255, 255) ' Reset to white
        End If

        ' Validate Date (valid date and not a future date)
        If Not IsDate(timesheetDate) Or timesheetDate > Date Then
            ws.Cells(i, 2).Interior.Color = RGB(255, 0, 0) ' Red color
            errorsCount = errorsCount + 1
        Else
            ws.Cells(i, 2).Interior.Color = RGB(255, 255, 255) ' Reset to white
        End If

        ' Validate Hours Worked (positive number and not exceeding 12 hours)
        If Not IsNumeric(hoursWorked) Or hoursWorked <= 0 Or hoursWorked > 12 Then
            ws.Cells(i, 3).Interior.Color = RGB(255, 0, 0) ' Red color
            errorsCount = errorsCount + 1
        Else
            ws.Cells(i, 3).Interior.Color = RGB(255, 255, 255) ' Reset to white
        End If

        ' Validate Project Code (4-character alphanumeric code)
        If Len(projectCode) <> 4 Or Not projectCode Like "*[A-Za-z0-9]*" Then
            ws.Cells(i, 4).Interior.Color = RGB(255, 0, 0) ' Red color
            errorsCount = errorsCount + 1
        Else
            ws.Cells(i, 4).Interior.Color = RGB(255, 255, 255) ' Reset to white
        End If

    Next i

    ' Display a summary message
    If errorsCount = 0 Then
        MsgBox "Timesheet validated successfully.", vbInformation
    Else
        MsgBox "Validation completed with " & errorsCount & " error(s) found.", vbCritical
    End If

End Sub
