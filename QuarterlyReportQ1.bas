'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'This was built in about 3 days with no time to optimize or clean up the code. Note that the other reports for
'Q2, Q3, and Q4 do literally the exact same thing but for different quarters of the year. This could be cleaned up
'a bunch but it doesn't seem like the best use of my time right now. Maybe someday
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Sub get_q1()

    Dim tab_interest As Worksheet
    Dim RowCount As Integer: RowCount = 0
    Dim rowLength As Long, j As Long
    Dim cellDate As Date
    Dim dateString As String
    Dim writeRow As Integer: writeRow = 1
    Dim totalDays As Integer
    Dim totalRows As Integer
    Dim avgDays As Double
    Dim my_resheet As Worksheet
    Dim WB_Orig As Workbook
    Dim myValue As Variant
    
    myValue = InputBox("What year would you like to run a report for?")
    
    Dim dt As Date: dt = "07/01/" & myValue
    Dim dt2 As Date: dt2 = "09/30/" & myValue
    
    'WB_Orig is the excel file that is running the macro, the one with our graphs and table
    Set WB_Orig = ThisWorkbook
    'overview is our Overview tab on our macro spreadsheet
    Set overview = WB_Orig.Worksheets("Overview")
    
    'prints at the top of the table what quarter and year we're running a report for
    overview.Cells(1, 1).Value = "Quarterly Report for Q1: " & dt & " to " & dt2
    
    'clear out our tabs, they may have data in them from previous reports
    Sheets("Stipends").UsedRange.ClearContents
    Sheets("Equities").UsedRange.ClearContents
    Sheets("Updated JD").UsedRange.ClearContents
    Sheets("Reclasses").UsedRange.ClearContents
    Sheets("Star").UsedRange.ClearContents
    Sheets("1Time").UsedRange.ClearContents
    
    '''''''''''''''''''''''''''''RECLASS'''''''''''''''''''''''''''''''
    
    'Open the Completed tab in the Reclass spreadsheet
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Reclass").Worksheets("Completed")
    'Find out how many rows we have
    rowLength = tab_interest.UsedRange.Rows.Count
    Set my_resheet = WB_Orig.Worksheets("Reclasses")
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
        
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 23).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 23).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'copy the entire row if the date is in our quarter
              tab_interest.Rows(j).EntireRow.Copy my_resheet.Rows(writeRow)
              'add 1 to writeRow so that we dont overwrite the row if we find something else
              writeRow = writeRow + 1
            End If
        End If
        
    Next j
    
    'Count "In Process" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Reclass").Worksheets("In Process")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    Dim process_count As Integer: process_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 23).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 23).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'we found something within our quarter, add one to our count
              process_count = process_count + 1
            End If
        End If
        
    Next j
    
    'Count "Withdrawn/On Hold" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Reclass").Worksheets("Withdrawn or On Hold")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    Dim withdrawn_count As Integer: withdrawn_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 23).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 23).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              withdrawn_count = withdrawn_count + 1
            End If
        End If
        
    Next j
    
    
    'Close the file, we dont want to keep it open
    Workbooks("Reclass").Close
    'find total days for completion
    totalDays = Application.WorksheetFunction.Sum(my_resheet.Columns("Y:Y"))
    'get total rows pulled
    totalRows = my_resheet.Cells(my_resheet.Rows.Count, "A").End(xlUp).Row
    'calculate average days to completion
    avgDays = totalDays / totalRows

    
    'Print total "in process" in our main sheet
    overview.Cells(6, 2).Value = process_count
    'Print total "complete" in our main sheet
    overview.Cells(6, 3).Value = totalRows
    'Print total "withdrawn" items in our main sheet
    overview.Cells(6, 4).Value = withdrawn_count
    'Print avg time to complete in our main sheet
    overview.Cells(6, 8).Value = avgDays
    
    '''''''''''''''''''''''''''''EQUITIES'''''''''''''''''''''''''''''''
    
    'Open the Completed tab in the Equities spreadsheet
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Equities").Worksheets("Completed")
    'Find out how many rows we have
    rowLength = tab_interest.UsedRange.Rows.Count
    Dim my_esheet As Worksheet
    
    Set my_esheet = WB_Orig.Worksheets("Equities")
    
    writeRow = 1
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
        
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 19).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 19).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'copy the entire row if the date is in our quarter
              tab_interest.Rows(j).EntireRow.Copy ThisWorkbook.Worksheets("Equities").Rows(writeRow)
              writeRow = writeRow + 1
            End If
        End If
        
    Next j
    
    'Count "In Process" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Equities").Worksheets("In Process")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    process_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 19).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 19).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              process_count = process_count + 1
            End If
        End If
        
    Next j
    
    'Count "Withdrawn/On Hold" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Equities").Worksheets("Withdrawn or On Hold")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    withdrawn_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 19).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 19).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              withdrawn_count = withdrawn_count + 1
            End If
        End If
        
    Next j
    
    
    
    Workbooks("Equities").Close
    'find total days for completion
    totalDays = Application.WorksheetFunction.Sum(my_esheet.Columns("U:U"))
    'get total rows pulled
    totalRows = my_esheet.Cells(Rows.Count, "A").End(xlUp).Row
    'calculate average days to completion
    avgDays = totalDays / totalRows

    
    
    'Print total "in process" in our main sheet
    overview.Cells(5, 2).Value = process_count
    'Print total "complete" in our main sheet
    overview.Cells(5, 3).Value = totalRows
    'Print total "withdrawn" items in our main sheet
    overview.Cells(5, 4).Value = withdrawn_count
    'Print avg time to complete in our main sheet
    overview.Cells(5, 8).Value = avgDays
    
    
    '''''''''''''''''''''''''''''Stipends'''''''''''''''''''''''''''''''
    
    'Open the Completed tab in the Equities spreadsheet
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Stipends").Worksheets("Completed")
    'Find out how many rows we have
    rowLength = tab_interest.UsedRange.Rows.Count
    Dim my_ssheet As Worksheet
    
    Set my_ssheet = WB_Orig.Worksheets("Stipends")
    
    writeRow = 1
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
        
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 23).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 23).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'copy the entire row if the date is in our quarter
              tab_interest.Rows(j).EntireRow.Copy ThisWorkbook.Worksheets("Stipends").Rows(writeRow)
              writeRow = writeRow + 1
            End If
        End If
        
    Next j
    
    'Count "In Process" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Stipends").Worksheets("In Process")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    process_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 23).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 23).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              process_count = process_count + 1
            End If
        End If
        
    Next j
    
    'Count "Withdrawn/On Hold" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Stipends").Worksheets("Withdrawn or On Hold")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    withdrawn_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 23).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 23).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              withdrawn_count = withdrawn_count + 1
            End If
        End If
        
    Next j
    
    
    
    Workbooks("Stipends").Close
    'find total days for completion
    totalDays = Application.WorksheetFunction.Sum(my_ssheet.Columns("U:U"))
    'get total rows pulled
    totalRows = my_ssheet.Cells(Rows.Count, "A").End(xlUp).Row
    'calculate average days to completion
    avgDays = totalDays / totalRows

    
    
    'Print total "in process" in our main sheet
    overview.Cells(4, 2).Value = process_count
    'Print total "complete" in our main sheet
    overview.Cells(4, 3).Value = totalRows
    'Print total "withdrawn" items in our main sheet
    overview.Cells(4, 4).Value = withdrawn_count
    'Print avg time to complete in our main sheet
    overview.Cells(4, 8).Value = avgDays
    

    '''''''''''''''''''''''''''''STAR Awards'''''''''''''''''''''''''''''''
    'Open the Completed tab in the Equities spreadsheet
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "STAR Awards").Worksheets("Completed")
    'Find out how many rows we have
    rowLength = tab_interest.UsedRange.Rows.Count
    Dim my_star_sheet As Worksheet
    
    Set my_star_sheet = WB_Orig.Worksheets("Star")
    
    writeRow = 1
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
        
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 24).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 24).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'copy the entire row if the date is in our quarter
              tab_interest.Rows(j).EntireRow.Copy ThisWorkbook.Worksheets("Star").Rows(writeRow)
              writeRow = writeRow + 1
            End If
        End If
        
    Next j
    
    'Count "In Process" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "STAR Awards").Worksheets("In Process")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    process_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 24).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 24).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              process_count = process_count + 1
            End If
        End If
        
    Next j
    
    'Count "Withdrawn/On Hold" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "STAR Awards").Worksheets("Withdrawn or On Hold")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    withdrawn_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 24).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 24).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              withdrawn_count = withdrawn_count + 1
            End If
        End If
        
    Next j
    
    
    
    Workbooks("STAR Awards").Close
    'find total days for completion
    totalDays = Application.WorksheetFunction.Sum(my_star_sheet.Columns("Z:Z"))
    'get total rows pulled
    totalRows = my_star_sheet.Cells(Rows.Count, "A").End(xlUp).Row
    'calculate average days to completion
    avgDays = totalDays / totalRows

    
    
    'Print total "in process" in our main sheet
    overview.Cells(7, 2).Value = process_count
    'Print total "complete" in our main sheet
    overview.Cells(7, 3).Value = totalRows
    'Print total "withdrawn" items in our main sheet
    overview.Cells(7, 4).Value = withdrawn_count
    'Print avg time to complete in our main sheet
    overview.Cells(7, 8).Value = avgDays
    
    
    '''''''''''''''''''''''''''''JD Updates'''''''''''''''''''''''''''''''
    'Open the Completed tab in the Equities spreadsheet
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Updated JD").Worksheets("Completed")
    'Find out how many rows we have
    rowLength = tab_interest.UsedRange.Rows.Count
    Dim my_jd_sheet As Worksheet
    
    Set my_jd_sheet = WB_Orig.Worksheets("Updated JD")
    
    writeRow = 1
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
        
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 13).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 13).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'copy the entire row if the date is in our quarter
              tab_interest.Rows(j).EntireRow.Copy ThisWorkbook.Worksheets("Updated JD").Rows(writeRow)
              writeRow = writeRow + 1
            End If
        End If
        
    Next j
    
    'Count "In Process" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Updated JD").Worksheets("In Process")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    process_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 13).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 13).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              process_count = process_count + 1
            End If
        End If
        
    Next j
    
    'Count "Withdrawn/On Hold" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "Updated JD").Worksheets("Withdrawn or On Hold")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    withdrawn_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 13).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 13).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              withdrawn_count = withdrawn_count + 1
            End If
        End If
        
    Next j
    
    
    
    Workbooks("Updated JD").Close
    'find total days for completion
    totalDays = Application.WorksheetFunction.Sum(my_jd_sheet.Columns("O:O"))
    'get total rows pulled
    totalRows = my_jd_sheet.Cells(Rows.Count, "A").End(xlUp).Row
    'calculate average days to completion
    avgDays = totalDays / totalRows

    
    
    'Print total "in process" in our main sheet
    overview.Cells(9, 2).Value = process_count
    'Print total "complete" in our main sheet
    overview.Cells(9, 3).Value = totalRows
    'Print total "withdrawn" items in our main sheet
    overview.Cells(9, 4).Value = withdrawn_count
    'Print avg time to complete in our main sheet
    overview.Cells(9, 8).Value = avgDays
    
    
    '''''''''''''''''''''''''''''1 Time Payments'''''''''''''''''''''''''''''''
    'Open the Completed tab in the Equities spreadsheet
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "One-Time Payments").Worksheets("(C) 1Time Payments")
    'Find out how many rows we have
    rowLength = tab_interest.UsedRange.Rows.Count
    Dim my_1_sheet As Worksheet
    
    Set my_1_sheet = WB_Orig.Worksheets("1Time")
    
    writeRow = 1
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
        
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 17).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 17).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'copy the entire row if the date is in our quarter
              
              tab_interest.Range("Q" & CStr(j) & ":S" & CStr(j)).Copy ThisWorkbook.Worksheets("1Time").Rows(writeRow)
              writeRow = writeRow + 1
            End If
        End If
        
    Next j
    
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "One-Time Payments").Worksheets("(C) INTER&MULT 1TIMEPAY")
    'Find out how many rows we have
    rowLength = tab_interest.UsedRange.Rows.Count
    For j = 3 To rowLength
        
        'make sure the column isnt blank
        If ((tab_interest.Cells(j, 29).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 29).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'copy the entire row if the date is in our quarter

               tab_interest.Range("AC" & CStr(j) & ":AE" & CStr(j)).Copy ThisWorkbook.Worksheets("1Time").Rows(writeRow)
              writeRow = writeRow + 1
            End If
        End If
        
    Next j
    
    'Count "In Process" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "One-Time Payments").Worksheets("(IP) 1TIME PAYMENTS")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    process_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 17).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 17).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              process_count = process_count + 1
            End If
        End If
        
    Next j
    
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "One-Time Payments").Worksheets("(IP) 1TIME PAYMENTS")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 29).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 29).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              process_count = process_count + 1
            End If
        End If
        
    Next j
    
    'Count "Withdrawn/On Hold" items
    Set tab_interest = Workbooks.Open(ThisWorkbook.Path & "\" & "One-Time Payments").Worksheets("Withdrawn or On Hold")
    'find how many rows we have in the sheet
    rowLength = tab_interest.UsedRange.Rows.Count
    'set this variable to 0, since we havent found anything yet (and might not)
    withdrawn_count = 0
    
    'start at row 3, since thats where the information begins
    For j = 3 To rowLength
    
        If ((tab_interest.Cells(j, 13).Value) <> "") Then
            'take the cell in the column "Date Recieved"
            dateString = tab_interest.Cells(j, 13).Value
            'this converts what we got from the cell into a "Date" that we can compare to other dates
            cellDate = CDate(dateString)
            
            'check that the date is in our quarter (between two dates)
            If (cellDate >= dt And cellDate <= dt2) Then
              'increment count
              withdrawn_count = withdrawn_count + 1
            End If
        End If
        
    Next j
    
    
    
    Workbooks("One-Time Payments").Close
    'find total days for completion
    totalDays = Application.WorksheetFunction.Sum(my_1_sheet.Columns("C:C"))
    'get total rows pulled
    totalRows = my_1_sheet.Cells(Rows.Count, "A").End(xlUp).Row
    'calculate average days to completion
    avgDays = totalDays / totalRows

    
    
    'Print total "in process" in our main sheet
    overview.Cells(8, 2).Value = process_count
    'Print total "complete" in our main sheet
    overview.Cells(8, 3).Value = totalRows
    'Print total "withdrawn" items in our main sheet
    overview.Cells(8, 4).Value = withdrawn_count
    'Print avg time to complete in our main sheet
    overview.Cells(8, 8).Value = avgDays
    
    

End Sub
