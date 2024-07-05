Sub btnSelectFile_Click()
    Dim sourceFilePath As Variant
    sourceFilePath = Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx*), *.xlsx*", Title:="Select a File")
    If sourceFilePath <> False Then
        Dim originalFileName As String
        Dim destFolderPath As String
        Dim wbSource As Workbook
        Dim wsAccessControl As Worksheet
        Dim emptyrows As Integer
        Dim wsToDelete As Worksheet
    
        ' Extract the original filename from the source file path
        originalFileName = Mid(sourceFilePath, InStrRev(sourceFilePath, "\") + 1)
        ' Set the destination folder path to the folder containing the source file
        destFolderPath = Left(sourceFilePath, InStrRev(sourceFilePath, "\"))
        
        ' Open the source workbook
        Set wbSource = Workbooks.Open(sourceFilePath)
        
        ' Specify the worksheet to delete
        Set wsToDelete = wbSource.Sheets("ValidationErrorSummary")
        ' Delete the worksheet
        Application.DisplayAlerts = False ' Suppress alerts to confirm deletion
        wsToDelete.Delete
        Application.DisplayAlerts = True
        
        ' Set the "AccessControl" worksheet
        Set wsAccessControl = wbSource.Sheets("AccessControl")
        ' Insert empty rows from row 2 to row 6
        For emptyrows = 1 To 5
            wsAccessControl.Rows(2).Insert Shift:=xlDown
        Next emptyrows
          
        ' there comes the hard part...
        Dim maxRowsinsheets As Long
        Dim sourceWorksheet As Worksheet
        Dim wsCount As Integer
        Dim filesNeeded As Long
        Dim fileCounter As Long
        Dim newWB As Workbook
        Dim newWS As Worksheet
        Dim i As Long, j As Long
        Dim ws2ToDelete As Worksheet
        ' Initialize variables for copying data rows
        Dim copyRange As Range
        Dim startRow As Long
        Dim endRow As Long
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        
        ' Determine the index based on the sheet with the most rows
        maxRowsinsheets = 0
        For Each sourceWorksheet In wbSource.Worksheets
            If sourceWorksheet.Cells(sourceWorksheet.Rows.Count, 1).End(xlUp).Row > maxRowsinsheets Then
                maxRowsinsheets = sourceWorksheet.Cells(sourceWorksheet.Rows.Count, 1).End(xlUp).Row
            End If
        Next sourceWorksheet
        
        ' Calculate the number of output files needed
        filesNeeded = WorksheetFunction.Ceiling((maxRowsinsheets - 7) / 350, 1)
        ' Debug.Print filesNeeded
        
      
        
        ' copying 350 rows in chunks into new files
        ' Loop through each output file needed
        For fileCounter = 1 To filesNeeded
            ' Create a new workbook for output
            Set newWB = Workbooks.Add
            
            ' Copy each worksheet from source to output workbook (in reverse order, that's why we use "Step -1")
            wsCount = wbSource.Worksheets.Count
            For j = wsCount To 1 Step -1
                Set sourceWorksheet = wbSource.Worksheets(j)
                ' Add a new worksheet in the output workbook
                Set newWS = newWB.Worksheets.Add
                ' Rename the new worksheet
                newWS.Name = sourceWorksheet.Name
                
                ' Copy the header rows to the new worksheet
                For i = 1 To 7
                    sourceWorksheet.Rows(i).Copy newWS.Rows(i)
                Next i
        
                ' Determine the start and end rows for copying data
                startRow = 8 + (fileCounter - 1) * 350
                endRow = Application.WorksheetFunction.Min(startRow + 349, sourceWorksheet.Cells(sourceWorksheet.Rows.Count, 1).End(xlUp).Row)
        
                ' Copy data rows to the new worksheet
                If startRow <= endRow Then
                    sourceWorksheet.Rows(startRow & ":" & endRow).Copy newWS.Rows(8)
                End If
            Next j
            
            ' Specify the worksheet to delete
            Set ws2ToDelete = newWB.Sheets("Sheet1")
            ' Delete the worksheet
            Application.DisplayAlerts = False ' Suppress alerts to confirm deletion
            ws2ToDelete.Delete
            Application.DisplayAlerts = True
            
            ' Save the output workbook
            newWB.SaveAs destFolderPath & "Edited_" & originalFileName & "_Split_" & fileCounter & ".xlsx", FileFormat:=xlOpenXMLWorkbook
            newWB.Close SaveChanges:=False
        Next fileCounter
    
        
        ' Close the source workbook
        wbSource.Close SaveChanges:=False
        
        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
        
        MsgBox "Files edited and saved successfully!"
    End If
End Sub
