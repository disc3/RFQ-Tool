Attribute VB_Name = "DynamicRoutineForEachPlant"
Sub GenerateERPRoutine()
    Dim wsSelected As Worksheet
    Dim wsRoutine As Worksheet
    Dim wsOutput As Worksheet
    Dim wsPlantVariables As Worksheet
    Dim tblRoutine As ListObject
    Dim tblOutput As ListObject
    Dim tblPlantFormats As ListObject
    Dim FinalProductList As Collection
    Dim product As Variant
    Dim routineRow As ListRow
    Dim destRow As ListRow
    Dim lastRow As Long
    Dim i As Long
    Dim selectedPlant As String
    Dim formatSheetName As String
    Dim outputTableName As String
    Dim headerCount As Long

    ' Define worksheets
    Set wsSelected = ThisWorkbook.Sheets("2. Routines")
    Set wsPlantVariables = ThisWorkbook.Sheets("Plant Variables")
    Set wsOutput = ThisWorkbook.Sheets("6. Routine uploaders")

    ' Clear the "6. Routine uploaders" sheet
    wsOutput.Cells.Clear

    ' Get the selected plant from "2. Routines" sheet
    selectedPlant = wsSelected.Range("D5").value

    ' Retrieve the appropriate format sheet name from "PlantExportFormats" table
    On Error Resume Next
    Set tblPlantFormats = wsPlantVariables.ListObjects("PlantExportFormats")
    On Error GoTo 0

    If tblPlantFormats Is Nothing Then
        MsgBox "Table 'PlantExportFormats' not found in 'Plant Variables' sheet.", vbCritical
        Exit Sub
    End If

    ' Find the format sheet name for the selected plant
    formatSheetName = ""
    For Each plantRow In tblPlantFormats.ListRows
        If plantRow.Range(tblPlantFormats.ListColumns("Selected Plant").Index).value = selectedPlant Then
            formatSheetName = plantRow.Range(tblPlantFormats.ListColumns("ERP Routing Format Sheet").Index).value
            Exit For
        End If
    Next plantRow

    ' Check if the format sheet name was found
    If formatSheetName = "" Then
        MsgBox "No ERP export format sheet found for the selected plant: " & selectedPlant, vbExclamation
        Exit Sub
    End If

    ' Define the table name as "ERPRouting{Plant}"
    outputTableName = "ERPRouting"

    ' Set wsRoutine to the format sheet found
    On Error Resume Next
    Set wsRoutine = ThisWorkbook.Sheets(formatSheetName)
    On Error GoTo 0

    If wsRoutine Is Nothing Then
        MsgBox "The ERP Routing Format sheet '" & formatSheetName & "' does not exist.", vbCritical
        Exit Sub
    End If

    ' Create the ERP Routing table with the name "ERPRouting{Plant}" if it doesn't already exist
    On Error Resume Next
    Set tblOutput = wsOutput.ListObjects(outputTableName)
    On Error GoTo 0
    If tblOutput Is Nothing Then
        Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, wsOutput.Range("A1:R1"), , xlYes)
        tblOutput.name = outputTableName
    End If

    ' Get the Routine table on the format sheet and its headers
    Set tblRoutine = wsRoutine.ListObjects(1) ' Assuming there is only one table on the format sheet

    ' Set headers in ERP Routine table from the format table
    tblRoutine.HeaderRowRange.Copy
    wsOutput.Range("A1").PasteSpecial xlPasteValues
    application.CutCopyMode = False

    ' Collect unique products from SelectedRoutines table
    Set FinalProductList = New Collection
    lastRow = wsSelected.Cells(wsSelected.Rows.Count, 1).End(xlUp).row
    On Error Resume Next
    For i = 2 To lastRow ' Start from row 2 to skip the header
        If Trim(wsSelected.Cells(i, 2).value) <> "" And wsSelected.Cells(i, 2).value <> "ERP Part Number" Then
            FinalProductList.Add wsSelected.Cells(i, 2).value, CStr(wsSelected.Cells(i, 2).value) ' Assuming ERP Part Number is in column 2
        End If
    Next i
    On Error GoTo 0

    ' Copy the structure for each unique product
    headerCount = tblRoutine.HeaderRowRange.Columns.Count

    For Each product In FinalProductList
        ' Loop through each row in the specified format table
        For Each routineRow In tblRoutine.ListRows
            ' Add a new row in the ERP Routine table
            Set destRow = tblOutput.ListRows.Add
            
            ' Populate the Product column and copy the rest from the format table
            destRow.Range(1).value = product ' Assuming Product is the first column
            For i = 2 To headerCount ' Start from the second column to copy remaining formulas
                destRow.Range(i).Formula = routineRow.Range(i).Formula ' Copy formulas instead of values
            Next i
        Next routineRow
    Next product

    'MsgBox "ERP Routing generated successfully for " & selectedPlant & "!", vbInformation
    'wsOutput.Activate
End Sub


