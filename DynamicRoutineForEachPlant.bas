Attribute VB_Name = "DynamicRoutineForEachPlant"
Sub GenerateERPRoutine()
    ' --- 1. SETTINGS & VARIABLES ---
    Dim wsSelected As Worksheet, wsRoutine As Worksheet, wsOutput As Worksheet, wsPlantVariables As Worksheet
    Dim tblRoutine As ListObject, tblOutput As ListObject, tblPlantFormats As ListObject
    Dim FinalProductList As Collection
    Dim product As Variant
    Dim routineRow As ListRow, destRow As ListRow
    Dim lastRow As Long, i As Long
    Dim selectedPlant As String, formatSheetName As String, outputTableName As String
    Dim headerCount As Long
    
    ' Speed Optimization: Turn off "lights" and calculation
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    application.EnableEvents = False

    ' Define worksheets
    Set wsSelected = ThisWorkbook.Sheets("2. Routines")
    Set wsPlantVariables = ThisWorkbook.Sheets("Plant Variables")
    Set wsOutput = ThisWorkbook.Sheets("6. Routine uploaders")

    ' Clear the "6. Routine uploaders" sheet fully
    wsOutput.Cells.Clear

    ' Get the selected plant
    selectedPlant = wsSelected.Range("D5").Text

    ' --- 2. FIND THE FORMAT SHEET ---
    On Error Resume Next
    Set tblPlantFormats = wsPlantVariables.ListObjects("PlantExportFormats")
    On Error GoTo 0

    If tblPlantFormats Is Nothing Then
        MsgBox "Table 'PlantExportFormats' not found.", vbCritical
        GoTo Cleanup ' Go to the end to turn settings back on
    End If

    ' Find the format sheet name (Optimized search logic)
    formatSheetName = ""
    Dim foundRow As Range
    ' Attempt to find the plant in the DataBodyRange of the table column
    On Error Resume Next
    formatSheetName = application.VLookup(selectedPlant, tblPlantFormats.DataBodyRange, _
                      tblPlantFormats.ListColumns("ERP Routing Format Sheet").Index, False)
    On Error GoTo 0

    If formatSheetName = "" Or formatSheetName = "Error 2042" Then
        MsgBox "No ERP export format sheet found for: " & selectedPlant, vbExclamation
        GoTo Cleanup
    End If

    ' Verify sheet exists
    On Error Resume Next
    Set wsRoutine = ThisWorkbook.Sheets(formatSheetName)
    On Error GoTo 0
    If wsRoutine Is Nothing Then
        MsgBox "Sheet '" & formatSheetName & "' does not exist.", vbCritical
        GoTo Cleanup
    End If

    ' --- 3. SETUP OUTPUT TABLE (DYNAMICALLY) ---
    outputTableName = "ERPRouting"
    Set tblRoutine = wsRoutine.ListObjects(1) ' The Source Table
    
    ' Count how many columns the Source has. This fixes your "New Column" bug.
    headerCount = tblRoutine.headerRowRange.Columns.Count

    ' Create output table sized exactly to the source columns
    ' We use .Resize to make the range match the source width automatically
    Set tblOutput = wsOutput.ListObjects.Add(xlSrcRange, wsOutput.Range("A1").Resize(1, headerCount), , xlYes)
    tblOutput.name = outputTableName

    ' Copy headers
    tblRoutine.headerRowRange.Copy
    wsOutput.Range("A1").PasteSpecial xlPasteValues
    application.CutCopyMode = False

    ' --- 4. GATHER DATA ---
    Set FinalProductList = New Collection
    lastRow = wsSelected.Cells(wsSelected.Rows.Count, 2).End(xlUp).row ' Check column 2 explicitly
    
    On Error Resume Next
    For i = 2 To lastRow
        Dim cellVal As Variant
        cellVal = wsSelected.Cells(i, 2).Text
        ' Basic check to ensure valid data
        If Len(cellVal) > 0 And cellVal <> "ERP Part Number" Then
            FinalProductList.Add cellVal, cellVal
        End If
    Next i
    On Error GoTo 0

    ' --- 5. THE BUILD LOOP (OPTIMIZED) ---
    
    ' If there are no products, stop here
    If FinalProductList.Count = 0 Then
        MsgBox "No products found to process.", vbExclamation
        GoTo Cleanup
    End If

    For Each product In FinalProductList
        For Each routineRow In tblRoutine.ListRows
            ' Add a new row
            Set destRow = tblOutput.ListRows.Add
            
            ' Set the Product Name (Column 1)
            destRow.Range(1, 1).value = product
            
            ' OPTIMIZATION: Copy ALL formulas in the row at once!
            ' We skip column 1 (Product) and copy from Col 2 to the end.
            ' This is ONE action per row, instead of 20+ loops per row.
            destRow.Range(1, 2).Resize(1, headerCount - 1).Formula = _
                routineRow.Range(1, 2).Resize(1, headerCount - 1).Formula
        Next routineRow
    Next product


    ' --- 6. APPLY PLANT 14 LOGIC (Post-Processing) ---
    ' Check if the selected plant string starts with "14"
    If Left(selectedPlant, 2) = "14" Then
        On Error Resume Next
        ' Attempt to reference the "Plnt" column
        Dim plntCol As ListColumn
        Set plntCol = tblOutput.ListColumns("Plnt")
        
        If Not plntCol Is Nothing Then
            ' If the column exists, overwrite the entire column with the plant code
            ' This is instant, even for 10,000 rows
            plntCol.DataBodyRange.value = selectedPlant
        Else
            ' Optional: Notify if the column is missing
            MsgBox "Note: You selected a '14' plant, but the output table has no 'Plnt' column.", vbExclamation
        End If
        On Error GoTo 0
    End If

    wsOutput.Activate
    MsgBox "ERP Routing generated successfully!", vbInformation

Cleanup:
    ' Restore settings so Excel works normally again
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    application.EnableEvents = True

End Sub

