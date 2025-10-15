Attribute VB_Name = "UpdateRoutinesBySelectedPlant"
Sub UpdateSelectedRoutines()
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim routineTable As ListObject
    Dim selectedTable As ListObject
    Dim destRow As ListRow
    Dim sourceRow As ListRow
    Dim workCenterCode As String
    Dim operationName As String
    Dim trValue As Variant
    Dim teValue As Variant
    Dim missingOperations As String
    Dim matchFound As Boolean
    Dim missingCount As Integer

    ' Set references to the sheets
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets("2. Routines")
    Set wsDestination = ThisWorkbook.Sheets("SelectedRoutines")
    Set routineTable = wsSource.ListObjects("RoutinesDB")
    Set selectedTable = wsDestination.ListObjects("SelectedRoutines")
    On Error GoTo 0

    ' Verify that the tables and worksheets exist
    If wsSource Is Nothing Or wsDestination Is Nothing Then
        MsgBox "Error: One of the worksheets does not exist.", vbCritical
        Exit Sub
    End If

    If routineTable Is Nothing Then
        MsgBox "Error: RoutinesDB table not found on the specified sheet.", vbCritical
        Exit Sub
    End If

    If selectedTable Is Nothing Then
        MsgBox "Error: SelectedRoutines table not found on the specified sheet.", vbCritical
        Exit Sub
    End If

    ' Initialize the missing operations message and missing count
    missingOperations = "No matching records found for the following items:" & vbCrLf
    missingCount = 0

    ' Loop through each row in SelectedRoutines
    For Each destRow In selectedTable.ListRows
        matchFound = False ' Reset matchFound for each destRow
        
        ' Capture the original values in SelectedRoutines for error reporting
        Dim originalPlant As String
        Dim originalProductType As String
        
        originalPlant = Trim(destRow.Range.Cells(1, selectedTable.ListColumns("Plant").Index).value)
        originalProductType = Trim(destRow.Range.Cells(1, selectedTable.ListColumns("Product Type").Index).value)
        workCenterCode = Trim(destRow.Range.Cells(1, selectedTable.ListColumns("Macrophase").Index).value)
        operationName = Trim(destRow.Range.Cells(1, selectedTable.ListColumns("Microphase").Index).value)

        ' Skip row if essential values are empty
        If originalPlant <> "" And originalProductType <> "" And workCenterCode <> "" And operationName <> "" Then
            
            ' Find the corresponding row in Routines for the selected plant and operation
            For Each sourceRow In routineTable.ListRows
                Dim srcPlant As String
                Dim srcProductType As String
                Dim srcWorkCenterCode As String
                Dim srcOperationName As String

                srcPlant = Trim(sourceRow.Range.Cells(1, routineTable.ListColumns("Plant").Index).value)
                srcWorkCenterCode = Trim(sourceRow.Range.Cells(1, routineTable.ListColumns("Macrophase").Index).value)
                srcOperationName = Trim(sourceRow.Range.Cells(1, routineTable.ListColumns("Microphase").Index).value)

                ' Check for match
                If srcPlant = originalPlant And _
                   srcProductType = originalProductType And _
                   srcWorkCenterCode = workCenterCode And _
                   srcOperationName = operationName Then

                    ' Retrieve tr and te values
                    trValue = sourceRow.Range.Cells(1, routineTable.ListColumns("tr").Index).value
                    teValue = sourceRow.Range.Cells(1, routineTable.ListColumns("te").Index).value

                    ' Update SelectedRoutines
                    destRow.Range.Cells(1, selectedTable.ListColumns("tr").Index).value = trValue
                    destRow.Range.Cells(1, selectedTable.ListColumns("te").Index).value = teValue

                    ' Also update Plant and Product Type in the destination table
                    destRow.Range.Cells(1, selectedTable.ListColumns("Plant").Index).value = srcPlant
                    destRow.Range.Cells(1, selectedTable.ListColumns("Product Type").Index).value = srcProductType

                    matchFound = True ' Match found, exit inner loop
                    Exit For
                End If
            Next sourceRow

            ' If no match was found, add to missingOperations
            If Not matchFound Then
                missingOperations = missingOperations & "Plant: " & originalPlant & _
                                    ", Product Type: " & originalProductType & _
                                    ", Work Center Code: " & workCenterCode & _
                                    ", Operation: " & operationName & vbCrLf
                missingCount = missingCount + 1
            End If
        End If
    Next destRow

    ' Display missing operations and open RoutineForm if necessary
    If missingCount > 0 Then
        MsgBox missingOperations, vbExclamation
        RoutineForm.Show  ' Open the form for user selection if some operations weren't found
    Else
        ' MsgBox "All operations updated successfully.", vbInformation
    End If
End Sub


