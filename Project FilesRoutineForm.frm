VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RoutineForm 
   Caption         =   "Add Manufacturing Operations"
   ClientHeight    =   6105
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   17460
   OleObjectBlob   =   "Project FilesRoutineForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "RoutineForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public workCenterCodeTemp As String ' Global variable to store the Work Center Code

Private mPreselectedComponent As String ' Private variable to store the preselected component


' Property to allow setting preselectedComponent from outside
Public Property Let PreselectedComponent(value As String)
    mPreselectedComponent = value
End Property
Public Sub SetupForm()
    Dim wsMain As Worksheet
    Dim wsRoutines As Worksheet
    Dim wsComponents As Worksheet
    Dim wsSelectedRoutines As Worksheet
    Dim tblRoutinesDB As ListObject
    Dim tblComponents As ListObject
    Dim componentRow As ListRow
    Dim routineRow As ListRow
    Dim selectedPlant As String
    Dim selectedProduct As String
    Dim teValue As Variant
    Dim trValue As Variant
    Dim uniqueMacrophases As Object
    Dim sortedMacrophases() As Variant
    Dim rowIdx As Long
    Dim sortOrder As Double
    Dim macrophaseName As Variant ' Changed to Variant for Dictionary iteration

    ' Set worksheet references
    Set wsMain = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsRoutines = ThisWorkbook.Sheets("RoutinesDB")
    Set wsSelectedRoutines = ThisWorkbook.Sheets("2. Routines")
    Set wsComponents = wsMain

    ' Validate table existence and set table references
    On Error Resume Next
    Set tblRoutinesDB = wsRoutines.ListObjects("RoutinesDB")
    Set tblComponents = wsComponents.ListObjects("BOMDefinition")
    On Error GoTo 0

    If tblRoutinesDB Is Nothing Or tblComponents Is Nothing Then Exit Sub

    ' Get the selected plant and product from the Product Specification sheet
    selectedPlant = Trim(wsMain.Range("C9").value)
    selectedProduct = Trim(wsSelectedRoutines.Range("D6").value)
    Me.lblselectedProduct.Caption = "Selected Final Product: " & selectedProduct

    ' Use a Dictionary to store unique Macrophases and their Sort Order
    Set uniqueMacrophases = CreateObject("Scripting.Dictionary")

    ' Collect unique Macrophases with their Sort Order for the selected plant
    For Each routineRow In tblRoutinesDB.ListRows
        If Trim(routineRow.Range(tblRoutinesDB.ListColumns("Plant").Index).value) = selectedPlant Then
            teValue = routineRow.Range(tblRoutinesDB.ListColumns("te").Index).value
            trValue = routineRow.Range(tblRoutinesDB.ListColumns("tr").Index).value
            macrophaseName = Trim(routineRow.Range(tblRoutinesDB.ListColumns("Macrophase").Index).value)
            sortOrder = routineRow.Range(tblRoutinesDB.ListColumns("Sort Order").Index).value

            ' Add unique Macrophases with Sort Order to the Dictionary
            If Not (IsEmpty(teValue) And IsEmpty(trValue)) Then
                If Not uniqueMacrophases.exists(macrophaseName) Then
                    uniqueMacrophases.Add macrophaseName, sortOrder
                End If
            End If
        End If
    Next routineRow

    ' Transfer unique Macrophases from the Dictionary to an array for sorting
    rowIdx = 0
    ReDim sortedMacrophases(0 To uniqueMacrophases.Count - 1)
    For Each macrophaseName In uniqueMacrophases.Keys
        sortedMacrophases(rowIdx) = Array(uniqueMacrophases(macrophaseName), macrophaseName)
        rowIdx = rowIdx + 1
    Next macrophaseName

    ' Sort the array by Sort Order
    Call BubbleSortBySortOrder(sortedMacrophases)

    ' Populate the MacrophaseSelect ComboBox with sorted unique Macrophases
    Me.MacrophaseSelect.Clear
    For rowIdx = LBound(sortedMacrophases) To UBound(sortedMacrophases)
        Me.MacrophaseSelect.AddItem sortedMacrophases(rowIdx)(1) ' Add the Macrophase name
    Next rowIdx

    ' Populate the Component ComboBox with components that match the selected product
    Me.cmbComponentSelect.Clear
    For Each componentRow In tblComponents.ListRows
        If componentRow.Range(tblComponents.ListColumns("Product Number").Index).value = selectedProduct Then
            Me.cmbComponentSelect.AddItem componentRow.Range(tblComponents.ListColumns("Material").Index).value
        End If
    Next componentRow

    ' Preselect the component only if mPreselectedComponent has a value
    If mPreselectedComponent <> "" Then
        Dim i As Long, found As Boolean
        found = False
        For i = 0 To Me.cmbComponentSelect.ListCount - 1
            If Me.cmbComponentSelect.List(i) = mPreselectedComponent Then
                Me.cmbComponentSelect.ListIndex = i
                found = True
                Exit For
            End If
        Next i
    End If

    ' Set up the ListView headers in the specified order
    With Me.OperationsListView
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .LabelEdit = lvwManual

        ' Add columns in the specified order
        .columnHeaders.Add , , "Operation", 110
        .columnHeaders.Add , , "Material", 100
        .columnHeaders.Add , , "Machine", 100
        .columnHeaders.Add , , "Wire/Cable Diameter", 65
        .columnHeaders.Add , , "Wire/Component Dimension", 60
        .columnHeaders.Add , , "Setup [sec]", 60
        .columnHeaders.Add , , "Manufacturing [sec]", 60
        .columnHeaders.Add , , "Number of Operations", 75
        ' Hidden Sort Order column (width 0)
        .columnHeaders.Add , , "Sort Order", 0
    End With
End Sub

Private Sub BubbleSortBySortOrder(arr() As Variant)
    Dim i As Long, j As Long
    Dim temp As Variant

    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i)(0) > arr(j)(0) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub MacrophaseSelect_Change()
    ' Clear the ListView and reset Work Center Code label
    Me.OperationsListView.ListItems.Clear
    Me.lblWorkCenterCode.Caption = "Work Center Code: Not found" ' Default message
    workCenterCodeTemp = "" ' Reset the temporary variable

    Dim tblRoutinesDB As ListObject
    Dim tblSelectedRoutines As ListObject
    Dim routineRow As ListRow
    Dim selectedRow As ListRow
    Dim selectedMacrophase As String
    Dim selectedPlant As String
    Dim selectedProduct As String
    Dim selectedComponent As String
    Dim listItem As listItem
    Dim trValue As Variant
    Dim teValue As Variant
    Dim numOperationsValue As Variant
    Dim microphaseValue As String
    Dim materialValue As String
    Dim machineValue As String
    Dim wireCableDimension As String
    Dim wireComponentDimension As String

    ' Set the tables for RoutinesDB and SelectedRoutines
    Set tblRoutinesDB = ThisWorkbook.Sheets("RoutinesDB").ListObjects("RoutinesDB")
    Set tblSelectedRoutines = ThisWorkbook.Sheets("2. Routines").ListObjects("SelectedRoutines") ' Assuming this is the name of the table for selections

    ' Get the selected plant, Macrophase, Product, and Component from the worksheet and ComboBox
    selectedPlant = Trim(ThisWorkbook.Sheets("1. BOM Definition").Range("C9").value)
    selectedMacrophase = Trim(Me.MacrophaseSelect.value) ' Remove any spaces around the selected Macrophase
    selectedProduct = Trim(ThisWorkbook.Sheets("2. Routines").Range("D6").value) ' Product Number from the worksheet
    selectedComponent = Trim(Me.cmbComponentSelect.value) ' Component selected from the ComboBox

    ' Loop through each row in RoutinesDB to populate the Operations ListView
    For Each routineRow In tblRoutinesDB.ListRows
        ' Check if the routine's Plant and Macrophase match the selected ones
        If Trim(routineRow.Range(tblRoutinesDB.ListColumns("Plant").Index).value) = selectedPlant And _
           Trim(routineRow.Range(tblRoutinesDB.ListColumns("Macrophase").Index).value) = selectedMacrophase Then
            
            ' Get routine details
            trValue = routineRow.Range(tblRoutinesDB.ListColumns("tr").Index).value
            teValue = routineRow.Range(tblRoutinesDB.ListColumns("te").Index).value
            microphaseValue = Trim(routineRow.Range(tblRoutinesDB.ListColumns("Microphase").Index).value)
            materialValue = Trim(routineRow.Range(tblRoutinesDB.ListColumns("Material").Index).value)
            machineValue = Trim(routineRow.Range(tblRoutinesDB.ListColumns("Machine").Index).value)
            wireCableDimension = Trim(routineRow.Range(tblRoutinesDB.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).value)
            wireComponentDimension = Trim(routineRow.Range(tblRoutinesDB.ListColumns("Wire/component dimensions  (mm)").Index).value)

            ' Skip rows without "tr" or "te" if the checkbox is not ticked
            If Me.chkShowAllOperations.value = False Then
                If IsEmpty(trValue) And IsEmpty(teValue) Then
                    GoTo NextRoutine
                End If
            End If

            ' Default Number of Operations to empty
            numOperationsValue = ""

           ' Check for a match in SelectedRoutines (including Component)
            For Each selectedRow In tblSelectedRoutines.ListRows
                If Trim(selectedRow.Range(tblSelectedRoutines.ListColumns("Product Number").Index).value) = selectedProduct And _
                   Trim(selectedRow.Range(tblSelectedRoutines.ListColumns("Macrophase").Index).value) = selectedMacrophase And _
                   Trim(selectedRow.Range(tblSelectedRoutines.ListColumns("Microphase").Index).value) = microphaseValue And _
                   Trim(selectedRow.Range(tblSelectedRoutines.ListColumns("Material").Index).value) = materialValue And _
                   Trim(selectedRow.Range(tblSelectedRoutines.ListColumns("Machine").Index).value) = machineValue And _
                   Trim(selectedRow.Range(tblSelectedRoutines.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).value) = wireCableDimension And _
                   Trim(selectedRow.Range(tblSelectedRoutines.ListColumns("Wire/component dimensions  (mm)").Index).value) = wireComponentDimension And _
                   Trim(selectedRow.Range(tblSelectedRoutines.ListColumns("Component").Index).value) = selectedComponent Then
            
                    ' Check if the Microphase is "Bunching"
                    If microphaseValue = "Bunching" Then
                        ' Display the evaluated formula result for the ListView
                        If selectedRow.Range(tblSelectedRoutines.ListColumns("Number of Operations").Index).HasFormula Then
                            numOperationsValue = selectedRow.Range(tblSelectedRoutines.ListColumns("Number of Operations").Index).value
                        Else
                            numOperationsValue = selectedRow.Range(tblSelectedRoutines.ListColumns("Number of Operations").Index).value
                        End If
                    Else
                        ' For non-Bunching rows, use the stored value
                        numOperationsValue = selectedRow.Range(tblSelectedRoutines.ListColumns("Number of Operations").Index).value
                    End If
                    Exit For
                End If
            Next selectedRow


            ' Set the Work Center Code if not already set and store it in the temporary variable
            If workCenterCodeTemp = "" Then
                workCenterCodeTemp = routineRow.Range(tblRoutinesDB.ListColumns("Work Center Code").Index).value
                Me.lblWorkCenterCode.Caption = "Work Center Code: " & workCenterCodeTemp
            End If

            ' Add a new row with operation details, following the header order
            Set listItem = Me.OperationsListView.ListItems.Add(, , microphaseValue)
            
            ' Populate the SubItems in the correct order
            listItem.SubItems(1) = materialValue ' Material
            listItem.SubItems(2) = machineValue ' Machine
            listItem.SubItems(3) = wireCableDimension ' Wire/Cable Diameter
            listItem.SubItems(4) = wireComponentDimension ' Wire/Component Dimension
            listItem.SubItems(5) = IIf(IsEmpty(trValue), "", trValue) ' Setup [sec]
            listItem.SubItems(6) = IIf(IsEmpty(teValue), "", teValue) ' Manufacturing [sec]
            listItem.SubItems(7) = IIf(IsEmpty(numOperationsValue), "", numOperationsValue) ' Number of Operations
            ' Populate Sort Order (hidden column)
            listItem.SubItems(8) = routineRow.Range(tblRoutinesDB.ListColumns("Sort Order").Index).value
        End If
NextRoutine:
    Next routineRow

    ' Notify if no operations were found
    ' If Me.OperationsListView.ListItems.Count = 0 Then
    '    MsgBox "No operations found for the selected Macrophase.", vbInformation
    ' End If
End Sub

Private Sub chkShowAllOperations_Click()
    ' Refresh the ListView when the checkbox value changes
    MacrophaseSelect_Change
End Sub


Private Sub AddButton_Click()
    Dim wsDestination As Worksheet
    Dim destTable As ListObject
    Dim newRow As ListRow
    Dim tblRoutinesDB As ListObject
    Dim routineRow As ListRow
    Dim listItem As listItem
    Dim microphaseCode As String
    Dim hasOperations As Boolean
    Dim selectedComponent As String
    Dim selectedProduct As String
    Dim decimalSeparator As String
    Dim validatedOperations As Double
    Dim operationExists As Boolean
    Dim matchingRow As ListRow
    Dim bundleCount As Variant ' Variable to store the number of bundles
    Dim currentMacrophase As String
    Dim temporaryMacrophase As String

    ' Get the system decimal separator
    decimalSeparator = application.International(xlDecimalSeparator)

    ' Get the selected component from ComboBox
    selectedComponent = Me.cmbComponentSelect.value
    selectedProduct = Trim(ThisWorkbook.Sheets("2. Routines").Range("D6").value)

    ' Set references to the relevant sheets and tables
    Set wsDestination = ThisWorkbook.Sheets("2. Routines")
    Set destTable = wsDestination.ListObjects("SelectedRoutines")
    Set tblRoutinesDB = ThisWorkbook.Sheets("RoutinesDB").ListObjects("RoutinesDB")

    hasOperations = False ' Initialize the flag

    ' Loop through each item in the ListView
    For Each listItem In Me.OperationsListView.ListItems
        ' Only add rows where the "Number of Operations" column is not empty
        If listItem.SubItems(7) <> "" Then
            hasOperations = True ' At least one operation has "Number of Operations" defined

            ' Correct the decimal separator in "Number of Operations"
            validatedOperations = Replace(listItem.SubItems(7), ".", decimalSeparator)
            validatedOperations = Replace(validatedOperations, ",", decimalSeparator)

            ' Validate the corrected operations
            'If Not IsNumeric(validatedOperations) Or Val(validatedOperations) <= 0 Then
            '    MsgBox "Invalid number of operations for " & listItem.Text & ".", vbExclamation
            '    GoTo NextOperation
            'End If
            Debug.Print "Number of operations: " & validatedOperations
            

            ' Retrieve the specific Microphase code based on selected Macrophase and Operation
            microphaseCode = Trim(listItem.Text)

            ' Check if the operation already exists in the SelectedRoutines table
            operationExists = False
            For Each matchingRow In destTable.ListRows
                If Trim(matchingRow.Range(destTable.ListColumns("Product Number").Index).value) = selectedProduct And _
                   Trim(matchingRow.Range(destTable.ListColumns("Macrophase").Index).value) = Me.MacrophaseSelect.value And _
                   Trim(matchingRow.Range(destTable.ListColumns("Microphase").Index).value) = microphaseCode And _
                   Trim(matchingRow.Range(destTable.ListColumns("Material").Index).value) = listItem.SubItems(1) And _
                   Trim(matchingRow.Range(destTable.ListColumns("Machine").Index).value) = listItem.SubItems(2) And _
                   Trim(matchingRow.Range(destTable.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).value) = listItem.SubItems(3) And _
                   Trim(matchingRow.Range(destTable.ListColumns("Wire/component dimensions  (mm)").Index).value) = listItem.SubItems(4) And _
                   Trim(matchingRow.Range(destTable.ListColumns("Component").Index).value) = selectedComponent Then

                    ' Update the "Number of Operations" column if the operation exists
                    If microphaseCode = "Bunching" Then
                        ' Restore the formula for "Number of Operations"
                        matchingRow.Range(destTable.ListColumns("Number of Operations").Index).Formula = _
                            "=SUMIFS(" & _
                            "SelectedRoutines[Number of operations]," & _
                            "SelectedRoutines[Macrophase],""Cutting""," & _
                            "SelectedRoutines[Machine],[@Machine]," & _
                            "SelectedRoutines[Microphase],""<>Bunching""," & _
                            "SelectedRoutines[Work Center Code],[@Work Center Code]," & _
                            "SelectedRoutines[te],""<>0""," & _
                            "SelectedRoutines[Product Number],[@Product Number])," & _
                            "[Component],[@Component])"
                    Else
                        ' Update the number of operations with the validated value for non-Bunching
                        matchingRow.Range(destTable.ListColumns("Number of Operations").Index).value = validatedOperations
                    End If

                    operationExists = True
                    Exit For
                End If
            Next matchingRow

            ' If the operation does not exist, add a new row
            If Not operationExists Then
                If destTable.ListRows.Count = 1 And destTable.ListRows(1).Range.Cells(1) = "" Then
                    Set newRow = destTable.ListRows(1)
                Else
                    Set newRow = destTable.ListRows.Add
                End If
                
                ' Copy data to the new row using column names
                newRow.Range(destTable.ListColumns("Plant").Index).value = wsDestination.Range("D5").value
                newRow.Range(destTable.ListColumns("Macrophase").Index).value = Me.MacrophaseSelect.value
                newRow.Range(destTable.ListColumns("Microphase").Index).value = microphaseCode
                newRow.Range(destTable.ListColumns("Material").Index).value = listItem.SubItems(1) ' Material
                newRow.Range(destTable.ListColumns("Machine").Index).value = listItem.SubItems(2) ' Machine
                newRow.Range(destTable.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).value = listItem.SubItems(3) ' Wire/Cable Diameter
                newRow.Range(destTable.ListColumns("Wire/component dimensions  (mm)").Index).value = listItem.SubItems(4) ' Wire/Component Dimension
                newRow.Range(destTable.ListColumns("tr").Index).value = listItem.SubItems(5) ' Setup time (tr)
                newRow.Range(destTable.ListColumns("te").Index).value = listItem.SubItems(6) ' Manufacturing time (te)
                newRow.Range(destTable.ListColumns("Sort Order").Index).value = listItem.SubItems(8) ' Sort Order
                newRow.Range(destTable.ListColumns("Number of Setups").Index).value = 1 ' Number of setups default

                ' Prompt for bundle count if the Microphase is "Bunching"
                If microphaseCode = "Bunching" Then
                    bundleCount = InputBox("Enter the number of bundles for the 'Bunching' operation:", "Input Bundles")
                    If Not IsNumeric(bundleCount) Or val(bundleCount) <= 0 Then
                        MsgBox "Invalid bundle count. Please enter a positive number.", vbExclamation
                        GoTo NextOperation
                    End If
                    newRow.Range(destTable.ListColumns("Bundles").Index).value = bundleCount

                    ' Set the formula for "Number of Operations"
                    newRow.Range(destTable.ListColumns("Number of Operations").Index).Formula = _
                        "=SUMIFS(" & _
                        "SelectedRoutines[Number of operations]," & _
                        "SelectedRoutines[Macrophase],""Cutting""," & _
                        "SelectedRoutines[Machine],[@Machine]," & _
                        "SelectedRoutines[Microphase],""<>Bunching""," & _
                        "SelectedRoutines[Work Center Code],[@Work Center Code]," & _
                        "SelectedRoutines[te],""<>0""," & _
                        "SelectedRoutines[Product Number],[@Product Number])"
                Else
                    ' Set the validated value for "Number of Operations"
                    newRow.Range(destTable.ListColumns("Number of Operations").Index).value = validatedOperations
                End If

                ' Add the Work Center Code
                newRow.Range(destTable.ListColumns("Work Center Code").Index).value = workCenterCodeTemp

                ' Add the selected component's Material Number to the "Component" column if one is selected
                If selectedComponent <> "" Then
                    newRow.Range(destTable.ListColumns("Component").Index).value = selectedComponent
                End If

                ' Copy the contents of "Product Specification" sheet cell B5 to the "Product" column in the new row
                newRow.Range(destTable.ListColumns("Product Number").Index).value = selectedProduct
                newRow.Range(destTable.ListColumns("Product Type").Index).value = wsDestination.Range("D8").value

               ' Add the Total Tr formula using structured references
                Dim formulaCell As Range
                Dim rawFormula As String
                
                Set formulaCell = ThisWorkbook.Sheets("2. Routines").Range("AD1") ' Update with actual Cell address
                rawFormula = formulaCell.Formula
                
                ' Remove leading apostrophe if Excel stored it as text
                Dim cleanFormula As String
                cleanFormula = Replace(formulaCell.Formula, "'", "")
                
                Debug.Print "Raw Formula: " & rawFormula
                Debug.Print "Cleaned Formula: " & cleanFormula
                
                
                ' Apply the formula to the structured table column
                newRow.Range(destTable.ListColumns("Total Tr").Index).Formula = rawFormula


                newRow.Range(destTable.ListColumns("Total Te").Index).Formula = "=[@[Number of operations]]*[@te]/60"
            End If
        End If
NextOperation:
    Next listItem

    ' Simulate changing the Macrophase to refresh the table
    If Me.MacrophaseSelect.ListCount > 1 Then
        currentMacrophase = Me.MacrophaseSelect.value ' Save the current selection
        temporaryMacrophase = Me.MacrophaseSelect.List(0) ' Choose a different Macrophase temporarily

        ' Ensure the temporary Macrophase is different from the current one
        If temporaryMacrophase = currentMacrophase Then
            temporaryMacrophase = Me.MacrophaseSelect.List(1)
        End If

        ' Set the Macrophase to the temporary value and trigger the change event
        Me.MacrophaseSelect.value = temporaryMacrophase
        Call MacrophaseSelect_Change

        ' Restore the original Macrophase and trigger the change event
        Me.MacrophaseSelect.value = currentMacrophase
        Call MacrophaseSelect_Change
    End If

    ' Check if no operations were added and show an error
    If Not hasOperations Then
        MsgBox "No operations added. Please double-click the needed operation and add the number.", vbExclamation
    Else
        ' MsgBox "Operations added or updated in Selected Routines!", vbInformation
    End If

    SortSelectedRoutingByProduct
End Sub



' --- UPDATED: double-click edits AND persists immediately ---
Private Sub OperationsListView_DblClick()
    Dim selectedItem As listItem
    Dim newValue As Variant
    Dim decimalSeparator As String

    decimalSeparator = application.International(xlDecimalSeparator)

    If Me.OperationsListView.selectedItem Is Nothing Then
        MsgBox "Please select an operation before double-clicking to edit.", vbExclamation
        Exit Sub
    End If

    Set selectedItem = Me.OperationsListView.selectedItem

    newValue = InputBox("Enter the number of operations:", _
                        "Edit Number of Operations", selectedItem.SubItems(7))

    If Len(newValue) = 0 Then Exit Sub

    If InStr(newValue, ".") > 0 And decimalSeparator <> "." Then
        newValue = Replace(newValue, ".", decimalSeparator)
    ElseIf InStr(newValue, ",") > 0 And decimalSeparator <> "," Then
        newValue = Replace(newValue, ",", decimalSeparator)
    End If

    If IsNumeric(newValue) Then
        ' update the list display
        selectedItem.SubItems(7) = newValue
        ' immediately add/update the table row
        PersistOperationFromListItem selectedItem
        ' refresh the list so formulas (e.g., Bunching) show evaluated values
        MacrophaseSelect_Change
        ' optional: keep things sorted as you already do elsewhere
        SortSelectedRoutingByProduct
    Else
        MsgBox "Please enter a valid number.", vbExclamation
    End If
End Sub


Private Sub cmbComponentSelect_Change()
    Dim wsComponents As Worksheet
    Dim tblComponents As ListObject
    Dim selectedComponent As String
    Dim componentRow As ListRow
    Dim materialDescription As String
    Dim cableDia As String
    Dim quantity As String

    ' Set the worksheet and table
    Set wsComponents = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblComponents = wsComponents.ListObjects("BOMDefinition")

    ' Get the selected component from the ComboBox
    selectedComponent = Me.cmbComponentSelect.value

    ' Check if the component is selected
    If selectedComponent = "" Then
        Me.Label9.Caption = "No component selected."
        Exit Sub
    End If

    ' Loop through each row in the table to find the selected component
    On Error Resume Next
    For Each componentRow In tblComponents.ListRows
        If componentRow.Range(tblComponents.ListColumns("Material").Index).value = selectedComponent Then
            ' Get the Material Description and Cable Dia
            materialDescription = componentRow.Range(tblComponents.ListColumns("Material Description").Index).value
            cableDia = componentRow.Range(tblComponents.ListColumns("Cable Dia.").Index).value
            quantity = componentRow.Range(tblComponents.ListColumns("Quantity").Index).value
            Exit For
        End If
    Next componentRow
    On Error GoTo 0

    ' Update the label with the component details
    If materialDescription <> "" Or cableDia <> "" Then
        Me.Label9.Caption = "Description: " & materialDescription & vbCrLf & "Cable Diameter: " & cableDia & vbCrLf & "Quantity: " & quantity
    Else
        Me.Label9.Caption = "Component details not found."
    End If
    
    ' Call MacrophaseSelect_Change to refresh the list
    Call MacrophaseSelect_Change
End Sub


Private Sub CopyRoutinesButton_Click()
    Dim wsRoutinesDB As Worksheet
    Dim tblRoutinesDB As ListObject
    Dim tblSelectedRoutines As ListObject
    Dim wsSelectedRoutines As Worksheet
    Dim selectedPlant As String
    Dim selectedProduct As String
    Dim routineRow As ListRow
    Dim newRow As ListRow
    Dim existingRow As ListRow
    Dim isDuplicate As Boolean
    Dim sortedRoutines As Collection
    Dim sortOrder As Double
    Dim sortedRoutineRows() As Variant
    Dim rowIdx As Long
    Dim col As ListColumn

    ' Get the selected plant and product from the "Product Specification" sheet
    Set wsSelectedRoutines = ThisWorkbook.Sheets("2. Routines")
    selectedPlant = Trim(wsSelectedRoutines.Range("D5").value)
    selectedProduct = Trim(wsSelectedRoutines.Range("D6").value)

    ' Validate input
    If selectedPlant = "" Or selectedProduct = "" Then
        MsgBox "Please select both a plant and a product in the Product Specification sheet.", vbExclamation
        Exit Sub
    End If

    ' Set references to the relevant sheets and tables
    Set wsRoutinesDB = ThisWorkbook.Sheets("RoutinesDB")
    Set tblRoutinesDB = wsRoutinesDB.ListObjects("RoutinesDB")
    Set tblSelectedRoutines = wsSelectedRoutines.ListObjects("SelectedRoutines")

    ' Collect all rows from RoutinesDB for the selected plant
    ReDim sortedRoutineRows(tblRoutinesDB.ListRows.Count - 1) ' Prepare array for sorting
    rowIdx = 0

    For Each routineRow In tblRoutinesDB.ListRows
        If Trim(routineRow.Range(tblRoutinesDB.ListColumns("Plant").Index).value) = selectedPlant Then
            ' Get the Sort Order value
            sortOrder = routineRow.Range(tblRoutinesDB.ListColumns("Sort Order").Index).value
            If IsNumeric(sortOrder) Then
                sortedRoutineRows(rowIdx) = Array(sortOrder, routineRow)
                rowIdx = rowIdx + 1
            End If
        End If
    Next routineRow

    ' Resize the array to the actual count of rows
    ReDim Preserve sortedRoutineRows(rowIdx - 1)

    ' Sort the array based on Sort Order
    If UBound(sortedRoutineRows) >= 0 Then
        Call BubbleSortBySortOrder(sortedRoutineRows)
    End If

    ' Copy sorted routines to SelectedRoutines table
    For rowIdx = LBound(sortedRoutineRows) To UBound(sortedRoutineRows)
        Set routineRow = sortedRoutineRows(rowIdx)(1) ' Extract the routine row from the sorted array

        ' Check if the routine already exists in the SelectedRoutines table
        isDuplicate = False
        For Each existingRow In tblSelectedRoutines.ListRows
            If Trim(existingRow.Range(tblSelectedRoutines.ListColumns("Product Number").Index).value) = selectedProduct Then
                isDuplicate = True
                
                ' Check all column values
                For Each col In tblSelectedRoutines.ListColumns
                    On Error Resume Next
                    Dim dbCol As ListColumn
                    Set dbCol = tblRoutinesDB.ListColumns(col.name)
                    On Error GoTo 0
                
                    ' Skip if the column does not exist in tblRoutinesDB
                    If Not dbCol Is Nothing Then
                        If Trim(existingRow.Range(col.Index).value) <> Trim(routineRow.Range(dbCol.Index).value) Then
                            isDuplicate = False
                            Exit For
                        End If
                    End If
                Next col

                
                If isDuplicate Then Exit For
            End If
        Next existingRow

        ' Add the routine to the SelectedRoutines table if it's not a duplicate
        If Not isDuplicate Then
            If tblSelectedRoutines.ListRows.Count = 1 And tblSelectedRoutines.ListRows(1).Range.Cells(1) = "" Then
                Set newRow = tblSelectedRoutines.ListRows(1)
            Else
                Set newRow = tblSelectedRoutines.ListRows.Add
            End If

            ' Copy all column values dynamically
            For Each col In tblSelectedRoutines.ListColumns
                On Error Resume Next
                newRow.Range(col.Index).value = routineRow.Range(tblRoutinesDB.ListColumns(col.name).Index).value
                On Error GoTo 0
            Next col

            ' Add the selected Product to the "Product Number" column
            newRow.Range(tblSelectedRoutines.ListColumns("Product Number").Index).value = selectedProduct

            ' Add the Product type from the Product Specification sheet
            newRow.Range(tblSelectedRoutines.ListColumns("Product Type").Index).value = wsSelectedRoutines.Range("D8").value

            ' Add the Total Tr and Total Te formulas
            ' Add the Total Tr formula using structured references
            Dim formulaCell As Range
            Dim rawFormula As String
            
            Set formulaCell = ThisWorkbook.Sheets("2. Routines").Range("AE1") ' Update with actual sheet name
            rawFormula = formulaCell.Formula
            
            ' Remove leading apostrophe if Excel stored it as text
            Dim cleanFormula As String
            cleanFormula = Replace(formulaCell.Formula, "'", "")
            
            Debug.Print "Raw Formula: " & rawFormula
            Debug.Print "Cleaned Formula: " & cleanFormula
            
            
            ' Apply the formula to the structured table column

            newRow.Range(tblSelectedRoutines.ListColumns("Number of Setups").Index).value = 1 ' Number of setups default
            newRow.Range(tblSelectedRoutines.ListColumns("ProductNumberText").Index).Formula = "= """" & [@[Product Number]]"
        End If
    Next rowIdx
    
    ' Sort table after addition
    SortSelectedRoutingByProduct
    MsgBox "Routines copied successfully for the selected plant.", vbInformation
End Sub


' --- helper: resolve "Number of Operations" column name safely ---
Private Function GetOpsCol(lo As ListObject, ByRef opsHeader As String) As Long
    Dim nm
    For Each nm In Array("Number of Operations", "Number of operations")
        On Error Resume Next
        GetOpsCol = lo.ListColumns(CStr(nm)).Index
        On Error GoTo 0
        If GetOpsCol > 0 Then
            opsHeader = lo.ListColumns(GetOpsCol).name
            Exit Function
        End If
    Next nm
    opsHeader = "Number of Operations"
    GetOpsCol = 0
End Function

' --- core: add or update the SelectedRoutines row for a single ListItem ---
Private Sub PersistOperationFromListItem(li As listItem)
    Dim wsDest As Worksheet, loDest As ListObject
    Dim selectedComponent As String, selectedProduct As String, mac As String
    Dim micro As String, material As String, machine As String
    Dim dia As String, dimm As String, trVal As Variant, teVal As Variant
    Dim sortOrder As Variant, bundles As Variant
    Dim r As ListRow, exists As Boolean
    Dim idxOps As Long, opsHeader As String
    Dim decSep As String, opsVal As Variant

    Set wsDest = ThisWorkbook.Sheets("2. Routines")
    Set loDest = wsDest.ListObjects("SelectedRoutines")

    mac = Me.MacrophaseSelect.value
    selectedComponent = Me.cmbComponentSelect.value
    selectedProduct = Trim$(wsDest.Range("D6").value)

    micro = Trim$(li.Text)
    material = Trim$(li.SubItems(1))
    machine = Trim$(li.SubItems(2))
    dia = Trim$(li.SubItems(3))
    dimm = Trim$(li.SubItems(4))
    trVal = li.SubItems(5)
    teVal = li.SubItems(6)
    sortOrder = li.SubItems(8)

    decSep = application.International(xlDecimalSeparator)
    opsVal = li.SubItems(7)
    If InStr(opsVal, ".") > 0 And decSep <> "." Then opsVal = Replace(opsVal, ".", decSep)
    If InStr(opsVal, ",") > 0 And decSep <> "," Then opsVal = Replace(opsVal, ",", decSep)

    idxOps = GetOpsCol(loDest, opsHeader)

    ' Try to find an existing matching row
    exists = False
    For Each r In loDest.ListRows
        If Trim$(r.Range(loDest.ListColumns("Product Number").Index).value) = selectedProduct _
        And Trim$(r.Range(loDest.ListColumns("Macrophase").Index).value) = mac _
        And Trim$(r.Range(loDest.ListColumns("Microphase").Index).value) = micro _
        And Trim$(r.Range(loDest.ListColumns("Material").Index).value) = material _
        And Trim$(r.Range(loDest.ListColumns("Machine").Index).value) = machine _
        And Trim$(r.Range(loDest.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).value) = dia _
        And Trim$(r.Range(loDest.ListColumns("Wire/component dimensions  (mm)").Index).value) = dimm _
        And Trim$(r.Range(loDest.ListColumns("Component").Index).value) = selectedComponent Then
            ' Update No. of Operations
            If micro = "Bunching" Then
                r.Range(idxOps).Formula = "=SUMIFS(" & _
                    "SelectedRoutines[" & opsHeader & "]," & _
                    "SelectedRoutines[Macrophase],""Cutting""," & _
                    "SelectedRoutines[Machine],[@Machine]," & _
                    "SelectedRoutines[Microphase],""<>Bunching""," & _
                    "SelectedRoutines[Work Center Code],[@Work Center Code]," & _
                    "SelectedRoutines[te],""<>0""," & _
                    "SelectedRoutines[Product Number],[@Product Number])"
            Else
                If idxOps > 0 Then r.Range(idxOps).value = opsVal
            End If
            exists = True
            Exit For
        End If
    Next r

    If exists Then Exit Sub

    ' Create new row
    Dim newRow As ListRow
    If loDest.ListRows.Count = 1 And loDest.ListRows(1).Range.Cells(1).value = "" Then
        Set newRow = loDest.ListRows(1)
    Else
        Set newRow = loDest.ListRows.Add
    End If

    ' Required fields
    newRow.Range(loDest.ListColumns("Plant").Index).value = wsDest.Range("D5").value
    newRow.Range(loDest.ListColumns("Macrophase").Index).value = mac
    newRow.Range(loDest.ListColumns("Microphase").Index).value = micro
    newRow.Range(loDest.ListColumns("Material").Index).value = material
    newRow.Range(loDest.ListColumns("Machine").Index).value = machine
    newRow.Range(loDest.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).value = dia
    newRow.Range(loDest.ListColumns("Wire/component dimensions  (mm)").Index).value = dimm
    newRow.Range(loDest.ListColumns("tr").Index).value = IIf(IsEmpty(trVal), "", trVal)
    newRow.Range(loDest.ListColumns("te").Index).value = IIf(IsEmpty(teVal), "", teVal)
    newRow.Range(loDest.ListColumns("Sort Order").Index).value = sortOrder
    newRow.Range(loDest.ListColumns("Number of Setups").Index).value = 1
    newRow.Range(loDest.ListColumns("Work Center Code").Index).value = workCenterCodeTemp
    If selectedComponent <> "" Then newRow.Range(loDest.ListColumns("Component").Index).value = selectedComponent
    newRow.Range(loDest.ListColumns("Product Number").Index).value = selectedProduct
    newRow.Range(loDest.ListColumns("Product Type").Index).value = wsDest.Range("D8").value

    ' Bunching special case
    If micro = "Bunching" Then
        bundles = InputBox("Enter the number of bundles for the 'Bunching' operation:", "Input Bundles")
        If IsNumeric(bundles) And val(bundles) > 0 Then
            On Error Resume Next
            newRow.Range(loDest.ListColumns("Bundles").Index).value = bundles
            On Error GoTo 0
        End If
        If idxOps > 0 Then
            newRow.Range(idxOps).Formula = "=SUMIFS(" & _
                "SelectedRoutines[" & opsHeader & "]," & _
                "SelectedRoutines[Macrophase],""Cutting""," & _
                "SelectedRoutines[Machine],[@Machine]," & _
                "SelectedRoutines[Microphase],""<>Bunching""," & _
                "SelectedRoutines[Work Center Code],[@Work Center Code]," & _
                "SelectedRoutines[te],""<>0""," & _
                "SelectedRoutines[Product Number],[@Product Number])"
        End If
    Else
        If idxOps > 0 Then newRow.Range(idxOps).value = opsVal
    End If

    ' Total Tr from a template cell (AD1). If missing, skip silently.
    On Error Resume Next
    Dim trTemplate As String
    trTemplate = wsDest.Range("AD1").Formula
    If Len(trTemplate) > 0 Then newRow.Range(loDest.ListColumns("Total Tr").Index).Formula = trTemplate
    ' Total Te uses whatever the actual header is
    If idxOps > 0 Then newRow.Range(loDest.ListColumns("Total Te").Index).Formula = "=[@[" & opsHeader & "]]*[@te]/60"
    On Error GoTo 0
End Sub




