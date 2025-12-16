VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RoutineForm 
   Caption         =   "Add Manufacturing Operations"
   ClientHeight    =   6105
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   17460
   OleObjectBlob   =   "RoutineForm.frx":0000
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

' --- LISTVIEW COLUMN MAPPING CONSTANTS ---
' This makes your code "Break-Proof" if you add more columns later.
Private Const COL_MICROPHASE As Long = 1
Private Const COL_MATERIAL As Long = 2
Private Const COL_MACHINE As Long = 3
'Private Const COL_NEW_COLUMN As Long = 4  ' <--- THIS IS YOUR NEW COLUMN
Private Const COL_WIRE_DIA As Long = 5
Private Const COL_COMP_DIM As Long = 6
Private Const COL_SETUP As Long = 7
Private Const COL_MFG As Long = 8
Private Const COL_NUM_OPS As Long = 9
Private Const COL_SORT As Long = 10


' Property to allow setting preselectedComponent from outside
Public Property Let PreselectedComponent(Value As String)
    mPreselectedComponent = Value
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
    selectedPlant = Trim(wsMain.Range("C9").Value)
    selectedProduct = Trim(wsSelectedRoutines.Range("D6").Value)
    Me.lblselectedProduct.Caption = "Selected Final Product: " & selectedProduct

    ' Use a Dictionary to store unique Macrophases and their Sort Order
    Set uniqueMacrophases = CreateObject("Scripting.Dictionary")

    ' Collect unique Macrophases with their Sort Order for the selected plant
    For Each routineRow In tblRoutinesDB.ListRows
        If Trim(routineRow.Range(tblRoutinesDB.ListColumns("Plant").Index).Value) = selectedPlant Then
            teValue = routineRow.Range(tblRoutinesDB.ListColumns("te").Index).Value
            trValue = routineRow.Range(tblRoutinesDB.ListColumns("tr").Index).Value
            macrophaseName = Trim(routineRow.Range(tblRoutinesDB.ListColumns("Macrophase").Index).Value)
            sortOrder = routineRow.Range(tblRoutinesDB.ListColumns("Sort Order").Index).Value

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
        If componentRow.Range(tblComponents.ListColumns("Product Number").Index).Value = selectedProduct Then
            Me.cmbComponentSelect.AddItem componentRow.Range(tblComponents.ListColumns("Material").Index).Value
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

' Set up the ListView headers
    With Me.OperationsListView
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
        .LabelEdit = lvwManual
        .ListItems.Clear ' Safety clear
        .columnHeaders.Clear ' Safety clear

        ' Add columns using our Constants to ensure alignment
        ' Note: The Index in .Add must match the Constant values defined above
        .columnHeaders.Add , , "Operation", 110                  ' Col 1
        .columnHeaders.Add , , "Material", 100                   ' Col 2
        .columnHeaders.Add , , "Machine", 100                    ' Col 3
        '.columnHeaders.Add , , "Extra Info", 80                  ' Col 4 <--- UPDATE NAME
        .columnHeaders.Add , , "Wire/Cable Diameter", 65         ' Col 5
        .columnHeaders.Add , , "Wire/Component Dimension", 60    ' Col 6
        .columnHeaders.Add , , "Setup [sec]", 60                 ' Col 7
        .columnHeaders.Add , , "Manufacturing [sec]", 60         ' Col 8
        .columnHeaders.Add , , "Number of Operations", 75        ' Col 9
        .columnHeaders.Add , , "Sort Order", 0                   ' Col 10 (Hidden)
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
    Dim newColValue As String ' Variable for the new column data

    ' Set the tables for RoutinesDB and SelectedRoutines
    Set tblRoutinesDB = ThisWorkbook.Sheets("RoutinesDB").ListObjects("RoutinesDB")
    Set tblSelectedRoutines = ThisWorkbook.Sheets("2. Routines").ListObjects("SelectedRoutines") ' Assuming this is the name of the table for selections

    ' Get the selected plant, Macrophase, Product, and Component from the worksheet and ComboBox
    selectedPlant = Trim(ThisWorkbook.Sheets("1. BOM Definition").Range("C9").Value)
    selectedMacrophase = Trim(Me.MacrophaseSelect.Value) ' Remove any spaces around the selected Macrophase
    selectedProduct = Trim(ThisWorkbook.Sheets("2. Routines").Range("D6").Value) ' Product Number from the worksheet
    selectedComponent = Trim(Me.cmbComponentSelect.Value) ' Component selected from the ComboBox

    ' Loop through each row in RoutinesDB to populate the Operations ListView
    For Each routineRow In tblRoutinesDB.ListRows
        ' Check if the routine's Plant and Macrophase match the selected ones
        If Trim(routineRow.Range(tblRoutinesDB.ListColumns("Plant").Index).Value) = selectedPlant And _
           Trim(routineRow.Range(tblRoutinesDB.ListColumns("Macrophase").Index).Value) = selectedMacrophase Then
            
            ' Get routine details
            trValue = routineRow.Range(tblRoutinesDB.ListColumns("tr").Index).Value
            teValue = routineRow.Range(tblRoutinesDB.ListColumns("te").Index).Value
            microphaseValue = Trim$(routineRow.Range(tblRoutinesDB.ListColumns("Microphase").Index).Value)
            materialValue = Trim$(routineRow.Range(tblRoutinesDB.ListColumns("Material").Index).Value)
            machineValue = Trim$(routineRow.Range(tblRoutinesDB.ListColumns("Machine").Index).Value)
            wireCableDimension = Trim$(routineRow.Range(tblRoutinesDB.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).Value)
            wireComponentDimension = Trim$(routineRow.Range(tblRoutinesDB.ListColumns("Wire/component dimensions  (mm)").Index).Value)
            If materialValue = "SSH according to length" Then
                Debug.Print materialValue & "; "; wireCableDimension & "; " & wireComponentDimension
            End If
            ' Skip rows without "tr" or "te" if the checkbox is not ticked
            If Me.chkShowAllOperations.Value = False Then
                If (IsEmpty(trValue) And IsEmpty(teValue)) Or (trValue = 0 And teValue = 0) Then
                    GoTo NextRoutine
                End If
            End If

            ' Default Number of Operations to empty
            numOperationsValue = ""

            ' Check for a match in SelectedRoutines (including Component)
            For Each selectedRow In tblSelectedRoutines.ListRows
                
                ' USE THE NEW HELPER FUNCTION HERE
                ' We use a boolean flag to track if all match, to make debugging easier
                Dim isMatch As Boolean
                isMatch = True
                
                If Not ValuesMatch(selectedRow.Range(tblSelectedRoutines.ListColumns("Product Number").Index).Value, selectedProduct) Then isMatch = False
                If isMatch And Not ValuesMatch(selectedRow.Range(tblSelectedRoutines.ListColumns("Macrophase").Index).Value, selectedMacrophase) Then isMatch = False
                If isMatch And Not ValuesMatch(selectedRow.Range(tblSelectedRoutines.ListColumns("Microphase").Index).Value, microphaseValue) Then isMatch = False
                If isMatch And Not ValuesMatch(selectedRow.Range(tblSelectedRoutines.ListColumns("Material").Index).Value, materialValue) Then isMatch = False
                If isMatch And Not ValuesMatch(selectedRow.Range(tblSelectedRoutines.ListColumns("Machine").Index).Value, machineValue) Then isMatch = False
                If isMatch And Not ValuesMatch(selectedRow.Range(tblSelectedRoutines.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).Value, wireCableDimension) Then isMatch = False
                If isMatch And Not ValuesMatch(selectedRow.Range(tblSelectedRoutines.ListColumns("Wire/component dimensions  (mm)").Index).Value, wireComponentDimension) Then isMatch = False
                If isMatch And Not ValuesMatch(selectedRow.Range(tblSelectedRoutines.ListColumns("Component").Index).Value, selectedComponent) Then isMatch = False
                
                ' If we survived all checks
                If isMatch Then
            
                    ' Check if the Microphase is "Bunching"
                    If microphaseValue = "Bunching" Then
                        ' Display the evaluated formula result for the ListView
                        numOperationsValue = selectedRow.Range(tblSelectedRoutines.ListColumns("Number of Operations").Index).Value
                    Else
                        ' For non-Bunching rows, use the stored value
                        numOperationsValue = selectedRow.Range(tblSelectedRoutines.ListColumns("Number of Operations").Index).Value
                    End If
                    Exit For
                End If
            Next selectedRow


            ' Set the Work Center Code if not already set and store it in the temporary variable
            If workCenterCodeTemp = "" Then
                workCenterCodeTemp = routineRow.Range(tblRoutinesDB.ListColumns("Work Center Code").Index).Value
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
            listItem.SubItems(8) = routineRow.Range(tblRoutinesDB.ListColumns("Sort Order").Index).Value
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
    selectedComponent = Me.cmbComponentSelect.Value
    selectedProduct = Trim(ThisWorkbook.Sheets("2. Routines").Range("D6").Value)

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
                If Trim(matchingRow.Range(destTable.ListColumns("Product Number").Index).Value) = selectedProduct And _
                   Trim(matchingRow.Range(destTable.ListColumns("Macrophase").Index).Value) = Me.MacrophaseSelect.Value And _
                   Trim(matchingRow.Range(destTable.ListColumns("Microphase").Index).Value) = microphaseCode And _
                   Trim(matchingRow.Range(destTable.ListColumns("Material").Index).Value) = listItem.SubItems(1) And _
                   Trim(matchingRow.Range(destTable.ListColumns("Machine").Index).Value) = listItem.SubItems(2) And _
                   Trim(matchingRow.Range(destTable.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).Value) = listItem.SubItems(3) And _
                   Trim(matchingRow.Range(destTable.ListColumns("Wire/component dimensions  (mm)").Index).Value) = listItem.SubItems(4) And _
                   Trim(matchingRow.Range(destTable.ListColumns("Component").Index).Value) = selectedComponent Then

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
                        matchingRow.Range(destTable.ListColumns("Number of Operations").Index).Value = validatedOperations
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
                newRow.Range(destTable.ListColumns("Plant").Index).Value = wsDestination.Range("D5").Value
                newRow.Range(destTable.ListColumns("Macrophase").Index).Value = Me.MacrophaseSelect.Value
                newRow.Range(destTable.ListColumns("Microphase").Index).Value = microphaseCode
                newRow.Range(destTable.ListColumns("Material").Index).Value = listItem.SubItems(1) ' Material
                newRow.Range(destTable.ListColumns("Machine").Index).Value = listItem.SubItems(2) ' Machine
                newRow.Range(destTable.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).Value = listItem.SubItems(3) ' Wire/Cable Diameter
                newRow.Range(destTable.ListColumns("Wire/component dimensions  (mm)").Index).Value = listItem.SubItems(4) ' Wire/Component Dimension
                newRow.Range(destTable.ListColumns("tr").Index).Value = listItem.SubItems(5) ' Setup time (tr)
                newRow.Range(destTable.ListColumns("te").Index).Value = listItem.SubItems(6) ' Manufacturing time (te)
                newRow.Range(destTable.ListColumns("Sort Order").Index).Value = listItem.SubItems(8) ' Sort Order
                newRow.Range(destTable.ListColumns("Number of Setups").Index).Value = 1 ' Number of setups default

                ' Prompt for bundle count if the Microphase is "Bunching"
                If microphaseCode = "Bunching" Then
                    bundleCount = InputBox("Enter the number of bundles for the 'Bunching' operation:", "Input Bundles")
                    If Not IsNumeric(bundleCount) Or val(bundleCount) <= 0 Then
                        MsgBox "Invalid bundle count. Please enter a positive number.", vbExclamation
                        GoTo NextOperation
                    End If
                    newRow.Range(destTable.ListColumns("Bundles").Index).Value = bundleCount

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
                    newRow.Range(destTable.ListColumns("Number of Operations").Index).Value = validatedOperations
                End If

                ' Add the Work Center Code
                newRow.Range(destTable.ListColumns("Work Center Code").Index).Value = workCenterCodeTemp

                ' Add the selected component's Material Number to the "Component" column if one is selected
                If selectedComponent <> "" Then
                    newRow.Range(destTable.ListColumns("Component").Index).Value = selectedComponent
                End If

                ' Copy the contents of "Product Specification" sheet cell B5 to the "Product" column in the new row
                newRow.Range(destTable.ListColumns("Product Number").Index).Value = selectedProduct
                newRow.Range(destTable.ListColumns("Product Type").Index).Value = wsDestination.Range("D8").Value

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
        currentMacrophase = Me.MacrophaseSelect.Value ' Save the current selection
        temporaryMacrophase = Me.MacrophaseSelect.List(0) ' Choose a different Macrophase temporarily

        ' Ensure the temporary Macrophase is different from the current one
        If temporaryMacrophase = currentMacrophase Then
            temporaryMacrophase = Me.MacrophaseSelect.List(1)
        End If

        ' Set the Macrophase to the temporary value and trigger the change event
        Me.MacrophaseSelect.Value = temporaryMacrophase
        Call MacrophaseSelect_Change

        ' Restore the original Macrophase and trigger the change event
        Me.MacrophaseSelect.Value = currentMacrophase
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
    selectedComponent = Me.cmbComponentSelect.Value

    ' Check if the component is selected
    If selectedComponent = "" Then
        Me.Label9.Caption = "No component selected."
        Exit Sub
    End If

    ' Loop through each row in the table to find the selected component
    On Error Resume Next
    For Each componentRow In tblComponents.ListRows
        If componentRow.Range(tblComponents.ListColumns("Material").Index).Value = selectedComponent Then
            ' Get the Material Description and Cable Dia
            materialDescription = componentRow.Range(tblComponents.ListColumns("Material Description").Index).Value
            cableDia = componentRow.Range(tblComponents.ListColumns("Cable Dia.").Index).Value
            quantity = componentRow.Range(tblComponents.ListColumns("Quantity").Index).Value
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
    Dim wsRoutinesDB As Worksheet, wsSelectedRoutines As Worksheet
    Dim tblRoutinesDB As ListObject, tblSelectedRoutines As ListObject
    Dim selectedPlant As String, selectedProduct As String
    Dim productType As String
    
    ' Data Arrays
    Dim arrSource As Variant
    Dim arrDest As Variant
    Dim arrNewData() As Variant
    
    ' Loop Variables
    Dim i As Long, j As Long, r As Long
    Dim countNew As Long
    Dim dictExisting As Object
    Dim key As String
    
    ' Column Mappings (Source Index -> Dest Index)
    Dim colMap As Object
    Dim srcHeader As Variant, destHeader As Variant
    
    ' Sorting
    Dim arrSort() As Variant
    Dim tempArr As Variant
    
    ' Initialize Sheets and Tables
    Set wsSelectedRoutines = ThisWorkbook.Sheets("2. Routines")
    Set wsRoutinesDB = ThisWorkbook.Sheets("RoutinesDB")
    
    On Error Resume Next
    Set tblRoutinesDB = wsRoutinesDB.ListObjects("RoutinesDB")
    Set tblSelectedRoutines = wsSelectedRoutines.ListObjects("SelectedRoutines")
    On Error GoTo 0
    
    If tblRoutinesDB Is Nothing Or tblSelectedRoutines Is Nothing Then Exit Sub
    
    ' Inputs
    selectedPlant = Trim(wsSelectedRoutines.Range("D5").Value)
    selectedProduct = Trim(wsSelectedRoutines.Range("D6").Value)
    productType = wsSelectedRoutines.Range("D8").Value
    
    If selectedPlant = "" Or selectedProduct = "" Then
        MsgBox "Please select both a plant and a product.", vbExclamation
        Exit Sub
    End If

    ' --- PHASE 1: PREPARATION & PERFORMANCE ON ---
    Call SpeedOn
    
    ' 1. Map Columns: Map DB Header Name to Selected Header Name
    ' This allows us to copy data dynamically even if columns move
    Set colMap = CreateObject("Scripting.Dictionary")
    
    ' Define which columns we copy FROM DB -> TO Selected
    ' Format: Array("DB_Header", "Selected_Header")
    ' Add more mappings here as needed
    Dim mappingList As Variant
    mappingList = Array( _
        Array("Macrophase", "Macrophase"), _
        Array("Microphase", "Microphase"), _
        Array("Material", "Material"), _
        Array("Machine", "Machine"), _
        Array("Wire/cable dimension diameter/section  (mm/mm2)", "Wire/cable dimension diameter/section  (mm/mm2)"), _
        Array("Wire/component dimensions  (mm)", "Wire/component dimensions  (mm)"), _
        Array("tr", "tr"), _
        Array("te", "te"), _
        Array("Sort Order", "Sort Order"), _
        Array("Work Center Code", "Work Center Code") _
    )
    
    ' 2. Load Existing Data to Dictionary (To check duplicates FAST)
    Set dictExisting = CreateObject("Scripting.Dictionary")
    
    If Not tblSelectedRoutines.DataBodyRange Is Nothing Then
        arrDest = tblSelectedRoutines.DataBodyRange.Value
        
        ' Build composite key: Product + Macrophase + Microphase + Material + Machine + WorkCenter
        ' Adjust these column numbers based on your table structure
        Dim idxProd As Long, idxMacro As Long, idxMicro As Long, idxMat As Long, idxMach As Long
        
        With tblSelectedRoutines.ListColumns
            idxProd = .item("Product Number").Index
            idxMacro = .item("Macrophase").Index
            idxMicro = .item("Microphase").Index
            idxMat = .item("Material").Index
            idxMach = .item("Machine").Index
        End With
        
        For i = 1 To UBound(arrDest, 1)
            If CStr(arrDest(i, idxProd)) = selectedProduct Then
                ' Create a unique string key
                key = selectedProduct & "|" & _
                      arrDest(i, idxMacro) & "|" & _
                      arrDest(i, idxMicro) & "|" & _
                      arrDest(i, idxMat) & "|" & _
                      arrDest(i, idxMach)
                dictExisting(key) = True
            End If
        Next i
    End If
    
    ' --- PHASE 2: READ SOURCE & FILTER ---
    If tblRoutinesDB.DataBodyRange Is Nothing Then GoTo Cleanup
    arrSource = tblRoutinesDB.DataBodyRange.Value
    
    ' Get Source Column Indices
    Dim srcCols As Object
    Set srcCols = CreateObject("Scripting.Dictionary")
    For i = 1 To tblRoutinesDB.ListColumns.Count
        srcCols(tblRoutinesDB.headerRowRange.Cells(1, i).Value) = i
    Next i
    
    ' Temporary storage for filtered rows
    ReDim arrSort(1 To UBound(arrSource, 1))
    countNew = 0
    
    Dim colPlantIdx As Long, colSortIdx As Long
    colPlantIdx = srcCols("Plant")
    colSortIdx = srcCols("Sort Order")
    
    For i = 1 To UBound(arrSource, 1)
        ' Filter by Plant
        If Trim(arrSource(i, colPlantIdx)) = selectedPlant Then
            
            ' Generate Key to check against existing
            ' Note: We need to look up source column indices to build the key
            key = selectedProduct & "|" & _
                  arrSource(i, srcCols("Macrophase")) & "|" & _
                  arrSource(i, srcCols("Microphase")) & "|" & _
                  arrSource(i, srcCols("Material")) & "|" & _
                  arrSource(i, srcCols("Machine"))
            
            If Not dictExisting.exists(key) Then
                countNew = countNew + 1
                ' Store the whole row index and the sort order in an array
                arrSort(countNew) = Array(arrSource(i, colSortIdx), i)
            End If
        End If
    Next i
    
    If countNew = 0 Then
        MsgBox "No new routines to add.", vbInformation
        GoTo Cleanup
    End If
    
    ' --- PHASE 3: SORT NEW DATA (Memory Bubble Sort) ---
    ' Sorting the small array of indices is faster than sorting the whole table
    Dim x As Long, y As Long
    For x = 1 To countNew - 1
        For y = x + 1 To countNew
            If arrSort(x)(0) > arrSort(y)(0) Then
                tempArr = arrSort(x)
                arrSort(x) = arrSort(y)
                arrSort(y) = tempArr
            End If
        Next y
    Next x
    
    ' --- PHASE 4: CONSTRUCT DESTINATION ARRAYS ---
    ' We need to prepare data for specific columns to write in bulk.
    ' We cannot simply write the whole 2D array because of Formulas.
    
    ' Add the rows first
    ' Using Resize is faster than ListRows.Add looped
    Dim startRow As Long
    Dim addedRange As Range
    
    ' Trick: Resize the listobject to accommodate new rows.
    ' Formulas will autofill automatically here!
    With tblSelectedRoutines
        If .DataBodyRange Is Nothing Then
            startRow = 1
            ' Insert first row safely
             .ListRows.Add
             If countNew > 1 Then .Resize .Range.Resize(.Range.Rows.Count + (countNew - 1))
        Else
            startRow = .ListRows.Count + 1
            .Resize .Range.Resize(.Range.Rows.Count + countNew)
        End If
    End With
    
    ' Now we construct arrays for each column we want to write
    ' This is much faster than writing cell by cell
    
    Dim itemMap As Variant
    Dim sourceRowIdx As Long
    Dim colDestName As String, colSourceName As String
    Dim targetCol As ListColumn
    Dim bulkArr() As Variant
    
    ' Loop through our mapping list
    For Each itemMap In mappingList
        colSourceName = itemMap(0)
        colDestName = itemMap(1)
        
        ' Check if columns exist
        If srcCols.exists(colSourceName) Then
            On Error Resume Next
            Set targetCol = tblSelectedRoutines.ListColumns(colDestName)
            On Error GoTo 0
            
            If Not targetCol Is Nothing Then
                ReDim bulkArr(1 To countNew, 1 To 1)
                
                ' Fill array from sorted source indices
                For i = 1 To countNew
                    sourceRowIdx = arrSort(i)(1) ' Get original row index
                    bulkArr(i, 1) = arrSource(sourceRowIdx, srcCols(colSourceName))
                Next i
                
                ' Dump array into the specific column of the NEW rows
                targetCol.DataBodyRange.Cells(startRow, 1).Resize(countNew, 1).Value = bulkArr
            End If
        End If
    Next itemMap
    
    ' --- PHASE 5: FILL STATIC VALUES (Product, Type, Defaults) ---
    ' Product Number
    ReDim bulkArr(1 To countNew, 1 To 1)
    For i = 1 To countNew: bulkArr(i, 1) = selectedProduct: Next i
    tblSelectedRoutines.ListColumns("Product Number").DataBodyRange.Cells(startRow, 1).Resize(countNew, 1).Value = bulkArr
    
    ' Product Type
    For i = 1 To countNew: bulkArr(i, 1) = productType: Next i
    tblSelectedRoutines.ListColumns("Product Type").DataBodyRange.Cells(startRow, 1).Resize(countNew, 1).Value = bulkArr
    
    ' Number of Setups (Default 1)
    For i = 1 To countNew: bulkArr(i, 1) = 1: Next i
    tblSelectedRoutines.ListColumns("Number of Setups").DataBodyRange.Cells(startRow, 1).Resize(countNew, 1).Value = bulkArr

    ' --- CLEANUP ---
Cleanup:
    Call SortSelectedRoutingByProduct ' Assuming this exists elsewhere
    Call SpeedOff
    MsgBox "Routines copied successfully.", vbInformation
    
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
    Dim dia As String, dimm As String, bundles As Variant
    Dim sortOrder As Variant ' Wird aus ListView gelesen, aber nicht geschrieben
    Dim r As ListRow, exists As Boolean
    Dim opsHeader As String
    
    ' Als Double für numerische Stabilität deklariert
    Dim trVal As Double
    Dim teVal As Double
    Dim opsVal As Double
    
    ' === START MAPPING ===
    ' Variablen zum Speichern ALLER Spaltenindizes
    Dim idxProd As Long, idxComp As Long, idxMac As Long, idxMicro As Long
    Dim idxMat As Long, idxMachine As Long, idxDia As Long, idxDimm As Long
    Dim idxWorkCenter As Long, idxPlant As Long, idxProdType As Long
    Dim idxTR As Long, idxTE As Long, idxOps As Long, idxSetups As Long
    ' === END MAPPING ===

    Set wsDest = ThisWorkbook.Sheets("2. Routines")
    Set loDest = wsDest.ListObjects("SelectedRoutines")

    mac = Me.MacrophaseSelect.Value
    selectedComponent = Me.cmbComponentSelect.Value
    selectedProduct = Trim$(wsDest.Range("D6").Value)

    ' ListView-Werte abrufen
    micro = Trim$(li.Text)
    material = Trim$(li.SubItems(1))
    machine = Trim$(li.SubItems(2))
    dia = Trim$(li.SubItems(3))
    dimm = Trim$(li.SubItems(4))
    sortOrder = li.SubItems(8) ' Gelesen, falls benötigt, aber nicht geschrieben

    ' Sichere numerische Konvertierung
    trVal = SafeCDbl(li.SubItems(5))
    teVal = SafeCDbl(li.SubItems(6))
    opsVal = SafeCDbl(li.SubItems(7))
    
    ' === START MAPPING ===
    ' Spaltenindizes VOR der Schleife abrufen, basierend auf Ihrer exakten Liste
    On Error Resume Next
    idxProd = loDest.ListColumns("Product Number").Index
    idxComp = loDest.ListColumns("Component").Index
    idxMac = loDest.ListColumns("Macrophase").Index
    idxMicro = loDest.ListColumns("Microphase").Index
    'idxNewCol = loDest.ListColumns("YourExcelHeaderName").Index
    idxMat = loDest.ListColumns("Material").Index
    idxMachine = loDest.ListColumns("Machine").Index
    idxDia = loDest.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index ' 2 Leerzeichen
    idxDimm = loDest.ListColumns("Wire/component dimensions  (mm)").Index ' 2 Leerzeichen
    idxWorkCenter = loDest.ListColumns("Work Center Code").Index
    idxPlant = loDest.ListColumns("Plant").Index
    idxProdType = loDest.ListColumns("Product Type").Index
    idxTR = loDest.ListColumns("tr").Index
    idxTE = loDest.ListColumns("te").Index
    opsHeader = "Number of operations" ' Exakter Name
    idxOps = loDest.ListColumns(opsHeader).Index
    idxSetups = loDest.ListColumns("Number of Setups").Index
    On Error GoTo 0

    ' Überprüfen, ob alle Spalten gefunden wurden
    If idxProd = 0 Or idxComp = 0 Or idxMac = 0 Or idxMicro = 0 Or _
       idxMat = 0 Or idxMachine = 0 Or idxDia = 0 Or idxDimm = 0 Or _
       idxWorkCenter = 0 Or idxPlant = 0 Or idxProdType = 0 Or _
       idxTR = 0 Or idxTE = 0 Or idxOps = 0 Or idxSetups = 0 Then
        
        MsgBox "Fehler: Kritische Spalten in 'SelectedRoutines' nicht gefunden." & vbCrLf & _
               "Bitte prüfen Sie Überschriften wie 'Number of operations' oder 'Wire/cable...'.", vbCritical
        Exit Sub ' Sicher beenden
    End If
    ' === END MAPPING ===

    ' Versuchen, eine vorhandene passende Zeile zu finden
    exists = False
    For Each r In loDest.ListRows
        ' Verwenden Sie die Indexvariablen
        If Trim$(r.Range(idxProd).Value) = selectedProduct _
        And Trim$(r.Range(idxMac).Value) = mac _
        And Trim$(r.Range(idxMicro).Value) = micro _
        And Trim$(r.Range(idxMat).Value) = material _
        And Trim$(r.Range(idxMachine).Value) = machine _
        And Trim$(r.Range(idxDia).Value) = dia _
        And Trim$(r.Range(idxDimm).Value) = dimm _
        And Trim$(r.Range(idxComp).Value) = selectedComponent Then
            
            ' UPDATE: Vorhandene Zeile aktualisieren
            If micro = "Bunching" Then
                ' Die Formel muss [Work Center Code] usw. referenzieren, stellen Sie sicher, dass diese Spalten existieren
                r.Range(idxOps).Formula = "=SUMIFS(" & _
                    "SelectedRoutines[" & opsHeader & "]," & _
                    "SelectedRoutines[Macrophase],""Cutting""," & _
                    "SelectedRoutines[Machine],[@Machine]," & _
                    "SelectedRoutines[Microphase],""<>Bunching""," & _
                    "SelectedRoutines[Work Center Code],[@Work Center Code]," & _
                    "SelectedRoutines[te],""<>0""," & _
                    "SelectedRoutines[Product Number],[@Product Number])"
            Else
                ' Wert als echten Double schreiben (oder leer, wenn 0)
                If idxOps > 0 Then r.Range(idxOps).Value = IIf(opsVal = 0, vbNullString, opsVal)
            End If
            exists = True
            Exit For
        End If
    Next r

    If exists Then Exit Sub

    ' ERSTELLEN: Neue Zeile
    Dim newRow As ListRow
    If loDest.ListRows.Count = 1 And loDest.ListRows(1).Range.Cells(1).Value = "" Then
        Set newRow = loDest.ListRows(1)
    Else
        Set newRow = loDest.ListRows.Add
    End If

    ' Erforderliche Felder (Verwendung der Indizes)
    newRow.Range(idxPlant).Value = wsDest.Range("D5").Value
    newRow.Range(idxMac).Value = mac
    newRow.Range(idxMicro).Value = micro
    newRow.Range(idxMat).Value = material
    newRow.Range(idxMachine).Value = machine
    newRow.Range(idxDia).Value = dia
    newRow.Range(idxDimm).Value = dimm
    
    ' Werte als echte Doubles schreiben (oder leer, wenn 0)
    newRow.Range(idxTR).Value = IIf(trVal = 0, vbNullString, trVal)
    newRow.Range(idxTE).Value = IIf(teVal = 0, vbNullString, teVal)
    
    ' ENTFERNT: newRow.Range(loDest.ListColumns("Sort Order").Index).Value = sortOrder
    
    newRow.Range(idxSetups).Value = 1
    newRow.Range(idxWorkCenter).Value = workCenterCodeTemp
    If selectedComponent <> "" Then newRow.Range(idxComp).Value = selectedComponent
    newRow.Range(idxProd).Value = selectedProduct
    newRow.Range(idxProdType).Value = wsDest.Range("D8").Value

    ' Bunching-Sonderfall
    If micro = "Bunching" Then
        bundles = InputBox("Enter the number of bundles for the 'Bunching' operation:", "Input Bundles")
        ' ENTFERNT: Code zum Schreiben von 'bundles', da Spalte nicht in der Liste
        ' If IsNumeric(bundles) And Val(bundles) > 0 Then
        '    On Error Resume Next
        '    newRow.Range(loDest.ListColumns("Bundles").Index).Value = bundles
        '    On Error GoTo 0
        ' End If
        
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
        ' Wert als echten Double schreiben (oder leer, wenn 0)
        If idxOps > 0 Then newRow.Range(idxOps).Value = IIf(opsVal = 0, vbNullString, opsVal)
    End If

End Sub




'''
' Konvertiert einen Variant/String-Wert sicher in einen Double-Wert.
' Behandelt internationale Dezimaltrennzeichen.
' Leere oder ungültige Werte werden als 0.0 zurückgegeben.
'''
Private Function SafeCDbl(ByVal inputValue As Variant) As Double
    Dim cleanValue As String
    Dim decSep As String
    
    ' Standardwert
    SafeCDbl = 0#
    
    ' Nichts zu konvertieren
    If IsEmpty(inputValue) Or inputValue = "" Then Exit Function

    ' System-Dezimaltrennzeichen abrufen
    decSep = application.International(xlDecimalSeparator)
    
    ' Eingabe in einen String umwandeln
    cleanValue = CStr(inputValue)
    
    ' 1. Falschen Trenner durch den richtigen ersetzen
    If decSep = "." Then
        cleanValue = Replace(cleanValue, ",", ".")
    Else
        cleanValue = Replace(cleanValue, ".", ",")
    End If
    
    ' 2. Auf Numerisch prüfen, bevor konvertiert wird
    If IsNumeric(cleanValue) Then
        SafeCDbl = CDbl(cleanValue)
    End If
End Function

' Helper function to compare two values safely
Private Function ValuesMatch(val1 As Variant, val2 As Variant) As Boolean
    Dim s1 As String
    Dim s2 As String
    
    ' 1. Convert to string and Lowercase
    s1 = LCase(CStr(val1 & ""))
    s2 = LCase(CStr(val2 & ""))
    
    ' 2. CLEANUP using the Allowlist (Fixes the Ghost ?)
    s1 = CleanString(s1)
    s2 = CleanString(s2)
    
    ' 3. Remove spaces for tighter comparison (optional but recommended)
    s1 = Replace(s1, " ", "")
    s2 = Replace(s2, " ", "")
    
    ' 4. Direct Comparison
    If s1 = s2 Then
        ValuesMatch = True
        Exit Function
    Else
        ' Optional: Handle numeric equivalency (e.g. "10" vs "10.0")
        If IsNumeric(s1) And IsNumeric(s2) Then
            If val(s1) = val(s2) Then ValuesMatch = True
        End If
    End If
End Function

' Helper: Only keeps characters we definitely want (A-Z, 0-9, and standard symbols)
Private Function CleanString(ByVal txt As String) As String
    Dim i As Long
    Dim char As String
    Dim result As String
    Dim code As Long
    
    For i = 1 To Len(txt)
        char = Mid(txt, i, 1)
        code = AscW(char) ' Use AscW for Unicode support
        
        Select Case code
            Case 48 To 57   ' 0-9
                result = result & char
            Case 65 To 90   ' A-Z
                result = result & char
            Case 97 To 122  ' a-z
                result = result & char
            Case 44, 46, 45, 47, 60, 62, 40, 41, 32 ' Standard symbols: , . - / < > ( ) [Space]
                result = result & char
            ' Add any specific German chars if needed (Ä=196, etc),
            ' but usually Dimensions don't need them.
            Case Else
                ' Do nothing - this strips the Ghost Character!
        End Select
    Next i
    
    CleanString = result
End Function
