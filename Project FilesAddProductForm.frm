VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddProductForm 
   Caption         =   "Add new final product"
   ClientHeight    =   4995
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8385.001
   OleObjectBlob   =   "Project FilesAddProductForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "AddProductForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub txtDescription_Change()
    Dim charCount As Long
    charCount = Len(Me.txtDescription.value)

    ' Update the label text
    Me.lblCharCount.Caption = "Characters: " & charCount

    ' Change color based on length
    If charCount > 40 Then
        Me.lblCharCount.ForeColor = RGB(255, 0, 0) ' Red
    Else
        Me.lblCharCount.ForeColor = RGB(0, 0, 0)   ' Black
    End If
End Sub



Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim productTypeRange As Range
    Dim productTypeCell As Range

    ' Reference the Global Variables sheet and ProductTypes table
    Set ws = ThisWorkbook.Sheets("Global Variables")
    Set tbl = ws.ListObjects("ProductTypes")

    ' Set the range for ProductType column
    Set productTypeRange = tbl.ListColumns("ProductType").DataBodyRange

    ' Clear existing items and headers
    Me.lvwProductTypes.ListItems.Clear
    Me.lvwProductTypes.columnHeaders.Clear ' Ensure previous headers are cleared

    ' Add a single column header for the ListView
    Me.lvwProductTypes.columnHeaders.Add , , "Product Type", 100 ' Adjust width as necessary

    ' Populate the ListView for product types
    For Each productTypeCell In productTypeRange
        If productTypeCell.value <> "" Then
            With Me.lvwProductTypes.ListItems.Add
                .Text = productTypeCell.value ' Add product type to the ListView
            End With
        End If
    Next productTypeCell

    ' Alert if no product types were found
    If Me.lvwProductTypes.ListItems.Count = 0 Then
        MsgBox "No product types found. Please check the ProductTypes table.", vbInformation
    End If
End Sub
Private Sub btnAdd_Click()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim newRow As ListRow
    Dim product As String
    Dim description As String
    Dim productType As String
    Dim batchSize As Double
    Dim aoq As Double
    Dim existingProduct As Range
    Dim isFirstMeaningfulProduct As Boolean
    Dim meaningfulRows As Long

    ' Get the values from the text boxes
    product = Trim(Me.txtProduct.value)
    description = Trim(Me.txtDescription.value)

    ' Validate numeric input for Batch Size
    If IsNumeric(Me.txtBatchSize.value) And CDbl(Me.txtBatchSize.value) > 0 Then
        batchSize = CDbl(Me.txtBatchSize.value)
    Else
        MsgBox "Please enter a valid numeric value greater than 0 for Batch Size.", vbExclamation
        Me.txtBatchSize.SetFocus
        Exit Sub
    End If

    ' Validate numeric input for Annual Order Quantity
    If IsNumeric(Me.txtAOQ.value) And CDbl(Me.txtAOQ.value) > 0 Then
        aoq = CDbl(Me.txtAOQ.value)
    Else
        MsgBox "Please enter a valid numeric value greater than 0 for Annual Order Quantity.", vbExclamation
        Me.txtAOQ.SetFocus
        Exit Sub
    End If

    ' Check if a product type is selected
    If Me.lvwProductTypes.selectedItem Is Nothing Then
        MsgBox "Please select an Product type from the list.", vbExclamation
        Exit Sub
    End If

    ' Get the selected product type from the ListView
    productType = Me.lvwProductTypes.selectedItem.Text

    ' Validate Product and description inputs
    If product = "" Or description = "" Then
        MsgBox "Please fill in all fields.", vbExclamation
        Exit Sub
    End If

    ' Set the FinalProductList table
    Set ws = ThisWorkbook.Sheets("Final Products")
    Set tbl = ws.ListObjects("FinalProductList")

    ' Check if the Product already exists in the FinalProductList table
    On Error Resume Next
    Set existingProduct = tbl.ListColumns("Product Number").DataBodyRange.Find(What:=product, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo 0

    If Not existingProduct Is Nothing Then
        MsgBox "The Product '" & product & "' already exists in the FinalProductList. Please add a unique Product.", vbExclamation
        Exit Sub
    End If

    ' Determine the number of meaningful rows (rows with non-empty Product Number)
    On Error Resume Next
    meaningfulRows = application.WorksheetFunction.CountA(tbl.ListColumns("Product Number").DataBodyRange)
    On Error GoTo 0

    ' If meaningfulRows is 0, this is the first meaningful Product
    isFirstMeaningfulProduct = (meaningfulRows = 0)

    ' Add a new row to the FinalProductList table
    ' Check if the table has exactly one row and if that row's "Product Number" cell is empty.
    If tbl.ListRows.Count = 1 And IsEmpty(tbl.ListRows(1).Range(tbl.ListColumns("Product Number").Index).value) Then
        ' If so, use the existing single row as the target.
        Set newRow = tbl.ListRows(1)
    Else
        ' Otherwise, add a new row as normal.
        Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
    End If

    ' Assign the Product, description, product type, batch size, and AOQ to the new row
    With newRow.Range
        .Cells(1, tbl.ListColumns("Product Number").Index).value = product

        .Cells(1, tbl.ListColumns("Product Description").Index).value = description
        .Cells(1, tbl.ListColumns("Product Type").Index).value = productType
        .Cells(1, tbl.ListColumns("Batch").Index).value = batchSize
        .Cells(1, tbl.ListColumns("AOQ").Index).value = aoq
    End With
    ' Set ProductNumberText formula *after* the row is created
    Dim formulaText As String
    formulaText = "=""&[@[Product Number]]"""
    newRow.Range.Cells(1, tbl.ListColumns("ProductNumberText").Index).Formula = "="""" & [@[Product Number]]"




    ' Check if the first row is empty after adding and delete if necessary
    If application.WorksheetFunction.CountA(tbl.ListRows(1).Range) = 0 Then
        tbl.ListRows(1).Delete
    End If

    ' If this is the first meaningful Product, call AddFirstProduct and set B5
    If isFirstMeaningfulProduct Then
        Call AddFirstProduct

        ' Set F11 in BOM Definition - the selection dropdown to the added Product number
        With ThisWorkbook.Sheets("1. BOM Definition").Range("F11")
            .value = product
        End With
        
        ' Set D6 in Routines - the selection dropdown to the added Product number
        With ThisWorkbook.Sheets("2. Routines").Range("D6")
            .value = product
        End With
        
        
    End If

    ' Clear the text boxes after adding
    Me.txtProduct.value = ""
    Me.txtDescription.value = ""
    Me.txtBatchSize.value = ""
    Me.txtAOQ.value = ""
    Me.lvwProductTypes.selectedItem = Nothing ' Clear selection from ListView

    ' Show success message
    MsgBox "Product added successfully!", vbInformation

    ' Call the UpdateProductDropdown to refresh the dropdown in Product Specification
    UpdateProductDropdown
    UpdateRoutineDropdown

    ' Clear validation after adding a new product
    Dim wsValidation As Worksheet
    Set wsValidation = ThisWorkbook.Sheets("3. Clarification Validation")
    
    ' Update cell C7 on the Product Specification sheet
    With wsValidation.Range("J7")
        .value = "New Product / component added. Please validate the RFQ"
        .Interior.Color = RGB(255, 255, 0) ' Yellow color
    End With
    
    ' Clear the contents and reset the interior color of O14 to O22
    
    With wsValidation.Range("O14:O24")
        .ClearContents ' Clear the cell contents
        .Interior.ColorIndex = xlNone ' Reset the interior color to transparent
    End With
    If LCase(productType) = "chain" Then
        Sheets("Page 1 - Chain RFQ Form").Visible = xlSheetVisible
        Sheets("Page 2 - Chain RFQ Form").Visible = xlSheetVisible
        Sheets("Chain Inner separation").Visible = xlSheetVisible
        ThisWorkbook.Sheets("1. BOM Definition").Shapes("btnOpenChainForm").Visible = True
    End If
    
    If LCase(productType) = "servo" Then
        Sheets("8. Servo calculation").Visible = xlSheetVisible
        ThisWorkbook.Sheets("1. BOM Definition").Shapes("btnOpenServoForm").Visible = True
    End If
    
    'Debug.Print "Calling AddMaterialPreparingRoutineIfNeeded with: " & Product & ", " & productType
    Call AddMaterialPreparingRoutineIfNeeded(product, productType)

    Unload Me
End Sub
Private Sub btnCancel_Click()
    ' Close the form
    Unload Me
End Sub
Public Sub AddMaterialPreparingRoutineIfNeeded(ByVal productNumber As String, ByVal productType As String)
    Dim selectedPlant As String
    selectedPlant = Trim(ThisWorkbook.Sheets("1. BOM Definition").Range("C9").value)

    'Debug.Print "Checking if default material preparing routine is needed for plant: " & selectedPlant
    'Debug.Print "Checking if default material preparing routine is needed for product: " & productNumber

    If selectedPlant <> "1410" And selectedPlant <> "1420" Then
        Debug.Print "Plant is not 1410 or 1420 — skipping default operation."
        Exit Sub
    End If

    Dim wsRoutinesDB As Worksheet, wsSelectedRoutines As Worksheet
    Dim tblRoutinesDB As ListObject, tblSelected As ListObject
    Dim routineRow As ListRow, newRow As ListRow
    Dim found As Boolean: found = False

    Set wsRoutinesDB = ThisWorkbook.Sheets("RoutinesDB")
    Set tblRoutinesDB = wsRoutinesDB.ListObjects("RoutinesDB")
    Set wsSelectedRoutines = ThisWorkbook.Sheets("2. Routines")
    Set tblSelected = wsSelectedRoutines.ListObjects("SelectedRoutines")

    For Each routineRow In tblRoutinesDB.ListRows
        If Trim(routineRow.Range(tblRoutinesDB.ListColumns("Plant").Index).value) = selectedPlant And _
           Trim(routineRow.Range(tblRoutinesDB.ListColumns("Macrophase").Index).value) = "Stock" And _
           Trim(routineRow.Range(tblRoutinesDB.ListColumns("Microphase").Index).value) = "Material preparing" Then

            Set newRow = tblSelected.ListRows.Add
            With newRow.Range
                .Cells(tblSelected.ListColumns("Plant").Index).value = selectedPlant
                .Cells(tblSelected.ListColumns("Product Number").Index).value = productNumber
                .Cells(tblSelected.ListColumns("Product Type").Index).value = productType
                .Cells(tblSelected.ListColumns("Macrophase").Index).value = "Stock"
                .Cells(tblSelected.ListColumns("Microphase").Index).value = "Material preparing"
                .Cells(tblSelected.ListColumns("Material").Index).value = routineRow.Range(tblRoutinesDB.ListColumns("Material").Index).value
                .Cells(tblSelected.ListColumns("Machine").Index).value = routineRow.Range(tblRoutinesDB.ListColumns("Machine").Index).value
                .Cells(tblSelected.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).value = routineRow.Range(tblRoutinesDB.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).value
                .Cells(tblSelected.ListColumns("Wire/component dimensions  (mm)").Index).value = routineRow.Range(tblRoutinesDB.ListColumns("Wire/component dimensions  (mm)").Index).value
                .Cells(tblSelected.ListColumns("Work Center Code").Index).value = routineRow.Range(tblRoutinesDB.ListColumns("Work Center Code").Index).value
                .Cells(tblSelected.ListColumns("tr").Index).value = routineRow.Range(tblRoutinesDB.ListColumns("tr").Index).value
                .Cells(tblSelected.ListColumns("te").Index).value = routineRow.Range(tblRoutinesDB.ListColumns("te").Index).value
                .Cells(tblSelected.ListColumns("Number of Operations").Index).value = 1
                .Cells(tblSelected.ListColumns("Number of Setups").Index).value = 1
                .Cells(tblSelected.ListColumns("Sort Order").Index).value = routineRow.Range(tblRoutinesDB.ListColumns("Sort Order").Index).value
            End With

            Debug.Print "Routine inserted for: " & productNumber & " (Plant " & selectedPlant & ")"
            found = True
            Exit For
        End If
    Next routineRow

    If Not found Then
        Debug.Print "? 'Material preparing' routine not found in RoutinesDB for plant " & selectedPlant
        MsgBox "Could not find 'Material preparing' operation in RoutinesDB for plant " & selectedPlant, vbExclamation
    End If
End Sub


