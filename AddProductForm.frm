VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddProductForm 
   Caption         =   "Add new final product"
   ClientHeight    =   4995
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   8385.001
   OleObjectBlob   =   "AddProductForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "AddProductForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' USERFORM EVENT: txtDescription_Change
' ==========================================================================
' PURPOSE: Provides real-time feedback to the user on the length of the
'          product description, changing the label color to red if it
'          exceeds the 40-character limit.
' ==========================================================================
Private Sub txtDescription_Change()
    Dim charCount As Long
    charCount = Len(Me.txtDescription.Value)

    ' Update the character count label
    Me.lblCharCount.Caption = "Characters: " & charCount

    ' Change the label's font color to indicate if the length is valid
    If charCount > 40 Then
        Me.lblCharCount.ForeColor = RGB(255, 0, 0) ' Red for invalid length
    Else
        Me.lblCharCount.ForeColor = RGB(0, 0, 0)   ' Black for valid length
    End If
End Sub

' ==========================================================================
' USERFORM EVENT: UserForm_Initialize
' ==========================================================================
' PURPOSE: Populates the product types ListView when the form is first opened.
'          Includes error handling to ensure the required sheet and table exist.
' ==========================================================================
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler

    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim productTypeCell As Range

    ' --- Pre-computation/Pre-flight Checks ---
    ' Verify that the required worksheet and table exist before proceeding.
    If Not SheetExists("Global Variables") Then
        MsgBox "Required sheet 'Global Variables' not found. Please check the workbook.", vbCritical, "Initialization Error"
        Unload Me
        Exit Sub
    End If

    Set ws = ThisWorkbook.Sheets("Global Variables")

    If Not TableExists(ws, "ProductTypes") Then
        MsgBox "Required table 'ProductTypes' not found on the 'Global Variables' sheet.", vbCritical, "Initialization Error"
        Unload Me
        Exit Sub
    End If

    Set tbl = ws.ListObjects("ProductTypes")

    ' --- ListView Population ---
    ' Clear any existing items and headers to prevent duplication.
    With Me.lvwProductTypes
        .ListItems.Clear
        .columnHeaders.Clear
        .columnHeaders.Add Text:="Product Type", Width:=120 ' Set header and width
    End With

    ' Loop through the "ProductType" column and add each non-empty cell to the ListView.
    For Each productTypeCell In tbl.ListColumns("ProductType").DataBodyRange.Cells
        If Trim(productTypeCell.Value) <> "" Then
            Me.lvwProductTypes.ListItems.Add Text:=CStr(productTypeCell.Value)
        End If
    Next productTypeCell

    ' Inform the user if no product types were loaded.
    If Me.lvwProductTypes.ListItems.Count = 0 Then
        MsgBox "No product types found in the 'ProductTypes' table.", vbInformation, "No Data"
    End If

    Exit Sub

ErrorHandler:
    ' General error handler for unexpected issues during initialization.
    MsgBox "An unexpected error occurred while initializing the form:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, vbCritical, "Fatal Error"
    Unload Me
End Sub

' ==========================================================================
' USERFORM EVENT: btnAdd_Click
' ==========================================================================
' PURPOSE: Validates user input and adds a new product to the "FinalProductList" table.
'          This sub is the core of the form's functionality and includes extensive
'          validation and error handling.
' ==========================================================================
Private Sub btnAdd_Click()
    On Error GoTo ErrorHandler

    ' --- Variable Declaration ---
    Dim wsProducts As Worksheet
    Dim tblProducts As ListObject
    Dim newRow As ListRow
    Dim product As String, description As String, productType As String
    Dim batchSize As Double, aoq As Double
    Dim existingProduct As Range
    Dim isFirstMeaningfulProduct As Boolean
    Dim meaningfulRows As Long

    ' --- Pre-computation/Pre-flight Checks ---
    If Not SheetExists("Final Products") Then
        MsgBox "Required sheet 'Final Products' not found. Please check the workbook.", vbCritical, "Operation Error"
        Exit Sub
    End If
    Set wsProducts = ThisWorkbook.Sheets("Final Products")
    If Not TableExists(wsProducts, "FinalProductList") Then
        MsgBox "Required table 'FinalProductList' not found on the 'Final Products' sheet.", vbCritical, "Operation Error"
        Exit Sub
    End If
    
    ' --- Input Validation ---
    If Trim(Me.txtProduct.Value) = "" Or Trim(Me.txtDescription.Value) = "" Or _
       Trim(Me.txtBatchSize.Value) = "" Or Trim(Me.txtAOQ.Value) = "" Then
        MsgBox "All fields are required. Please fill in all the information.", vbExclamation, "Missing Information"
        Exit Sub
    End If
    If Len(Trim(Me.txtDescription.Value)) > 40 Then
        MsgBox "The 'Product Description' cannot exceed 40 characters.", vbExclamation, "Invalid Input"
        Me.txtDescription.SetFocus
        Exit Sub
    End If
    If IsNumeric(Me.txtBatchSize.Value) Then
        batchSize = CDbl(Me.txtBatchSize.Value)
        If batchSize <= 0 Then
            MsgBox "Please enter a numeric value greater than 0 for 'Batch Size'.", vbExclamation, "Invalid Input"
            Me.txtBatchSize.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "The value entered for 'Batch Size' is not a valid number.", vbExclamation, "Invalid Input"
        Me.txtBatchSize.SetFocus
        Exit Sub
    End If
    If IsNumeric(Me.txtAOQ.Value) Then
        aoq = CDbl(Me.txtAOQ.Value)
        If aoq <= 0 Then
            MsgBox "Please enter a numeric value greater than 0 for 'Annual Order Quantity'.", vbExclamation, "Invalid Input"
            Me.txtAOQ.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "The value entered for 'Annual Order Quantity' is not a valid number.", vbExclamation, "Invalid Input"
        Me.txtAOQ.SetFocus
        Exit Sub
    End If
    If Me.lvwProductTypes.selectedItem Is Nothing Then
        MsgBox "Please select a 'Product Type' from the list.", vbExclamation, "Selection Required"
        Exit Sub
    End If
    
    ' --- Data Processing ---
    product = Trim(Me.txtProduct.Value)
    description = Trim(Me.txtDescription.Value)
    productType = Me.lvwProductTypes.selectedItem.Text
    Set tblProducts = wsProducts.ListObjects("FinalProductList")

    On Error Resume Next
    Set existingProduct = tblProducts.ListColumns("Product Number").DataBodyRange.Find(What:=product, LookIn:=xlValues, LookAt:=xlWhole)
    On Error GoTo ErrorHandler
    
    If Not existingProduct Is Nothing Then
        MsgBox "The Product '" & product & "' already exists. Please enter a unique Product Number.", vbExclamation, "Duplicate Entry"
        Me.txtProduct.SetFocus
        Exit Sub
    End If

    ' Determine the number of meaningful rows before adding the new one.
    meaningfulRows = application.WorksheetFunction.CountA(tblProducts.ListColumns("Product Number").DataBodyRange)
    isFirstMeaningfulProduct = (meaningfulRows = 0)

    ' --- Add Data to Table ---
    If tblProducts.ListRows.Count = 1 And application.WorksheetFunction.CountA(tblProducts.ListRows(1).Range) = 0 Then
        Set newRow = tblProducts.ListRows(1)
    Else
        Set newRow = tblProducts.ListRows.Add(AlwaysInsert:=True)
    End If

    With newRow
        .Range(tblProducts.ListColumns("Product Number").Index).Value = product
        .Range(tblProducts.ListColumns("Product Description").Index).Value = description
        .Range(tblProducts.ListColumns("Product Type").Index).Value = productType
        .Range(tblProducts.ListColumns("Batch").Index).Value = batchSize
        .Range(tblProducts.ListColumns("AOQ").Index).Value = aoq
        .Range(tblProducts.ListColumns("ProductNumberText").Index).Formula = "=[@[Product Number]]"
    End With

    ' --- Post-Addition Tasks ---
    ' RESTORED: This call to an external sub was missing and is critical for the first product's setup.
    If isFirstMeaningfulProduct Then
        Call AddFirstProduct
        ThisWorkbook.Sheets("1. BOM Definition").Range("F11").Value = product
        ThisWorkbook.Sheets("2. Routines").Range("D6").Value = product
    End If
    
    ' Clear form fields for next entry
    Me.txtProduct.Value = ""
    Me.txtDescription.Value = ""
    Me.txtBatchSize.Value = ""
    Me.txtAOQ.Value = ""
    Me.lvwProductTypes.selectedItem = Nothing

    MsgBox "Product '" & product & "' was added successfully!", vbInformation, "Success"

    ' Update related dropdowns and validation statuses.
    UpdateProductDropdown
    UpdateRoutineDropdown
    
    With ThisWorkbook.Sheets("3. Clarification Validation")
        .Range("J7").Value = "New Product / component added. Please validate the RFQ"
        .Range("J7").Interior.Color = RGB(255, 255, 0) ' Yellow
        .Range("O14:O24").ClearContents
        .Range("O14:O24").Interior.ColorIndex = xlNone
    End With

    ' Conditionally show sheets or buttons based on product type.
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
    
    Call AddMaterialPreparingRoutineIfNeeded(product, productType)
    
    Unload Me ' Close the form.

    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred while adding the product:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, vbCritical, "Fatal Error"
End Sub

' ==========================================================================
' USERFORM EVENT: btnCancel_Click
' ==========================================================================
' PURPOSE: Closes the user form without saving any data.
' ==========================================================================
Private Sub btnCancel_Click()
    Unload Me
End Sub

' ==========================================================================
' PUBLIC SUBROUTINE: AddMaterialPreparingRoutineIfNeeded
' ==========================================================================
' PURPOSE: Adds a default "Material preparing" routine for a new product if the
'          selected plant is "1410" or "1420". It will use the first row of the
'          destination table if it's empty, otherwise it adds a new row.
' ==========================================================================
Public Sub AddMaterialPreparingRoutineIfNeeded(ByVal productNumber As String, ByVal productType As String)
    On Error GoTo ErrorHandler

    ' --- Pre-computation/Pre-flight Checks ---
    If Not SheetExists("1. BOM Definition") Or Not SheetExists("RoutinesDB") Or Not SheetExists("2. Routines") Then
        MsgBox "One or more required sheets for routine creation are missing. Aborting operation.", vbCritical, "Missing Sheet"
        Exit Sub
    End If
    Dim wsRoutinesDB As Worksheet: Set wsRoutinesDB = ThisWorkbook.Sheets("RoutinesDB")
    Dim wsSelectedRoutines As Worksheet: Set wsSelectedRoutines = ThisWorkbook.Sheets("2. Routines")
    If Not TableExists(wsRoutinesDB, "RoutinesDB") Or Not TableExists(wsSelectedRoutines, "SelectedRoutines") Then
        MsgBox "One or more required tables for routine creation are missing. Aborting operation.", vbCritical, "Missing Table"
        Exit Sub
    End If

    ' --- Main Logic ---
    Dim selectedPlant As String
    selectedPlant = Trim(ThisWorkbook.Sheets("1. BOM Definition").Range("C9").Value)

    If selectedPlant <> "1410" And selectedPlant <> "1420" Then
        Exit Sub
    End If

    Dim tblRoutinesDB As ListObject: Set tblRoutinesDB = wsRoutinesDB.ListObjects("RoutinesDB")
    Dim tblSelected As ListObject: Set tblSelected = wsSelectedRoutines.ListObjects("SelectedRoutines")
    Dim routineRow As ListRow, newRow As ListRow
    Dim routineFound As Boolean: routineFound = False
    
    Dim trValue As Variant, teValue As Variant
    Dim trConverted As Double, teConverted As Double

    For Each routineRow In tblRoutinesDB.ListRows
        If Trim(routineRow.Range(tblRoutinesDB.ListColumns("Plant").Index).Value) = selectedPlant And _
           Trim(routineRow.Range(tblRoutinesDB.ListColumns("Macrophase").Index).Value) = "Stock" And _
           Trim(routineRow.Range(tblRoutinesDB.ListColumns("Microphase").Index).Value) = "Material preparing" Then
            
            ' --- UPDATED: Logic to handle placeholder row ---
            ' If the table has one row and the "Product Number" cell is empty, use that row.
            ' Otherwise, add a new row to the table.
            If tblSelected.ListRows.Count = 1 And IsEmpty(tblSelected.ListRows(1).Range(tblSelected.ListColumns("Product Number").Index).Value) Then
                Set newRow = tblSelected.ListRows(1)
            Else
                Set newRow = tblSelected.ListRows.Add(AlwaysInsert:=True)
            End If
            
            ' --- Safe Conversion for tr/te values ---
            trValue = routineRow.Range(tblRoutinesDB.ListColumns("tr").Index).Value
            teValue = routineRow.Range(tblRoutinesDB.ListColumns("te").Index).Value
            
            If IsNumeric(trValue) Then
                trConverted = CDbl(trValue)
            Else
                trConverted = 0
            End If
            
            If IsNumeric(teValue) Then
                teConverted = CDbl(teValue)
            Else
                teConverted = 0
            End If

            ' Assign values cell-by-cell for maximum safety and clarity.
            With newRow.Range
                .Cells(1, tblSelected.ListColumns("Plant").Index).Value = selectedPlant
                .Cells(1, tblSelected.ListColumns("Product Number").Index).Value = productNumber
                .Cells(1, tblSelected.ListColumns("Product Type").Index).Value = productType
                .Cells(1, tblSelected.ListColumns("Macrophase").Index).Value = "Stock"
                .Cells(1, tblSelected.ListColumns("Microphase").Index).Value = "Material preparing"
                .Cells(1, tblSelected.ListColumns("Material").Index).Value = routineRow.Range(tblRoutinesDB.ListColumns("Material").Index).Value
                .Cells(1, tblSelected.ListColumns("Machine").Index).Value = routineRow.Range(tblRoutinesDB.ListColumns("Machine").Index).Value
                .Cells(1, tblSelected.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).Value = routineRow.Range(tblRoutinesDB.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).Value
                .Cells(1, tblSelected.ListColumns("Wire/component dimensions  (mm)").Index).Value = routineRow.Range(tblRoutinesDB.ListColumns("Wire/component dimensions  (mm)").Index).Value
                .Cells(1, tblSelected.ListColumns("Work Center Code").Index).Value = routineRow.Range(tblRoutinesDB.ListColumns("Work Center Code").Index).Value
                .Cells(1, tblSelected.ListColumns("tr").Index).Value = trConverted
                .Cells(1, tblSelected.ListColumns("te").Index).Value = teConverted
                .Cells(1, tblSelected.ListColumns("Number of Operations").Index).Value = 1
                .Cells(1, tblSelected.ListColumns("Number of Setups").Index).Value = 1
                .Cells(1, tblSelected.ListColumns("Sort Order").Index).Value = routineRow.Range(tblRoutinesDB.ListColumns("Sort Order").Index).Value
            End With
            
            routineFound = True
            Exit For
        End If
    Next routineRow

    If Not routineFound Then
        MsgBox "Could not find a default 'Material preparing' routine in RoutinesDB for plant " & selectedPlant & ".", vbExclamation, "Routine Not Found"
    End If

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in 'AddMaterialPreparingRoutineIfNeeded':" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, vbCritical, "Routine Error"
End Sub

' ==========================================================================
' HELPER FUNCTION: SheetExists
' ==========================================================================
' PURPOSE: Checks if a worksheet with the given name exists in the workbook.
' RETURNS: Boolean (True if the sheet exists, False otherwise).
' ==========================================================================
Private Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

' ==========================================================================
' HELPER FUNCTION: TableExists
' ==========================================================================
' PURPOSE: Checks if a table (ListObject) with the given name exists on a specified worksheet.
' RETURNS: Boolean (True if the table exists, False otherwise).
' ==========================================================================
Private Function TableExists(ByVal ws As Worksheet, ByVal tableName As String) As Boolean
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects(tableName)
    On Error GoTo 0
    TableExists = Not tbl Is Nothing
End Function

