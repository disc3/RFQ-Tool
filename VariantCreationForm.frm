VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} VariantCreationForm 
   Caption         =   "Create Variants based on an article"
   ClientHeight    =   4860
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   4845
   OleObjectBlob   =   "VariantCreationForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "VariantCreationForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    Dim wsComponents As Worksheet
    Dim tblComponents As ListObject
    Dim productRow As ListRow
    Dim uniqueProducts As Object
    Dim productNumber As String, ProductDescription As String
    Dim rowIndex As Long

    ' Debug: Start of initialization
    Debug.Print "UserForm_Initialize: Start initialization"

    ' Set worksheet and table
    Set wsComponents = ThisWorkbook.Sheets("1. BOM Definition")
    If wsComponents Is Nothing Then
        Err.Raise vbObjectError + 1, "UserForm_Initialize", "Worksheet 'Selected Components MMD' not found."
    End If
    Debug.Print "Worksheet found: " & wsComponents.name

    Set tblComponents = wsComponents.ListObjects("BOMDefinition")
    If tblComponents Is Nothing Then
        Err.Raise vbObjectError + 2, "UserForm_Initialize", "Table 'BOMDefinition' not found in the worksheet."
    End If
    Debug.Print "Table found: " & tblComponents.name

    ' Use a Dictionary to store unique Product numbers
    Set uniqueProducts = CreateObject("Scripting.Dictionary")
    Debug.Print "Dictionary initialized for unique products"

    ' Configure ComboBox for two columns
    cmbBaseProduct.ColumnCount = 2
    cmbBaseProduct.ColumnWidths = "70;150"
    cmbBaseProduct.Clear

    ' Populate ComboBox with unique Product Numbers and Descriptions
    rowIndex = 0
    For Each productRow In tblComponents.ListRows
        On Error Resume Next ' Prevent crash in case of unexpected data
        productNumber = productRow.Range(tblComponents.ListColumns("Product Number").Index).Value
        ProductDescription = productRow.Range(tblComponents.ListColumns("Product Description").Index).Value
        On Error GoTo ErrorHandler

        ' Debug: Print product info
        Debug.Print "Product Number: " & productNumber & ", Description: " & ProductDescription

        ' Add unique Products only
        If productNumber <> "" And Not uniqueProducts.exists(productNumber) Then
            uniqueProducts.Add productNumber, ProductDescription
            cmbBaseProduct.AddItem
            cmbBaseProduct.List(rowIndex, 0) = productNumber
            cmbBaseProduct.List(rowIndex, 1) = ProductDescription
            rowIndex = rowIndex + 1
        End If
    Next productRow

    Debug.Print "UserForm_Initialize: Completed successfully"
    Exit Sub

ErrorHandler:
    ' Handle and log the error
    MsgBox "Error in UserForm_Initialize: " & Err.description & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Procedure: " & Err.Source, vbCritical
    Debug.Print "Error in UserForm_Initialize: " & Err.description & " (Error Number: " & Err.Number & ")"
    Debug.Print "Procedure: " & Err.Source
    Resume Next ' Resume for debugging purposes
End Sub

Private Sub cmbBaseProduct_Change()
    Dim wsComponents As Worksheet
    Dim tblComponents As ListObject
    Dim productRow As ListRow
    Dim selectedProduct As String
    Dim rowIndex As Long

    ' Set worksheet and table
    Set wsComponents = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblComponents = wsComponents.ListObjects("BOMDefinition")

    ' Get the selected Product Number
    If cmbBaseProduct.ListIndex = -1 Then Exit Sub
    selectedProduct = cmbBaseProduct.List(cmbBaseProduct.ListIndex, 0)

    ' Debug: Print the selected Product number
    Debug.Print "Selected Product: " & selectedProduct

    ' Update Product Description
    txtBaseProductDesc.Text = "" ' Clear previous description
    For Each productRow In tblComponents.ListRows
        If productRow.Range(tblComponents.ListColumns("Product Number").Index).Value = selectedProduct Then
            txtBaseProductDesc.Text = productRow.Range(tblComponents.ListColumns("Product Description").Index).Value
            Exit For
        End If
    Next productRow

    ' Populate Variable Component ComboBox with Material and Quantity
    cmbVariableComponent.Clear
    cmbVariableComponent.ColumnCount = 3 ' Display Material, Description, and Quantity
    cmbVariableComponent.ColumnWidths = "70;150;50" ' Adjust column widths

    rowIndex = 0
    For Each productRow In tblComponents.ListRows
        If productRow.Range(tblComponents.ListColumns("Product Number").Index).Value = selectedProduct Then
            ' Add Material, Material Description, and Quantity
            cmbVariableComponent.AddItem
            cmbVariableComponent.List(rowIndex, 0) = productRow.Range(tblComponents.ListColumns("Material").Index).Value
            cmbVariableComponent.List(rowIndex, 1) = productRow.Range(tblComponents.ListColumns("Material Description").Index).Value
            cmbVariableComponent.List(rowIndex, 2) = productRow.Range(tblComponents.ListColumns("Quantity").Index).Value
            rowIndex = rowIndex + 1
        End If
    Next productRow

    ' Clear Variable Component Description TextBox
    txtVariableComponentDesc.Text = ""
    txtOriginalQty.Text = ""
End Sub

Private Sub cmbVariableComponent_Change()
    Dim selectedComponent As String
    Dim selectedDescription As String
    Dim selectedQuantity As Double

    ' Verify selection
    If cmbVariableComponent.ListIndex = -1 Then Exit Sub

    ' Retrieve selected values from the ComboBox
    selectedComponent = cmbVariableComponent.List(cmbVariableComponent.ListIndex, 0) ' Material
    selectedDescription = cmbVariableComponent.List(cmbVariableComponent.ListIndex, 1) ' Material Description
    selectedQuantity = val(cmbVariableComponent.List(cmbVariableComponent.ListIndex, 2)) ' Quantity

    ' Debug: Print selected values
    Debug.Print "Selected Component: " & selectedComponent
    Debug.Print "Selected Description: " & selectedDescription
    Debug.Print "Selected Quantity: " & selectedQuantity

    ' Update the Material Description TextBox
    txtVariableComponentDesc.Text = selectedDescription

    ' Update the quantity text box with the selected component's quantity
    txtOriginalQty.Text = selectedQuantity

    ' Debug: Confirm updates
    Debug.Print "Updated Original Quantity: " & txtOriginalQty.Text
End Sub
Private Sub btnCreateVariants_Click()
    Dim wsComponents As Worksheet, wsProducts As Worksheet
    Dim tblComponents As ListObject, tblFinalProductList As ListObject
    Dim baseProduct As String, baseProductDescription As String
    Dim variableComponent As String, variableDescription As String
    Dim variableRow As ListRow, baseFinalProductListRow As ListRow
    Dim variantRow As ListRow, newProductRow As ListRow
    Dim NumVariants As Long, variantQuantities() As Double
    Dim i As Long
    Dim negativeQuantity As Double
    Dim variantName As String, VariantDescription As String
    Dim sourceCell As Range, targetCell As Range
    Dim j As Long
    Dim existingVariants As Collection
    Dim highestVariant As Long, nextVariant As Long
    Dim productNumber As String
    Dim VariantNames() As String
    
    
    If cmbBaseProduct.ColumnCount >= 2 Then
        baseProductDescription = cmbBaseProduct.List(cmbBaseProduct.ListIndex, 1)
    Else
        baseProductDescription = ""
    End If



    ' Validate Base Product selection
    If cmbBaseProduct.ListIndex = -1 Then
        MsgBox "Please select a base Product.", vbExclamation
        Exit Sub
    End If
    baseProduct = cmbBaseProduct.List(cmbBaseProduct.ListIndex, 0)
    baseProductDescription = cmbBaseProduct.List(cmbBaseProduct.ListIndex, 1)

    ' Validate Variable Component selection
    If cmbVariableComponent.ListIndex = -1 Then
        MsgBox "Please select a variable component.", vbExclamation
        Exit Sub
    End If
    variableComponent = cmbVariableComponent.List(cmbVariableComponent.ListIndex, 0)
    variableDescription = cmbVariableComponent.List(cmbVariableComponent.ListIndex, 1)

    ' Validate Number of Variants
    If Not IsNumeric(txtNumVariants.Value) Or val(txtNumVariants.Value) <= 0 Then
        MsgBox "Enter a valid number of variants.", vbExclamation
        Exit Sub
    End If
    NumVariants = CLng(txtNumVariants.Value)
    ReDim VariantNames(1 To NumVariants)
    

    ' Prompt for quantities
    ReDim variantQuantities(1 To NumVariants)
    For i = 1 To NumVariants
        Do
            variantQuantities(i) = application.InputBox("Enter quantity for Variant " & i & ":", "Quantity Input", Type:=1)
            If IsNumeric(variantQuantities(i)) And variantQuantities(i) > 0 Then Exit Do
            MsgBox "Please enter a valid positive number.", vbExclamation
        Loop
    Next i

    ' Set worksheets and tables
    Set wsComponents = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblComponents = wsComponents.ListObjects("BOMDefinition")
    Set wsProducts = ThisWorkbook.Sheets("Final Products")
    Set tblFinalProductList = wsProducts.ListObjects("FinalProductList")

    ' Find the row with the selected variable component in Components Table
    Dim selectedQuantity As Double
    selectedQuantity = CDbl(txtOriginalQty.Text) ' Use the original quantity shown in the form
    
    For Each variableRow In tblComponents.ListRows
        If variableRow.Range(tblComponents.ListColumns("Product Number").Index).Value = baseProduct _
        And variableRow.Range(tblComponents.ListColumns("Material").Index).Value = variableComponent _
        And variableRow.Range(tblComponents.ListColumns("Quantity").Index).Value = selectedQuantity Then
            Exit For
        End If
    Next variableRow


    If variableRow Is Nothing Then
        MsgBox "Variable component not found in the base Product.", vbExclamation
        Exit Sub
    End If

    ' Retrieve original quantity for the negative row
    negativeQuantity = variableRow.Range(tblComponents.ListColumns("Quantity").Index).Value

    ' Find base Product row in FinalProductList
    For Each baseFinalProductListRow In tblFinalProductList.ListRows
        If baseFinalProductListRow.Range(tblFinalProductList.ListColumns("Product Number").Index).Value = baseProduct Then Exit For
    Next baseFinalProductListRow

    If baseFinalProductListRow Is Nothing Then
        MsgBox "Base Product not found in the Products table.", vbExclamation
        Exit Sub
    End If

    ' ---------------------------
    ' Check for existing variants
    ' ---------------------------
    Set existingVariants = New Collection
    nextVariant = GetNextVariantNumber(baseProduct, tblComponents)


    ' ---------------------------
    ' Create Variants
    ' ---------------------------
    Debug.Print "Highest Variant: " & highestVariant
    Debug.Print "Starting nextVariant at: " & nextVariant

    ManualOverrides.SuppressChangeTracking = True
    For i = 1 To NumVariants
        Debug.Print "Creating variant: " & variantName

        variantName = baseProduct & "-V" & nextVariant
        VariantDescription = baseProductDescription & " | Changed: " & variableComponent & " = " & variantQuantities(i)
        VariantNames(i) = variantName
        ' Row 1: Negative quantity in Components Table
        Set variantRow = tblComponents.ListRows.Add
        For j = 1 To tblComponents.ListColumns.Count
            Set sourceCell = variableRow.Range(j)
            Set targetCell = variantRow.Range(j)
            If tblComponents.ListColumns(j).name = "Product Number" Then
                targetCell.Value = variantName
            ElseIf tblComponents.ListColumns(j).name = "Product Description" Then
                targetCell.Value = VariantDescription
            ElseIf tblComponents.ListColumns(j).name = "Variant of" Then
                targetCell.Value = baseProduct
            ElseIf tblComponents.ListColumns(j).name = "Quantity" Then
                targetCell.Value = -negativeQuantity
            ElseIf Not sourceCell.HasFormula Then
                targetCell.Value = sourceCell.Value
            End If
        Next j

        ' Row 2: Positive quantity in Components Table
        Set variantRow = tblComponents.ListRows.Add
        For j = 1 To tblComponents.ListColumns.Count
            Set sourceCell = variableRow.Range(j)
            Set targetCell = variantRow.Range(j)
            If tblComponents.ListColumns(j).name = "Product Number" Then
                targetCell.Value = variantName
            ElseIf tblComponents.ListColumns(j).name = "Product Description" Then
                targetCell.Value = VariantDescription
            ElseIf tblComponents.ListColumns(j).name = "Variant of" Then
                targetCell.Value = baseProduct
            ElseIf tblComponents.ListColumns(j).name = "Quantity" Then
                targetCell.Value = variantQuantities(i)
            ElseIf Not sourceCell.HasFormula Then
                targetCell.Value = sourceCell.Value
            End If
        Next j

        ' Add new variant to Products Table
        Set newProductRow = tblFinalProductList.ListRows.Add
        For j = 1 To tblFinalProductList.ListColumns.Count
            Set sourceCell = baseFinalProductListRow.Range(j)
            Set targetCell = newProductRow.Range(j)
            If tblFinalProductList.ListColumns(j).name = "Product Number" Then
                targetCell.Value = variantName
            ElseIf tblFinalProductList.ListColumns(j).name = "Product Description" Then
                targetCell.Value = VariantDescription
            ElseIf tblFinalProductList.ListColumns(j).name = "Variant of" Then
                targetCell.Value = baseProduct
            ElseIf Not sourceCell.HasFormula Then
                targetCell.Value = sourceCell.Value
            End If
        Next j

        nextVariant = nextVariant + 1 ' Increment for next variant
    Next i
    ManualOverrides.SuppressChangeTracking = False

    ' ---------------------------
    ' Show Form to Select Base Routine
    ' ---------------------------
    Dim frm As frmSelectRoutineVariants
    Set frm = New frmSelectRoutineVariants
    
    ' Pass base Product and number of variants to the form
    frm.baseProduct = baseProduct
    frm.NumVariants = NumVariants
    frm.VariantNames = VariantNames

    
    ' Initialize and show the form
    frm.InitializeForm
    frm.Show
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


Function GetNextVariantNumber(baseProduct As String, tbl As ListObject) As Long
    Dim cell As Range, maxVar As Long, varNum As Long
    Dim variantPrefix As String, suffix As String

    variantPrefix = LCase(baseProduct) & "-v"

    maxVar = 0
    For Each cell In tbl.ListColumns("Product Number").DataBodyRange
        If LCase(Left(cell.Value, Len(variantPrefix))) = variantPrefix Then
            suffix = Mid(cell.Value, Len(variantPrefix) + 1)
            If IsNumeric(suffix) Then
                varNum = CLng(suffix)
                If varNum > maxVar Then maxVar = varNum
            End If
        End If
    Next cell

    GetNextVariantNumber = maxVar + 1
End Function

