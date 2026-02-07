VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVariantConfigurator 
   Caption         =   "Create Variant"
   ClientHeight    =   7140
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12465
   OleObjectBlob   =   "frmVariantConfigurator.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmVariantConfigurator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' --- Module-level cache to hold the base product's data for high-speed lookups ---
Private m_BaseProductData As Object
Public WasCancelled As Boolean
'====================================================================================================
'                                       FORM EVENTS & CONTROLS
'====================================================================================================

'This code is now robust and handles cases with one or many rows.
Private Sub UserForm_Initialize()
    ' --- DECLARATION BLOCK ---
    Dim ws As Worksheet, tbl As ListObject
    Dim productData As Variant, descriptionData As Variant
    Dim dict As Object
    Dim productList() As String
    Dim prodNum As String
    Dim key As Variant
    Dim i As Long

    ' --- INITIAL SETUP ---
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")
    productData = tbl.ListColumns("Product Number").DataBodyRange.Value2
    descriptionData = tbl.ListColumns("Product Description").DataBodyRange.Value2

    ' --- DATA PROCESSING ---
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    ' Handle both single-row (non-array) and multi-row (array) cases
    If IsArray(productData) Then
        ' Case 1: Multiple rows exist (2D array). Loop normally.
        For i = 1 To UBound(productData, 1)
            prodNum = CStr(productData(i, 1))
            If Len(prodNum) > 0 And Not dict.exists(prodNum) Then
                dict.Add prodNum, CStr(descriptionData(i, 1))
            End If
        Next i
    Else
        ' Case 2: Only one row exists (single value). Process it directly.
        prodNum = CStr(productData)
        If Len(prodNum) > 0 And Not dict.exists(prodNum) Then
            dict.Add prodNum, CStr(descriptionData)
        End If
    End If

    ' --- POPULATE COMBOBOX ---
    With cmbBaseProduct
        .Clear
        .ColumnCount = 2
        .ColumnWidths = "100;150"
    End With
    
    If dict.Count > 0 Then
        ReDim productList(0 To dict.Count - 1, 0 To 1)
        i = 0 ' Reset counter for the next loop
        For Each key In dict.Keys
            productList(i, 0) = key
            productList(i, 1) = dict(key)
            i = i + 1
        Next key
        cmbBaseProduct.List = productList
    End If
End Sub

Private Sub cmbBaseProduct_Change()
    If cmbBaseProduct.ListIndex = -1 Then Exit Sub

    Dim selectedProduct As String
    selectedProduct = cmbBaseProduct.column(0)
    txtBaseProductDesc.Text = cmbBaseProduct.column(1)

    LoadComponentList selectedProduct

    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("1. BOM Definition").ListObjects("BOMDefinition")
    
    txtVariantName.Text = NextFreeVariantName(selectedProduct, tbl)
    txtVariantDescription.Text = selectedProduct & " | Modified variant"
End Sub

Private Sub btnLoadComponents_Click()
    If cmbBaseProduct.ListIndex = -1 Then
        MsgBox "Please select a base product.", vbExclamation
        Exit Sub
    End If

    txtBaseProductDesc.Text = cmbBaseProduct.column(1)
    LoadComponentList cmbBaseProduct.column(0)
End Sub

Private Sub lvwComponents_DblClick()
    If lvwComponents.selectedItem Is Nothing Then Exit Sub
    Dim selectedItem As listItem: Set selectedItem = lvwComponents.selectedItem
    Dim newValue As String
    
    newValue = InputBox("Enter new quantity for " & selectedItem.Text, "Edit Quantity", selectedItem.SubItems(2))
    If newValue = "" Then Exit Sub
    
    ' Validate that the input is a valid, non-negative number, respecting regional settings
    Dim tempValue As String
    tempValue = Replace(newValue, ".", application.International(xlDecimalSeparator))
    tempValue = Replace(tempValue, ",", application.International(xlDecimalSeparator))
    
    If IsNumeric(tempValue) Then
        If CDbl(tempValue) >= 0 Then
            selectedItem.SubItems(2) = newValue ' Keep original user input format
        Else
            MsgBox "Quantity cannot be negative. Please enter a value greater than or equal to zero.", vbExclamation, "Invalid Quantity"
        End If
    Else
        MsgBox "Please enter a valid number for the quantity.", vbExclamation, "Invalid Input"
    End If
End Sub

Private Sub txtVariantName_Change()
    ' Live feedback if PN already exists in BOM
    If VariantExistsInBOM(txtVariantName.Text) Then
        txtVariantName.BackColor = RGB(255, 230, 230) ' light red
    Else
        txtVariantName.BackColor = vbWhite
    End If
End Sub

Private Sub btnCancel_Click()
    Me.WasCancelled = True
    Unload Me
End Sub

'''
' @Description Main routine to create the new variant.
' @Version 2.1 - Added comprehensive component data validation before execution.
'''
Private Sub btnCreateVariant_Click()
    ' --- DECLARATION BLOCK ---
    Dim wsBom As Worksheet, wsProducts As Worksheet
    Dim tblBOM As ListObject, tblProducts As ListObject
    Dim newRow As ListRow, baseRow As ListRow, prodRow As ListRow
    Dim baseProduct As String, variantName As String, variantDesc As String
    Dim resp As VbMsgBoxResult, altName As String
    Dim errorMsg As String
    Dim missingQty As Collection, negativeQty As Collection, invalidQty As Collection
    Dim compName As String, qtyString As String
    Dim materialKey As String
    Dim quantity As Double
    Dim sourceDataRow As Variant
    Dim i As Long, colIndex As Long

    ' --- SETUP AND VALIDATION ---
    On Error GoTo ErrorHandler

    If cmbBaseProduct.ListIndex = -1 Then
        MsgBox "Please select a base product first.", vbExclamation, "Input Required"
        Exit Sub
    End If
    variantName = Trim$(txtVariantName.Text)
    If variantName = "" Then
        MsgBox "Please enter a name for the new variant product.", vbExclamation, "Input Required"
        txtVariantName.SetFocus
        Exit Sub
    End If
    If lvwComponents.ListItems.Count = 0 Then
        MsgBox "There are no components to create a variant from. Please load them first.", vbExclamation, "No Components"
        Exit Sub
    End If
    baseProduct = cmbBaseProduct.column(0)
    variantDesc = txtVariantDescription.Text

    On Error Resume Next
    Set wsBom = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblBOM = wsBom.ListObjects("BOMDefinition")
    Set wsProducts = ThisWorkbook.Sheets("Final Products")
    Set tblProducts = wsProducts.ListObjects("FinalProductList")
    On Error GoTo ErrorHandler
    
    If tblBOM Is Nothing Or wsBom Is Nothing Or tblProducts Is Nothing Or wsProducts Is Nothing Then
        MsgBox "A required worksheet or table could not be found. Please check sheet names '1. BOM Definition' and 'Final Products'.", vbCritical, "Missing Object"
        Exit Sub
    End If

    ' --- COMPONENT DATA VALIDATION ---
    Set missingQty = New Collection
    Set negativeQty = New Collection
    Set invalidQty = New Collection
    For i = 1 To lvwComponents.ListItems.Count
        compName = lvwComponents.ListItems(i).Text
        qtyString = lvwComponents.ListItems(i).SubItems(2)
        
        If Trim(qtyString) = "" Then
            missingQty.Add compName
        ElseIf Not IsNumeric(Replace(Replace(qtyString, ",", "."), ".", application.International(xlDecimalSeparator))) Then
            invalidQty.Add compName
        ElseIf SafeCDbl(qtyString) < 0 Then
            negativeQty.Add compName
        End If
    Next i
    
    If missingQty.Count > 0 Then
        errorMsg = "The following components are missing a quantity:" & vbCrLf & "- " & JoinCollection(missingQty, vbCrLf & "- ")
    End If
    If invalidQty.Count > 0 Then
        If errorMsg <> "" Then errorMsg = errorMsg & vbCrLf & vbCrLf
        errorMsg = errorMsg & "The following have an invalid, non-numeric quantity:" & vbCrLf & "- " & JoinCollection(invalidQty, vbCrLf & "- ")
    End If
    If negativeQty.Count > 0 Then
        If errorMsg <> "" Then errorMsg = errorMsg & vbCrLf & vbCrLf
        errorMsg = errorMsg & "The following have a negative quantity (must be >= 0):" & vbCrLf & "- " & JoinCollection(negativeQty, vbCrLf & "- ")
    End If
    
    If errorMsg <> "" Then
        MsgBox "Please correct the following issues before creating the variant:" & vbCrLf & vbCrLf & errorMsg, vbExclamation, "Data Validation Failed"
        Exit Sub
    End If
    
    ' --- LOGIC & DATA HANDLING ---
    If VariantExistsInBOM(variantName) Then
        altName = NextFreeVariantName(baseProduct, tblBOM)
        resp = MsgBox("Product Number '" & variantName & "' already exists." & vbCrLf & vbCrLf & _
                      "Use the next available name '" & altName & "' instead?", _
                      vbExclamation + vbYesNoCancel, "Duplicate Product Number")
        
        If resp <> vbYes Then Exit Sub
        variantName = altName
        txtVariantName.Text = altName
    End If

    ' --- EXECUTION ---
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual

    ' --- 5.1. Create BOM Rows ---
    For i = 1 To lvwComponents.ListItems.Count
        quantity = SafeCDbl(lvwComponents.ListItems(i).SubItems(2))
        If quantity <> 0 Then
            Set newRow = tblBOM.ListRows.Add
            materialKey = lvwComponents.ListItems(i).Text
            If m_BaseProductData.exists(materialKey) Then
                sourceDataRow = m_BaseProductData(materialKey)
            Else
                ReDim sourceDataRow(1 To tblBOM.ListColumns.Count)
            End If
            For colIndex = 1 To tblBOM.ListColumns.Count
                With newRow.Range(colIndex)
                    If Not .HasFormula Then
                        Select Case tblBOM.ListColumns(colIndex).name
                            Case "Product Number": .Value = variantName
                            Case "Variant of": .Value = baseProduct
                            Case "Quantity": .Value = quantity
                            Case Else: .Value = sourceDataRow(colIndex)
                        End Select
                    End If
                End With
            Next colIndex
        End If
    Next i

    ' --- 5.2. Add to Final Products Table ---
    Set baseRow = Nothing ' Ensure baseRow is reset before the loop
    If Not tblProducts.DataBodyRange Is Nothing Then
        For Each baseRow In tblProducts.ListRows
            If CStr(baseRow.Range(tblProducts.ListColumns("Product Number").Index).Value) = baseProduct Then Exit For
        Next baseRow
    End If
    Set prodRow = tblProducts.ListRows.Add
    For colIndex = 1 To tblProducts.ListColumns.Count
        With prodRow.Range(colIndex)
            Select Case tblProducts.ListColumns(colIndex).name
                Case "Product Number": .Value = variantName
                Case "Product Description": .Value = variantDesc
                Case "Variant of": .Value = baseProduct
                Case Else
                    If Not .HasFormula And Not baseRow Is Nothing Then
                        .Value = baseRow.Range(colIndex).Value
                    End If
            End Select
        End With
    Next colIndex

    ' --- CLEANUP & NEXT STEPS ---
    Dim frmRoutine As New frmRoutineVariantEditor
    Unload Me
    Call Utils.RunProductBasedFormatting("1. BOM Definition", "BOMDefinition", "Helper Format BOMs")
    With frmRoutine
        .baseProduct = baseProduct
        .variantName = variantName
        .VariantDescription = variantDesc
        .InitializeForm
        .Show
    End With
    
    GoTo CleanExit

ErrorHandler:
    MsgBox "An unexpected error occurred." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.description, vbCritical, "Create Variant Error"
CleanExit:
    application.Calculation = xlCalculationAutomatic
    application.ScreenUpdating = True
End Sub

' --- NEW HELPER FUNCTION ---
' This small function is needed to format the error messages nicely.
' Please add this to your HELPER FUNCTIONS section.
Private Function JoinCollection(col As Collection, delimiter As String) As String
    Dim item As Variant
    Dim result As String
    If col.Count = 0 Then Exit Function
    
    For Each item In col
        result = result & CStr(item) & delimiter
    Next item
    ' Remove the trailing delimiter
    JoinCollection = Left(result, Len(result) - Len(delimiter))
End Function

'''
' @Description Populates the ListView with component data for the selected base product.
' @Version 4.0 - Refactored to perfectionist standards with a dedicated helper
'              function for array dimension checking, ensuring maximum clarity and robustness.
'''
Sub LoadComponentList(baseProduct As String)
    ' --- 1. DECLARATION BLOCK ---
    Dim ws As Worksheet, tbl As ListObject
    Dim componentsData As Variant
    Dim colIndices As Object
    Dim materialKey As String
    Dim i As Long
    Dim li As listItem ' ListViewItem object

    ' --- 2. INITIAL SETUP ---
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")
    componentsData = GetFilteredTableData(tbl, "Product Number", baseProduct)
    
    Set m_BaseProductData = CreateObject("Scripting.Dictionary")
    m_BaseProductData.CompareMode = vbTextCompare

    ' --- 3. CONFIGURE AND CLEAR LISTVIEW ---
    With Me.lvwComponents
        .ListItems.Clear
        ' Set up headers only if they don't already exist to prevent flickering.
        If .columnHeaders.Count = 0 Then
            .View = lvwReport
            .Gridlines = True
            .FullRowSelect = True
            .columnHeaders.Add , , "Material", 90
            .columnHeaders.Add , , "Material Description", 150
            .columnHeaders.Add , , "Quantity", 60
            .columnHeaders.Add , , "Base unit of component", 70
            .columnHeaders.Add , , "Vendor name", 100
            .columnHeaders.Add , , "Price per 1 unit", 90
        End If
    End With

    ' --- 4. DATA VALIDATION (Guard Clause) ---
    ' If the filter returned no data, there's nothing more to do.
    If IsEmpty(componentsData) Then Exit Sub

    Set colIndices = GetColumnIndices(tbl)

    ' --- 5. POPULATE LISTVIEW ---
    ' Use our robust helper function to determine how to process the data.
    If Is2DArray(componentsData) Then
        ' Case 1: MULTIPLE ROWS were found (2D Array).
        For i = 1 To UBound(componentsData, 1)
            materialKey = CStr(componentsData(i, colIndices("Material")))
            
            Set li = Me.lvwComponents.ListItems.Add(, , materialKey)
            li.SubItems(1) = CStr(componentsData(i, colIndices("Material description")))
            li.SubItems(2) = CStr(componentsData(i, colIndices("Quantity")))
            li.SubItems(3) = CStr(componentsData(i, colIndices("Base unit of component")))
            li.SubItems(4) = CStr(componentsData(i, colIndices("Vendor name")))
            li.SubItems(5) = CStr(componentsData(i, colIndices("Price per 1 unit")))

            If Not m_BaseProductData.exists(materialKey) Then
                m_BaseProductData.Add materialKey, GetRowAsArray(componentsData, i)
            End If
        Next i
    Else
        ' Case 2: A SINGLE ROW was found (1D Array).
        materialKey = CStr(componentsData(colIndices("Material")))

        Set li = Me.lvwComponents.ListItems.Add(, , materialKey)
        li.SubItems(1) = CStr(componentsData(colIndices("Material description")))
        li.SubItems(2) = CStr(componentsData(colIndices("Quantity")))
        li.SubItems(3) = CStr(componentsData(colIndices("Base unit of component")))
        li.SubItems(4) = CStr(componentsData(colIndices("Vendor name")))
        li.SubItems(5) = CStr(componentsData(colIndices("Price per 1 unit")))

        If Not m_BaseProductData.exists(materialKey) Then
            m_BaseProductData.Add materialKey, componentsData
        End If
    End If
End Sub

'====================================================================================================
'                                       HELPER FUNCTIONS
'====================================================================================================

Private Function VariantExistsInBOM(ByVal productNumber As String) As Boolean
    If Len(Trim$(productNumber)) = 0 Then Exit Function
    Dim tbl As ListObject: Set tbl = ThisWorkbook.Sheets("1. BOM Definition").ListObjects("BOMDefinition")
    If tbl.DataBodyRange Is Nothing Then Exit Function
    VariantExistsInBOM = (application.CountIf(tbl.ListColumns("Product Number").DataBodyRange, productNumber) > 0)
End Function

Private Function NextFreeVariantName(ByVal baseProduct As String, ByVal tbl As ListObject) As String
    Dim n As Long
    n = GetNextVariantNumber(baseProduct, tbl)
    Dim proposedName As String
    proposedName = baseProduct & "-V" & n
    ' Just in case of gaps, bump until free
    Do While VariantExistsInBOM(proposedName)
        n = n + 1
        proposedName = baseProduct & "-V" & n
    Loop
    NextFreeVariantName = proposedName
End Function

Function GetNextVariantNumber(baseProduct As String, tbl As ListObject) As Long
    ' --- DECLARATION BLOCK ---
    Dim bomData As Variant
    Dim prefix As String
    Dim cellValue As String, suffix As String
    Dim varNum As Long, maxNum As Long
    Dim i As Long

    ' --- INITIAL SETUP ---
    maxNum = 0
    If tbl.DataBodyRange Is Nothing Then
        GetNextVariantNumber = 1
        Exit Function
    End If
    
    bomData = tbl.ListColumns("Product Number").DataBodyRange.Value2
    prefix = LCase$(baseProduct) & "-v"
    
    ' --- DATA PROCESSING ---
    If IsArray(bomData) Then
        ' Multiple rows exist, loop through the array
        For i = 1 To UBound(bomData, 1)
            cellValue = CStr(bomData(i, 1))
            If LCase$(Left$(cellValue, Len(prefix))) = prefix Then
                suffix = Mid$(cellValue, Len(prefix) + 1)
                If IsNumeric(suffix) Then
                    varNum = CLng(suffix)
                    If varNum > maxNum Then maxNum = varNum
                End If
            End If
        Next i
    Else
        ' Only one row exists, process the single value
        cellValue = CStr(bomData)
        If LCase$(Left$(cellValue, Len(prefix))) = prefix Then
            suffix = Mid$(cellValue, Len(prefix) + 1)
            If IsNumeric(suffix) Then
                varNum = CLng(suffix)
                If varNum > maxNum Then maxNum = varNum
            End If
        End If
    End If
    
    GetNextVariantNumber = maxNum + 1
End Function


'''
' @Description Safely converts a string to a Double, handling different decimal separators.
' @Returns A Double. Returns 0 if the string is not a valid number.
'''
Private Function SafeCDbl(ByVal inputText As String) As Double
    On Error GoTo IsNotNumeric
    Dim cleanText As String
    Dim sysSep As String: sysSep = application.International(xlDecimalSeparator)
    Dim thousandSep As String: thousandSep = application.International(xlThousandsSeparator)
    
    ' Remove thousand separators first
    If Len(thousandSep) > 0 Then
        cleanText = Replace(inputText, thousandSep, "")
    Else
        cleanText = inputText
    End If
    
    ' Replace the non-system decimal separator with the system one
    If sysSep = "." Then
        cleanText = Replace(cleanText, ",", ".")
    Else ' sysSep is ","
        cleanText = Replace(cleanText, ".", ",")
    End If
    
    SafeCDbl = CDbl(cleanText)
    Exit Function
IsNotNumeric:
    SafeCDbl = 0
End Function

'''
' @Description Extracts a single row from a 2D array and returns it as a 1D array.
'''
Private Function GetRowAsArray(ByVal dataArray As Variant, ByVal rowIndex As Long) As Variant
    Dim result() As Variant
    ReDim result(1 To UBound(dataArray, 2))
    Dim i As Long
    For i = 1 To UBound(dataArray, 2)
        result(i) = dataArray(rowIndex, i)
    Next i
    GetRowAsArray = result
End Function

'''
' @Description Filters a ListObject and returns the data as a 2D array.
'''
Private Function GetFilteredTableData(ByVal sourceTable As ListObject, ByVal fieldName As String, ByVal criteria As String) As Variant
    Dim visibleRows As Range
    With sourceTable
        If .AutoFilter.FilterMode Then .AutoFilter.ShowAllData
        .Range.AutoFilter Field:=.ListColumns(fieldName).Index, Criteria1:=criteria
        On Error Resume Next
        Set visibleRows = .DataBodyRange.SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        .AutoFilter.ShowAllData
    End With
    If Not visibleRows Is Nothing Then
        GetFilteredTableData = visibleRows.Value2
    End If
End Function

'''
' @Description Gets all column indices from a table for fast lookups.
'              ROBUST VERSION: Uses a local variable for the dictionary to avoid conflicts.
'''
Private Function GetColumnIndices(ByVal tbl As ListObject) As Object
    ' Use a local dictionary object to build the results
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim col As ListColumn
    For Each col In tbl.ListColumns
        ' Check if the header value is valid before adding
        Dim headerValue As Variant
        headerValue = col.Range.Cells(1, 1).Value
        
        If Not IsError(headerValue) Then
            Dim colName As String
            colName = CStr(headerValue)
            
            If Len(colName) > 0 Then
                dict(colName) = col.Index
            End If
        End If
    Next col
    
    ' Return the completed dictionary
    Set GetColumnIndices = dict
End Function

'''
' @Description Checks if a given Variant contains a 2-dimensional array.
' @param arr The Variant to check.
' @return True if the Variant is a 2D array, otherwise False.
'''
Private Function Is2DArray(arr As Variant) As Boolean
    ' A robust check must first confirm the variant is an array at all.
    If Not IsArray(arr) Then
        Is2DArray = False
        Exit Function
    End If
    
    ' The standard VBA method to check for a second dimension is to
    ' attempt to access its boundary and see if an error is generated.
    On Error Resume Next
    Dim check As Long
    check = LBound(arr, 2)
    
    ' If Err.Number is 0, the operation succeeded, meaning a 2nd dimension exists.
    Is2DArray = (Err.Number = 0)
    
    ' Always reset the error handler to its default state.
    On Error GoTo 0
End Function
