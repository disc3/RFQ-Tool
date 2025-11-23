Attribute VB_Name = "ProductsButton"
Sub ShowAddProductForm()
    AddProductForm.Show
End Sub
Public Sub UpdateRoutineDropdown()
    Dim ws As Worksheet
    Dim wsRoutine As Worksheet
    Dim tbl As ListObject
    Dim FinalProductList As Range
    Dim namedRange As Range

    ' Reference the "Final Products" sheet and FinalProductList table
    Set ws = ThisWorkbook.Sheets("Final Products")
    Set tbl = ws.ListObjects("FinalProductList")

    ' Reference the "2. Routines" sheet
    Set wsRoutine = ThisWorkbook.Sheets("2. Routines")

    ' Define the range for the dropdown in "FinalProductList"
    On Error Resume Next
    Set FinalProductList = tbl.ListColumns("Product Number").DataBodyRange
    On Error GoTo 0

    ' Check if FinalProductList is valid and contains data
    If Not FinalProductList Is Nothing And application.WorksheetFunction.CountA(FinalProductList) > 0 Then
        ' Define a named range to avoid issues with commas
        Set namedRange = FinalProductList

        ' Create a named range in the workbook (overwrite if it exists)
        On Error Resume Next
        ThisWorkbook.names.Add name:="RoutineDropdown", RefersTo:=namedRange
        On Error GoTo 0

        ' Clear existing validation if any
        wsRoutine.Range("D6").Validation.Delete

        ' Apply the validation using the named range
        With wsRoutine.Range("D6").Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=RoutineDropdown"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Else
        ' Clear validation if no items are available to populate
        wsRoutine.Range("D6").Validation.Delete
        MsgBox "No products found to populate the dropdown in '2. Routines'.", vbInformation
    End If
End Sub

Public Sub UpdateProductDropdown()
    Dim ws As Worksheet
    Dim wsProductSpec As Worksheet
    Dim tbl As ListObject
    Dim FinalProductList As Range
    Dim namedRange As Range

    ' Reference the Products sheet and FinalProductList table
    Set ws = ThisWorkbook.Sheets("Final Products")
    Set tbl = ws.ListObjects("FinalProductList")

    ' Reference the Product Specification sheet
    Set wsProductSpec = ThisWorkbook.Sheets("1. BOM definition")

    ' Define the range for the dropdown in Product Specification
    On Error Resume Next
    Set FinalProductList = tbl.ListColumns("Product Number").DataBodyRange
    On Error GoTo 0

    ' Check if FinalProductList is valid and contains data
    If Not FinalProductList Is Nothing And application.WorksheetFunction.CountA(FinalProductList) > 0 Then
        ' Define a named range to avoid issues with commas
        Set namedRange = FinalProductList

        ' Create a named range in the workbook (overwrite if it exists)
        On Error Resume Next
        ThisWorkbook.names.Add name:="ProductDropdown", RefersTo:=namedRange
        On Error GoTo 0

        ' Clear existing validation if any
        wsProductSpec.Range("F11").Validation.Delete

        ' Apply the validation using the named range
        With wsProductSpec.Range("F11").Validation
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="=ProductDropdown"
            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True
        End With
    Else
        ' Clear validation if no items are available to populate
        wsProductSpec.Range("F11").Validation.Delete
        MsgBox "No products found to populate the dropdown.", vbInformation
    End If
End Sub
Sub DeleteAllProducts()
    Dim wsProducts As Worksheet
    Dim wsSelectedRoutines As Worksheet
    Dim wsSelectedComponents As Worksheet
    Dim wsPlantVariables As Worksheet
    Dim tblProducts As ListObject
    Dim tblPlantFormats As ListObject
    Dim confirmDelete As VbMsgBoxResult
    Dim plantRow As ListRow
    Dim outputSheetName As String
    Dim wsOutput As Worksheet
    'Dim wsValidation As Worksheet

    ' Set the worksheets and tables
    Set wsProducts = ThisWorkbook.Sheets("Final Products")
    Set wsSelectedComponents = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsSelectedRoutines = ThisWorkbook.Sheets("2. Routines")
    'Set wsValidation = ThisWorkbook.Sheets("3. Clarification Validation")

    Set wsPlantVariables = ThisWorkbook.Sheets("Plant Variables")
    Set tblProducts = wsProducts.ListObjects("FinalProductList")
    Set wsOutput = ThisWorkbook.Sheets("4. Sales Calculation (Internal)") ' Set reference to the Output sheet

    ' Set the PlantExportFormats table
    On Error Resume Next
    Set tblPlantFormats = wsPlantVariables.ListObjects("PlantExportFormats")
    On Error GoTo 0

    ' Check if PlantExportFormats table exists
    If tblPlantFormats Is Nothing Then
        MsgBox "Table 'PlantExportFormats' not found in 'Plant Variables' sheet.", vbExclamation
        Exit Sub
    End If

    ' Prompt user for confirmation before deleting all products and routines
    confirmDelete = MsgBox("Are you sure you want to delete all products, selected routines, selected components, and generated sheets?", vbYesNo + vbQuestion, "Confirm Delete")

    ' If user confirms, clear the tables
    If confirmDelete = vbYes Then
        ' Delete all rows in the FinalProductList table except the first row, then clear the first row
        If tblProducts.ListRows.Count > 0 Then
            ' If there's more than one row, delete all rows except the first
            If tblProducts.ListRows.Count > 1 Then
                tblProducts.DataBodyRange.offset(1, 0).Resize(tblProducts.ListRows.Count - 1).Rows.Delete
            End If
            ' Clear contents of the first row to leave an empty row
            tblProducts.ListRows(1).Range.ClearContents
        End If

        ' Call the function to clear selected components and routines
        ClearSelectedComponentsTable
        ClearSelectedRoutinesTable
        ClearProjectDataColumns
        ClearMassUploadTable

        ' Delete generated sheets based on "Output Routing" and "Output BOM" columns in "PlantExportFormats"
        On Error Resume Next ' Suppress errors if sheets do not exist
        For Each plantRow In tblPlantFormats.ListRows
            ' Delete the "Output Routing" sheet if it exists
            outputSheetName = plantRow.Range(tblPlantFormats.ListColumns("Output Routing").Index).Value
            If outputSheetName <> "" Then
                application.DisplayAlerts = False ' Suppress prompt to confirm sheet deletion
                ThisWorkbook.Sheets(outputSheetName).Delete
                application.DisplayAlerts = True
            End If
            
            ' Delete the "Output BOM" sheet if it exists
            outputSheetName = plantRow.Range(tblPlantFormats.ListColumns("Output BOM").Index).Value
            If outputSheetName <> "" Then
                application.DisplayAlerts = False ' Suppress prompt to confirm sheet deletion
                ThisWorkbook.Sheets(outputSheetName).Delete
                application.DisplayAlerts = True
            End If
        Next plantRow
        On Error GoTo 0 ' Reset error handling

        ' Clear specific cells in the "Output" sheet
        On Error Resume Next
        wsOutput.Range("A1").ClearContents
        
        On Error GoTo 0

        ' Show completion message
        MsgBox "All products, selected routines, selected components, and generated sheets have been deleted.", vbInformation

        ' Clear the dropdown selection in Product Specification
        ThisWorkbook.Sheets("1. BOM definition").Range("F11").ClearContents
        ThisWorkbook.Sheets("2. Routines").Range("D6").ClearContents
        
        ' Clear the statuses for customer and purchasing clarification
        ThisWorkbook.Sheets("3. Clarification Validation").Range("E6:G23").ClearContents
        ThisWorkbook.Sheets("3. Clarification Validation").Range("O14:O24").ClearContents
        ThisWorkbook.Sheets("3. Clarification Validation").Range("O14:O24").Interior.ColorIndex = xlNone ' Reset the interior color to transparent
        
        ' Clear validation after deleting all products
        
        ' Update cell J7 on the Clarification/Validation sheet
        With ThisWorkbook.Sheets("3. Clarification Validation").Range("J7")
            .Value = "All Products cleared. Please add new products and validate the RFQ"
            .Interior.Color = RGB(255, 255, 0) ' Yellow color
        End With
        
        ' Clear RFQ sent info
        ThisWorkbook.Sheets("4. Sales Calculation (Internal)").Range("N1").ClearContents
        
        ' Call the UpdateProductDropdown to refresh the dropdown in Product Specification
        UpdateProductDropdown
        
        ' Clear the BOM Exporter sheet
        ThisWorkbook.Sheets("Template_BOM_Connect").Range("A3:X999").ClearContents
        
         ' Clear the Routing Exporter sheet
        ThisWorkbook.Sheets("Template_Routing_Connect").Range("A4:X999").ClearContents
        
        HideChainSheets
    Else
        UpdateProductDropdown
        MsgBox "No items were deleted.", vbInformation
    End If
End Sub

Sub ClearOutputRange()
    Dim wsOutput As Worksheet

    ' Set the Output worksheet reference
    On Error Resume Next
    Set wsOutput = ThisWorkbook.Sheets("Output")
    On Error GoTo 0

    ' Ensure the sheet exists
    If wsOutput Is Nothing Then
        MsgBox "The Output sheet does not exist.", vbExclamation
        Exit Sub
    End If

    ' Unprotect the sheet if it is protected
    On Error Resume Next
    On Error GoTo 0

    ' Clear the contents of the specified range
    On Error Resume Next
    wsOutput.Range("E10:F99").Value = "" ' Explicitly clear the range
    On Error GoTo 0

    ' Reprotect the sheet if necessary
    On Error Resume Next
    On Error GoTo 0

    'MsgBox "The contents of E10:F99 have been cleared.", vbInformation
End Sub


Sub HideChainSheets()
    Dim sheetName As Variant
    Dim chainSheets As Variant
    chainSheets = Array("Page 1 Chain RFQ Form", "Page 2 Chain RFQ Form", "Page 3 Chain RFQ Form", "Example Template Chain Layout", "Example Connection Plan")

    For Each sheetName In chainSheets
        On Error Resume Next
        ThisWorkbook.Sheets(sheetName).Visible = xlSheetHidden
        On Error GoTo 0
    Next sheetName

    ' Hide the button as well
    On Error Resume Next
    ThisWorkbook.Sheets("1. BOM Definition").Shapes("btnOpenChainForm").Visible = False
    On Error GoTo 0
End Sub


