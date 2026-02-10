Attribute VB_Name = "Search_copy_components"
Option Explicit

'================================================================================
' MODULE-LEVEL CONSTANTS
'================================================================================
' --- Define constants here to ensure consistency across all subs ---
Private Const BOM_SHEET_NAME As String = "1. BOM Definition"
Private Const BOM_TABLE_NAME As String = "BOMDefinition"
Private Const DATA_SHEET_NAME As String = "Purchasing Info Records"
Private Const DATA_TABLE_NAME As String = "LoadedData"

'================================================================================
' CORE REFACTORED SUBROUTINES
' These three subs form the new, modular core of your component handling logic.
'================================================================================

'''
' Inserts a new row into the BOMDefinition table and populates it with primary metadata.
' @param {String} partNumber The material number of the component.
' @param {String} plant The plant associated with the component.
' @param {Double} quantity The required quantity of the component.
' @param {String} alternateInfo Information about alternates.
' @param {String} [finalProductPn="Missing"] The final product number (Optional).
' @param {String} [finalProductCategory="Missing"] The final product category (Optional).
' @return {ListRow} The newly created ListRow object, or Nothing if an error occurred.
'''
Private Function InsertComponentRow(partNumber As String, ByVal plant As String, quantity As Double, alternateInfo As String, Optional finalProductPn As String = "Missing", Optional finalProductCategory As String = "Missing") As ListRow
    Dim wsBom As Worksheet
    Dim loBom As ListObject
    Dim targetRow As ListRow ' Changed from newRow to reflect it might be an existing row

    On Error GoTo ErrorHandler
    Set wsBom = ThisWorkbook.Sheets(BOM_SHEET_NAME)
    Set loBom = wsBom.ListObjects(BOM_TABLE_NAME)

    If loBom Is Nothing Then
        MsgBox "Fatal Error: The table '" & BOM_TABLE_NAME & "' could not be found.", vbCritical
        Exit Function
    End If

    ' Handle optional parameters
    If finalProductPn = "Missing" Then finalProductPn = wsBom.Range("F11").Value
    If finalProductCategory = "Missing" Then finalProductCategory = wsBom.Range("F13").Value

    '--- MODIFICATION START ---
    ' Check if the table has exactly one row and if that row's "Material" cell is empty.
    If loBom.ListRows.Count = 1 And IsEmpty(loBom.ListRows(1).Range(loBom.ListColumns("Material").Index).Value) Then
        ' If so, use the existing single row as the target.
        Set targetRow = loBom.ListRows(1)
    Else
        ' Otherwise, add a new row as normal.
        Set targetRow = loBom.ListRows.Add(AlwaysInsert:=True)
    End If
    '--- MODIFICATION END ---

    ' Populate the metadata in the target row
    With targetRow
        .Range(loBom.ListColumns("Material").Index).Value = partNumber
        .Range(loBom.ListColumns("Plant").Index).Value = plant
        .Range(loBom.ListColumns("Quantity").Index).Value = quantity
        .Range(loBom.ListColumns("Alternate").Index).Value = alternateInfo
        .Range(loBom.ListColumns("Product Number").Index).Value = finalProductPn
    End With

    ' Return the populated row
    Set InsertComponentRow = targetRow
    Exit Function

ErrorHandler:
    Set InsertComponentRow = Nothing
    MsgBox "An error occurred in InsertComponentRow: " & Err.description, vbCritical
End Function

'''
' Updates the details of a given BOM row by looking up data in the LoadedData table.
' @param {ListRow} targetRow The ListRow object in the BOMDefinition table to be updated.
' @param {Object} [overridesDict=Nothing] Optional Dictionary of manual overrides.
'        If provided, protected columns with recorded overrides are skipped.
' @return {String} A status string: "Updated", "NotFound", or "NoChange".
'''
Private Function UpdateComponentDetails(targetRow As ListRow, Optional overridesDict As Object = Nothing) As String
    Dim wsBom As Worksheet, wsData As Worksheet
    Dim loBom As ListObject, loData As ListObject
    Dim col As ListColumn, targetCell As Range
    Dim bomColIndex As Variant, sourceVal As Variant, destVal As Variant
    Dim partNumber As String, plant As String, searchKey As String
    Dim dataUpdated As Boolean, dataFound As Boolean
    Dim overrideKey As String

    dataUpdated = False
    dataFound = False

    On Error GoTo ErrorHandler

    ' --- 1. SETUP ---
    Set wsBom = ThisWorkbook.Sheets(BOM_SHEET_NAME)
    Set loBom = wsBom.ListObjects(BOM_TABLE_NAME)
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set loData = wsData.ListObjects(DATA_TABLE_NAME)

    If loBom Is Nothing Or loData Is Nothing Then
        MsgBox "A required table ('" & BOM_TABLE_NAME & "' or '" & DATA_TABLE_NAME & "') could not be found.", vbCritical
        UpdateComponentDetails = "NoChange" ' Prevent further errors
        Exit Function
    End If

    partNumber = targetRow.Range(loBom.ListColumns("Material").Index).Value
    plant = targetRow.Range(loBom.ListColumns("Plant").Index).Value
    searchKey = partNumber & " " & plant

    ' Read Product Number for override key (unique per BOM row)
    Dim productNum As String
    productNum = CStr(targetRow.Range(loBom.ListColumns("Product Number").Index).Value)

    ' --- 2. VERIFY a match exists before proceeding ---
    If application.WorksheetFunction.CountIf(loData.ListColumns("MatPlantID").DataBodyRange, searchKey) > 0 Then
        dataFound = True
    Else
        UpdateComponentDetails = "NotFound"
        Exit Function
    End If

    ' --- 3. AUTOMATED LOOKUP AND UPDATE ---
    application.ScreenUpdating = False

    For Each col In loData.ListColumns
        On Error Resume Next
        bomColIndex = application.Match(col.name, loBom.headerRowRange, 0)
        On Error GoTo ErrorHandler

        If Not IsError(bomColIndex) Then
            Select Case col.name
                ' Ignore key columns that are manually set or are identifiers
                Case "Material", "Plant", "Quantity", "Alternate", "Product Number", _
                     "MatPlantID", "SearchColumn", "MatSourceID", "LAPP Item"
                    ' Do nothing

                ' Skip "Price per 1 unit" - it is now a formula column (=[@Price]/[@[Price Unit]])
                Case "Price per 1 unit"
                    ' Do nothing - handled by Price mapping below

                Case Else
                    Set targetCell = targetRow.Range(bomColIndex)

                    ' Skip formula cells
                    If targetCell.HasFormula Then GoTo NextCol

                    ' Check if this column has a manual override that should be preserved
                    If Not overridesDict Is Nothing Then
                        If ManualOverrides.IsProtectedColumn(col.name) Then
                            overrideKey = partNumber & "|" & plant & "|" & productNum & "|" & col.name
                            If overridesDict.exists(overrideKey) Then
                                ' Override exists - skip this cell to preserve manual edit
                                GoTo NextCol
                            End If
                        End If
                    End If

                    ' Perform the lookup to get the source value
                    sourceVal = application.WorksheetFunction.XLookup(searchKey, _
                        loData.ListColumns("MatPlantID").DataBodyRange, _
                        loData.ListColumns(col.name).DataBodyRange, "")

                    ' Update the cell
                    dataUpdated = True
                    targetCell.Value = sourceVal
            End Select
        End If
NextCol:
    Next col

    ' --- 3b. PRICE MAPPING ---
    ' Source has "Price per 1 unit" -> map to BOM "Price" column, set "Price Unit" = 1
    Dim srcPriceVal As Variant
    On Error Resume Next
    srcPriceVal = application.WorksheetFunction.XLookup(searchKey, _
        loData.ListColumns("MatPlantID").DataBodyRange, _
        loData.ListColumns("Price per 1 unit").DataBodyRange, "")
    On Error GoTo ErrorHandler

    ' Write to "Price" column (if not overridden)
    Dim priceColIdx As Long, priceUnitColIdx As Long
    priceColIdx = 0
    priceUnitColIdx = 0
    On Error Resume Next
    priceColIdx = loBom.ListColumns("Price").Index
    priceUnitColIdx = loBom.ListColumns("Price Unit").Index
    On Error GoTo ErrorHandler

    If priceColIdx > 0 Then
        Set targetCell = targetRow.Range(priceColIdx)
        If Not targetCell.HasFormula Then
            Dim skipPrice As Boolean: skipPrice = False
            If Not overridesDict Is Nothing Then
                overrideKey = partNumber & "|" & plant & "|" & productNum & "|Price"
                If overridesDict.exists(overrideKey) Then skipPrice = True
            End If
            If Not skipPrice Then
                targetCell.Value = srcPriceVal
                dataUpdated = True
            End If
        End If
    End If

    ' Write "Price Unit" = 1 (if not overridden)
    If priceUnitColIdx > 0 Then
        Set targetCell = targetRow.Range(priceUnitColIdx)
        If Not targetCell.HasFormula Then
            Dim skipPriceUnit As Boolean: skipPriceUnit = False
            If Not overridesDict Is Nothing Then
                overrideKey = partNumber & "|" & plant & "|" & productNum & "|Price Unit"
                If overridesDict.exists(overrideKey) Then skipPriceUnit = True
            End If
            If Not skipPriceUnit Then
                targetCell.Value = 1
                dataUpdated = True
            End If
        End If
    End If

    ' --- 4. APPLY FORMULAS (Logic Update) ---
    On Error Resume Next
    Dim lappColIndex As Long
    lappColIndex = loBom.ListColumns("LAPP Item").Index
    On Error GoTo ErrorHandler

    If lappColIndex > 0 Then
        Dim targetRange As Range
        Set targetRange = targetRow.Range(lappColIndex)

        targetRange.NumberFormat = "General"

        Dim strFormula As String
        strFormula = "=IF(COUNTIF(LAPPCompanies[Firm], [@[Vendor name]]) > 0, ""Yes"", """")"

        If targetRange.Formula2 <> strFormula Then
            targetRange.Formula2 = strFormula
            dataUpdated = True
        End If
    End If

    Utils.ApplyRowFormatting targetRow
    ' --- 5. RETURN STATUS ---
    If dataUpdated Then
        UpdateComponentDetails = "Updated"
    Else
        UpdateComponentDetails = "NoChange"
    End If

    application.ScreenUpdating = True
    Exit Function

ErrorHandler:
    UpdateComponentDetails = "NoChange" ' Default return on error
    application.ScreenUpdating = True
    MsgBox "An error occurred in UpdateComponentDetails for material " & partNumber & ": " & Err.description, vbExclamation
End Function

'''
' A composite sub that first inserts a new component row, then updates its details.
' This replaces the old 'CopySelectedComponent' for adding new items.
' @param {String} partNumber The material number of the component.
' @param {String} plant The plant associated with the component.
' @param {Double} quantity The required quantity of the component.
' @param {String} alternateInfo Information about alternates.
'''
Public Sub AddFullComponent(partNumber As String, ByVal plant As String, quantity As Double, alternateInfo As String, Optional EndMaterialPn As String = "Missing")
    Dim newRow As ListRow

    ' Step 1: Insert the new row with basic metadata
    ManualOverrides.SuppressChangeTracking = True
    If EndMaterialPn = "Missing" Then
        Set newRow = InsertComponentRow(partNumber, plant, quantity, alternateInfo)
    Else
        Set newRow = InsertComponentRow(partNumber, plant, quantity, alternateInfo, EndMaterialPn)
    End If

    ' Step 2: If the row was created successfully, update it with details from the data source
    If Not newRow Is Nothing Then
        application.StatusBar = "Updating details for " & partNumber & "..."
        Call UpdateComponentDetails(newRow)
        application.StatusBar = False
    End If
    ManualOverrides.SuppressChangeTracking = False
End Sub


'================================================================================
' PRIMARY USER-FACING SUBROUTINES
' These are the main procedures that will be called by buttons or other processes.
'================================================================================

'''
' REPLACES AND REWRITES 'UpdateComponentIfChanged' and 'RefreshBomPrices'.
' Iterates through the BOMDefinition table, refreshes data from the source,
' and applies color-coding based on the update status.
'''
Public Sub RefreshBOMData()
    Dim wsBom As Worksheet
    Dim loBom As ListObject
    Dim currentRow As ListRow
    Dim material As String
    Dim updateStatus As String
    Dim updatedCount As Long, notFoundCount As Long
    Dim materialCell As Range
    Dim dict As Object

    On Error GoTo ErrorHandler

    ' Load Database if necessary
    FilterComponents.LoadDatabase True

    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual

    Set wsBom = ThisWorkbook.Sheets(BOM_SHEET_NAME)
    Set loBom = wsBom.ListObjects(BOM_TABLE_NAME)

    If loBom Is Nothing Then
        MsgBox "The table '" & BOM_TABLE_NAME & "' could not be found. Aborting.", vbCritical
        GoTo CleanExit
    End If

    ' Clear colors only from the Material column
    loBom.ListColumns("Material").DataBodyRange.Interior.ColorIndex = xlNone

    updatedCount = 0
    notFoundCount = 0

    ' Load manual overrides dictionary for protected column checking
    Set dict = ManualOverrides.LoadOverridesDict()

    ' Suppress change tracking for all programmatic writes
    ManualOverrides.SuppressChangeTracking = True

    ' Loop through each row in the BOMDefinition table
    For Each currentRow In loBom.ListRows
        Set materialCell = currentRow.Range.Cells(1, loBom.ListColumns("Material").Index)
        material = CStr(materialCell.Value)

        ' Skip "NEW" components
        If Not UCase(material) Like "NEW*" And material <> "" Then
            application.StatusBar = "Checking: " & material

            ' Call the update function with overrides dict
            updateStatus = UpdateComponentDetails(currentRow, dict)
            If Len(Trim$(updateStatus)) = 0 Then updateStatus = "NoChange"

            ' Apply color-coding to only the Material cell
            Select Case StatusSeverity(updateStatus)
                Case 1 ' Updated
                    materialCell.Interior.Color = RGB(255, 255, 153)
                    updatedCount = updatedCount + 1
                Case 2 ' Not Found
                    materialCell.Interior.Color = RGB(255, 102, 102)
                    notFoundCount = notFoundCount + 1
                Case Else
                    ' Leave it blank (no color)
            End Select
        End If
    Next currentRow

    MsgBox "BOM data refresh complete." & vbCrLf & vbCrLf & _
           "Updated Rows: " & updatedCount & vbCrLf & _
           "Rows Not Found: " & notFoundCount, vbInformation, "Refresh Complete"
    Call Utils.RunProductBasedFormatting(BOM_SHEET_NAME, BOM_TABLE_NAME, "Helper Format BOMs")
CleanExit:
    ManualOverrides.SuppressChangeTracking = False
    application.StatusBar = False
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "A critical error occurred in RefreshBOMData: " & Err.description, vbCritical
    Resume CleanExit
End Sub

Private Function StatusSeverity(ByVal s As String) As Long
    Dim u As String
    u = UCase$(Trim$(CStr(s)))
    u = Replace(u, vbCr, " ")
    u = Replace(u, vbLf, " ")

    If u Like "*NOT*FOUND*" Or u Like "*NO*MATCH*" Or u Like "*MISSING*" Then
        StatusSeverity = 2 ' Red
    ElseIf u Like "*UPDAT*" Or u Like "*CHANG*" Or u Like "*REFRESH*" Then
        StatusSeverity = 1 ' Yellow
    Else
        StatusSeverity = 0 ' No color
    End If
End Function




Public Sub ProcessMassUploadData()
    Dim wsGlobal As Worksheet, wsData As Worksheet, wsMassUpload As Worksheet
    Dim loMassUpload As ListObject, loLoadedData As ListObject
    Dim rComponentCell As Range
    Dim sPlantGlobal As String, sSourceGlobal As String, sCurrentComponent As String, sFinalProduct As String
    Dim vLookupResult As Variant
    Dim dQuantity As Double, lCurrentRow As Long
    Dim netPriceUnit As Variant
    Dim notFoundList As String

    On Error GoTo ErrorHandler
    
    ' Load database if needed
    FilterComponents.LoadDatabase True
    
    application.ScreenUpdating = False

    ' --- Initialization ---
    Set wsGlobal = ThisWorkbook.Sheets("Global Variables")
    Set wsData = ThisWorkbook.Sheets(DATA_SHEET_NAME)
    Set wsMassUpload = ThisWorkbook.Sheets("MassUpload")
    Set loMassUpload = wsMassUpload.ListObjects("MassUploadTable")
    Set loLoadedData = wsData.ListObjects(DATA_TABLE_NAME)

    If loMassUpload Is Nothing Or loLoadedData Is Nothing Then
        MsgBox "A required table could not be found. Aborting.", vbCritical
        GoTo CleanExit
    End If

    sPlantGlobal = wsGlobal.Range("B3").Value
    sSourceGlobal = wsGlobal.Range("B2").Value
    notFoundList = ""

    ' Suppress change tracking for all programmatic writes
    ManualOverrides.SuppressChangeTracking = True

    ' --- Loop through each row of the MassUploadTable ---
    For Each rComponentCell In loMassUpload.ListColumns("Component").DataBodyRange.Cells
        sCurrentComponent = CStr(rComponentCell.Value)
        If sCurrentComponent = "" Then GoTo NextRow
        
        lCurrentRow = rComponentCell.row - loMassUpload.headerRowRange.row
        dQuantity = loMassUpload.ListColumns("Quantity").DataBodyRange.Cells(lCurrentRow).Value
        sFinalProduct = CStr(loMassUpload.ListColumns("Product Number").DataBodyRange.Cells(lCurrentRow).Value)
        netPriceUnit = Empty
        On Error Resume Next
        netPriceUnit = loMassUpload.ListColumns("Price per 1 unit").DataBodyRange.Cells(lCurrentRow).Value
        On Error GoTo 0

        Dim sAlternate As String: sAlternate = "No" ' keep as in your original

        ' --- Prioritized Search Logic ---
        ' 1) Component & Global Plant
        If application.WorksheetFunction.CountIf(loLoadedData.ListColumns("MatPlantID").DataBodyRange, _
                                                 sCurrentComponent & " " & sPlantGlobal) > 0 Then
            Call AddFullComponent(sCurrentComponent, sPlantGlobal, dQuantity, sAlternate, sFinalProduct)
            GoTo NextRow
        End If

        ' 2) Component & Global Source  -> returns Plant
        vLookupResult = application.WorksheetFunction.XLookup( _
                             sCurrentComponent & " " & sSourceGlobal, _
                             loLoadedData.ListColumns("MatSourceID").DataBodyRange, _
                             loLoadedData.ListColumns("Plant").DataBodyRange, "")
        If vLookupResult <> "" Then
            Call AddFullComponent(sCurrentComponent, CStr(vLookupResult), dQuantity, sAlternate, sFinalProduct)
            GoTo NextRow
        End If
        
        ' 3) Component & "TP List"
        If application.WorksheetFunction.CountIf(loLoadedData.ListColumns("MatPlantID").DataBodyRange, _
                                                 sCurrentComponent & " TP List") > 0 Then
            Call AddFullComponent(sCurrentComponent, "TP List", dQuantity, sAlternate, sFinalProduct)
            GoTo NextRow
        End If
        
        ' 4) Material only (first plant found)
        vLookupResult = application.WorksheetFunction.XLookup( _
                             sCurrentComponent, _
                             loLoadedData.ListColumns("Material").DataBodyRange, _
                             loLoadedData.ListColumns("Plant").DataBodyRange, "")
        If vLookupResult <> "" Then
            Call AddFullComponent(sCurrentComponent, CStr(vLookupResult), dQuantity, sAlternate, sFinalProduct)
        Else
            ' --- NEW: if still not found, add placeholder into BOMDefinition and highlight Material ---
            notFoundList = notFoundList & sCurrentComponent & vbCrLf
            On Error Resume Next
            AddPlaceholderComponentToBOM sCurrentComponent, dQuantity, sFinalProduct, netPriceUnit
            If Err.Number <> 0 Then
                ' Fallback (only if helper is missing)
                Err.Clear
                Dim wsDest As Worksheet, tblDest As ListObject, newRow As ListRow
                Dim cMat As Long, cQty As Long, cProd As Long, cPrice As Long
                Set wsDest = ThisWorkbook.Sheets("1. BOM Definition")
                Set tblDest = wsDest.ListObjects("BOMDefinition")
                Set newRow = tblDest.ListRows.Add
                On Error Resume Next
                cMat = tblDest.ListColumns("Material").Index
                cQty = tblDest.ListColumns("Quantity").Index
                cProd = tblDest.ListColumns("Product Number").Index
                cPrice = tblDest.ListColumns("Price").Index
                Dim cPU As Long
                cPU = tblDest.ListColumns("Price Unit").Index
                On Error GoTo 0
                If cProd > 0 Then If Not newRow.Range.Cells(1, cProd).HasFormula Then newRow.Range.Cells(1, cProd).Value = sFinalProduct
                If cMat > 0 Then If Not newRow.Range.Cells(1, cMat).HasFormula Then newRow.Range.Cells(1, cMat).Value = sCurrentComponent
                If cQty > 0 Then If Not newRow.Range.Cells(1, cQty).HasFormula Then newRow.Range.Cells(1, cQty).Value = dQuantity
                If cPrice > 0 Then
                    If IsNumeric(netPriceUnit) And CDbl(netPriceUnit) <> 0 Then
                        If Not newRow.Range.Cells(1, cPrice).HasFormula Then newRow.Range.Cells(1, cPrice).Value = CDbl(netPriceUnit)
                    End If
                End If
                If cPU > 0 Then If Not newRow.Range.Cells(1, cPU).HasFormula Then newRow.Range.Cells(1, cPU).Value = 1
                If cMat > 0 Then newRow.Range.Cells(1, cMat).Interior.Color = RGB(255, 255, 0)
            End If
            On Error GoTo ErrorHandler
        End If

NextRow:
    Next rComponentCell

    If notFoundList <> "" Then
        MsgBox "Processed with placeholders for components not found in the database. Added to BOMDefinition and highlighted:" & _
               vbCrLf & notFoundList, vbExclamation
    Else
        MsgBox "Processing of Mass Upload is complete.", vbInformation
    End If

    ThisWorkbook.Sheets(BOM_SHEET_NAME).Activate
    SortSelectedComponentsByProduct
CleanExit:
    ManualOverrides.SuppressChangeTracking = False
    application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "An unexpected error occurred: " & Err.description, vbCritical, "Error in ProcessMassUploadData"
    Resume CleanExit
End Sub

Public Sub AddPlaceholderComponentToBOM( _
    ByVal materialPN As String, _
    ByVal qty As Double, _
    ByVal productNumber As String, _
    Optional ByVal netPriceUnit As Variant)

    Dim wsDest As Worksheet
    Dim tblDest As ListObject
    Dim newRow As ListRow
    Dim cMat As Long, cQty As Long, cProd As Long, cPrice As Long, cPriceUnit As Long

    Set wsDest = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblDest = wsDest.ListObjects("BOMDefinition")

    ' Add row
    Set newRow = tblDest.ListRows.Add

    ' Resolve destination column indices safely
    On Error Resume Next
    cMat = tblDest.ListColumns("Material").Index
    cQty = tblDest.ListColumns("Quantity").Index
    cProd = tblDest.ListColumns("Product Number").Index
    cPrice = tblDest.ListColumns("Price").Index
    cPriceUnit = tblDest.ListColumns("Price Unit").Index
    On Error GoTo 0

    ' Write values only if target cells are not formula cells
    If cProd > 0 Then
        If Not newRow.Range.Cells(1, cProd).HasFormula Then newRow.Range.Cells(1, cProd).Value = productNumber
    End If

    If cMat > 0 Then
        If Not newRow.Range.Cells(1, cMat).HasFormula Then newRow.Range.Cells(1, cMat).Value = materialPN
    End If

    If cQty > 0 Then
        If Not newRow.Range.Cells(1, cQty).HasFormula Then newRow.Range.Cells(1, cQty).Value = qty
    End If

    If cPrice > 0 Then
        If Not IsMissing(netPriceUnit) Then
            If IsNumeric(netPriceUnit) And CDbl(netPriceUnit) <> 0 Then
                If Not newRow.Range.Cells(1, cPrice).HasFormula Then
                    newRow.Range.Cells(1, cPrice).Value = CDbl(netPriceUnit)
                End If
            End If
        End If
    End If

    ' Set Price Unit to 1
    If cPriceUnit > 0 Then
        If Not newRow.Range.Cells(1, cPriceUnit).HasFormula Then newRow.Range.Cells(1, cPriceUnit).Value = 1
    End If

    ' Highlight the Material cell so it's easy to spot
    If cMat > 0 Then newRow.Range.Cells(1, cMat).Interior.Color = RGB(255, 255, 0)
End Sub

'================================================================================
' UNCHANGED HELPER SUBROUTINES
' These subs were not part of the refactoring request and remain as they were.
'================================================================================

Public Sub RefreshSearch()
    Dim selectedProduct As String
    selectedProduct = ThisWorkbook.Sheets("1. BOM definition").Range("F11").Value
 
    If selectedProduct = "" Then
        MsgBox "No product is selected. Please select a product or add a new one.", vbExclamation
        AddProductForm.Show ' Assuming this form exists
        Exit Sub
    End If
    
    ResultsForm.Show ' Assuming this form exists
End Sub




Public Sub GoToMassUpload()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("MassUpload")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet 'MassUpload' not found!", vbExclamation
    Else
        ws.Activate
    End If
End Sub








