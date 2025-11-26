Attribute VB_Name = "AddSpecialRoutineSteps"
Option Explicit

' ==========================================================================
' PUBLIC SUBROUTINE: AddMaterialPreparingRoutineIfNeeded (Single Mode)
' ==========================================================================
' PURPOSE: Adds a routine for a SINGLE product.
'          Ensures formulas in undefined columns are preserved/copied down.
' ==========================================================================
Public Sub AddMaterialPreparingRoutineIfNeeded(ByVal productNumber As String)
    On Error GoTo ErrorHandler
    
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual

    ' 1. Get Template Data
    Dim templateData As Variant
    templateData = GetMaterialPreparingTemplateData()
    If IsEmpty(templateData) Then GoTo CleanExit

    ' 2. Define Destination
    Dim wsSelected As Worksheet: Set wsSelected = ThisWorkbook.Sheets("2. Routines")
    Dim tblSelected As ListObject: Set tblSelected = wsSelected.ListObjects("SelectedRoutines")
    
    ' 3. Determine Target Row
    Dim destRow As ListRow
    ' Logic: If table has 1 row and it's empty, use it. Otherwise, add new.
    If tblSelected.ListRows.Count = 1 And _
       IsEmpty(tblSelected.DataBodyRange(1, tblSelected.ListColumns("Product Number").Index).Value) Then
        Set destRow = tblSelected.ListRows(1)
    Else
        Set destRow = tblSelected.ListRows.Add(AlwaysInsert:=True)
    End If
    
    ' 4. CRITICAL: Fill Formulas Down BEFORE writing data
    ' (ListRows.Add usually does this, but this guarantees it for the placeholder scenario too)
    If tblSelected.ListRows.Count > 1 Then
        destRow.Range.FillDown
    End If
    
    ' 5. Write Hard Values to specific columns only
    With destRow
        .Range(1, tblSelected.ListColumns("Plant").Index).Value = templateData(8)
        .Range(1, tblSelected.ListColumns("Product Number").Index).Value = productNumber
        .Range(1, tblSelected.ListColumns("Macrophase").Index).Value = "Stock"
        .Range(1, tblSelected.ListColumns("Microphase").Index).Value = "Material preparing"
        
        .Range(1, tblSelected.ListColumns("Material").Index).Value = templateData(0)
        .Range(1, tblSelected.ListColumns("Machine").Index).Value = templateData(1)
        .Range(1, tblSelected.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index).Value = templateData(2)
        .Range(1, tblSelected.ListColumns("Wire/component dimensions  (mm)").Index).Value = templateData(3)
        .Range(1, tblSelected.ListColumns("Work Center Code").Index).Value = templateData(4)
        .Range(1, tblSelected.ListColumns("tr").Index).Value = templateData(5)
        .Range(1, tblSelected.ListColumns("te").Index).Value = templateData(6)
        .Range(1, tblSelected.ListColumns("Number of Operations").Index).Value = 1
        .Range(1, tblSelected.ListColumns("Number of Setups").Index).Value = 1
        .Range(1, tblSelected.ListColumns("Sort Order").Index).Value = templateData(7)
    End With

CleanExit:
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error in Single Add: " & Err.description, vbCritical
    Resume CleanExit
End Sub

' ==========================================================================
' PUBLIC SUBROUTINE: AddMaterialPreparingRoutines_Bulk (Mass Mode)
' ==========================================================================
' PURPOSE: 1. Calculates data in memory.
'          2. Resizes table.
'          3. Performs FILLDOWN to carry over formulas in non-touched columns.
'          4. Overwrites specific columns with new hard data.
' ==========================================================================
Public Sub AddMaterialPreparingRoutines_Bulk()
    On Error GoTo ErrorHandler
    
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual

    ' --- 1. Get Template Data ---
    Dim templateData As Variant
    templateData = GetMaterialPreparingTemplateData()
    If IsEmpty(templateData) Then GoTo CleanExit

    ' --- 2. Load Product List ---
    Dim wsProducts As Worksheet: Set wsProducts = ThisWorkbook.Sheets("Final Products")
    Dim tblProducts As ListObject: Set tblProducts = wsProducts.ListObjects("FinalProductList")
    
    If tblProducts.DataBodyRange Is Nothing Then GoTo CleanExit
    
    Dim arrProdData As Variant
    arrProdData = tblProducts.DataBodyRange.Value
    
    Dim idxProdNum As Long: idxProdNum = tblProducts.ListColumns("Product Number").Index
    Dim idxHelper As Long: idxHelper = tblProducts.ListColumns("Helper NeedsMaterialPreparingRoutine").Index

    ' --- 3. Identify Items to Add ---
    Dim countToAdd As Long: countToAdd = 0
    Dim i As Long
    
    For i = 1 To UBound(arrProdData, 1)
        If arrProdData(i, idxHelper) = True Then countToAdd = countToAdd + 1
    Next i
    
    If countToAdd = 0 Then GoTo CleanExit
    
    ' --- 4. Prepare Destination Table ---
    Dim wsSelected As Worksheet: Set wsSelected = ThisWorkbook.Sheets("2. Routines")
    Dim tblSelected As ListObject: Set tblSelected = wsSelected.ListObjects("SelectedRoutines")
    
    ' Map the columns we will WRITE to.
    ' Any column NOT in this list will be handled by FillDown (Formulas).
    Dim colNames As Variant
    colNames = Array("Plant", "Product Number", "Macrophase", "Microphase", _
                     "Material", "Machine", "Wire/cable dimension diameter/section  (mm/mm2)", _
                     "Wire/component dimensions  (mm)", "Work Center Code", _
                     "tr", "te", "Number of Operations", "Number of Setups", "Sort Order")
    
    ' Prepare a 2D array to hold ONLY the data we are writing
    ' Dimensions: (1 to countToAdd, 1 to Number of Columns in our list above)
    Dim arrData() As Variant
    ReDim arrData(1 To countToAdd, 1 To UBound(colNames) + 1)
    
    ' --- 5. Fill Data Array in Memory ---
    Dim currRow As Long: currRow = 1
    
    For i = 1 To UBound(arrProdData, 1)
        If arrProdData(i, idxHelper) = True Then
            ' Fill the array row based on the order in colNames
            ' 0: Plant
            arrData(currRow, 1) = templateData(8)
            ' 1: Product Number
            arrData(currRow, 2) = arrProdData(i, idxProdNum)
            ' 2: Macrophase
            arrData(currRow, 3) = "Stock"
            ' 3: Microphase
            arrData(currRow, 4) = "Material preparing"
            ' 4: Material
            arrData(currRow, 5) = templateData(0)
            ' 5: Machine
            arrData(currRow, 6) = templateData(1)
            ' 6: Wire 1
            arrData(currRow, 7) = templateData(2)
            ' 7: Wire 2
            arrData(currRow, 8) = templateData(3)
            ' 8: Work Center
            arrData(currRow, 9) = templateData(4)
            ' 9: tr
            arrData(currRow, 10) = templateData(5)
            ' 10: te
            arrData(currRow, 11) = templateData(6)
            ' 11: Ops
            arrData(currRow, 12) = 1
            ' 12: Setups
            arrData(currRow, 13) = 1
            ' 13: Sort Order
            arrData(currRow, 14) = templateData(7)
            
            currRow = currRow + 1
        End If
    Next i
    
    ' --- 6. Resize Table & FillDown Formulas ---
    Dim initialRows As Long: initialRows = tblSelected.ListRows.Count
    Dim startWriteRow As Long
    Dim cProdIdx As Long: cProdIdx = tblSelected.ListColumns("Product Number").Index

    ' Determine where to start writing
    If initialRows = 1 And IsEmpty(tblSelected.DataBodyRange(1, cProdIdx)) Then
        ' Case: Table has placeholder empty row
        startWriteRow = 1
        If countToAdd > 1 Then
            tblSelected.Resize tblSelected.Range.Resize(tblSelected.Range.Rows.Count + countToAdd - 1)
        End If
    Else
        ' Case: Table has data
        startWriteRow = initialRows + 1
        tblSelected.Resize tblSelected.Range.Resize(tblSelected.Range.Rows.Count + countToAdd)
    End If

    ' ** THE MAGIC STEP: Propagate Formulas **
    ' We assume the row ABOVE the new block has the correct formulas.
    ' We FillDown from that row through the new block.
    If startWriteRow > 1 Then
        ' Fill from row above startWriteRow, down to the end of new data
        tblSelected.DataBodyRange.Rows(startWriteRow - 1).Resize(countToAdd + 1).FillDown
    Else
        ' If we are at row 1, formulas should already be there in the placeholder.
        ' But if not, we can't fill down from row 0. We assume the placeholder row is correct.
    End If

    ' --- 7. Write Data Columns ---
    ' Now we overwrite the specific columns with our Hard Values.
    ' The columns NOT in this loop retain the formulas from Step 6.
    
    Dim cName As Variant
    Dim cIndex As Long
    Dim arrayCol As Long
    Dim destRange As Range
    Dim dataSlice As Variant
    
    For arrayCol = 0 To UBound(colNames)
        cName = colNames(arrayCol)
        cIndex = tblSelected.ListColumns(cName).Index
        
        ' 1. Slice the specific column from our 2D array
        ' (Get column arrayCol + 1 because Excel Index is 1-based)
        dataSlice = application.Index(arrData, 0, arrayCol + 1)
        
        ' 2. Define the destination range for just this column
        Set destRange = tblSelected.DataBodyRange.Cells(startWriteRow, cIndex).Resize(countToAdd, 1)
        
        ' 3. Paste the values
        destRange.Value = dataSlice
    Next arrayCol
    
    MsgBox countToAdd & " routines added successfully.", vbInformation, "Bulk Process Complete"

CleanExit:
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error in Bulk Add: " & Err.description, vbCritical
    Resume CleanExit
End Sub

' ==========================================================================
' PRIVATE HELPER: GetMaterialPreparingTemplateData (Unchanged)
' ==========================================================================
Private Function GetMaterialPreparingTemplateData() As Variant
    Dim wsDef As Worksheet, wsDB As Worksheet
    On Error Resume Next
    Set wsDef = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsDB = ThisWorkbook.Sheets("RoutinesDB")
    On Error GoTo 0
    
    If wsDef Is Nothing Or wsDB Is Nothing Then Exit Function
    
    Dim selectedPlant As String
    selectedPlant = Trim(wsDef.Range("C9").Value)
    If selectedPlant <> "1410" And selectedPlant <> "1420" Then Exit Function
    
    Dim tblDB As ListObject
    Set tblDB = wsDB.ListObjects("RoutinesDB")
    Dim arrDB As Variant
    arrDB = tblDB.DataBodyRange.Value
    
    Dim idxPlant As Long: idxPlant = tblDB.ListColumns("Plant").Index
    Dim idxMacro As Long: idxMacro = tblDB.ListColumns("Macrophase").Index
    Dim idxMicro As Long: idxMicro = tblDB.ListColumns("Microphase").Index
    
    Dim i As Long
    For i = 1 To UBound(arrDB, 1)
        If Trim(arrDB(i, idxPlant)) = selectedPlant And _
           Trim(arrDB(i, idxMacro)) = "Stock" And _
           Trim(arrDB(i, idxMicro)) = "Material preparing" Then
            
            Dim res(8) As Variant
            res(0) = arrDB(i, tblDB.ListColumns("Material").Index)
            res(1) = arrDB(i, tblDB.ListColumns("Machine").Index)
            res(2) = arrDB(i, tblDB.ListColumns("Wire/cable dimension diameter/section  (mm/mm2)").Index)
            res(3) = arrDB(i, tblDB.ListColumns("Wire/component dimensions  (mm)").Index)
            res(4) = arrDB(i, tblDB.ListColumns("Work Center Code").Index)
            
            Dim vTr As Variant: vTr = arrDB(i, tblDB.ListColumns("tr").Index)
            Dim vTe As Variant: vTe = arrDB(i, tblDB.ListColumns("te").Index)
            If IsNumeric(vTr) Then res(5) = CDbl(vTr) Else res(5) = 0
            If IsNumeric(vTe) Then res(6) = CDbl(vTe) Else res(6) = 0
            
            res(7) = arrDB(i, tblDB.ListColumns("Sort Order").Index)
            res(8) = selectedPlant
            GetMaterialPreparingTemplateData = res
            Exit Function
        End If
    Next i
End Function

