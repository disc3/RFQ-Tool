Attribute VB_Name = "ServoCalculation"
Dim originalRoutineQuantities() As Variant
Dim routineRowIndexes() As Long

Public Sub RunServoCalculation()
    Dim wsBom As Worksheet, wsSales As Worksheet, wsServo As Worksheet
    Dim tblBOM As ListObject
    Dim selectedProduct As String, cableMaterial As String
    Dim labConPrice As Double, cableCostPerM As Double
    Dim copperWeightPerKm As Double, netWeightPerM As Double
    Dim otherComponentWeight As Double
    Dim rowRange As Range, cell As Range
    Dim cableQuantityCell As Range, originalQuantity As Variant
    Dim productRow As Range
    Dim i As Long
    
    Call FixDecimalSeparatorsInTables
    
    Set wsBom = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsSales = ThisWorkbook.Sheets("4. Sales Calculation (Internal)")
    Set wsServo = ThisWorkbook.Sheets("8. Servo calculation")
    Set tblBOM = wsBom.ListObjects("BOMDefinition")

    selectedProduct = Trim(wsBom.Range("F11").value)
    If selectedProduct = "" Then
        MsgBox "Please select a product in cell F11 first.", vbExclamation
        Exit Sub
    End If

    ' Select cable
    Dim cableForm As ServoCableSelectForm
    Set cableForm = New ServoCableSelectForm
    cableForm.Show
    If cableForm.SelectedCableMaterial = "" Then
        MsgBox "No cable component selected.", vbExclamation
        Exit Sub
    End If
    cableMaterial = cableForm.SelectedCableMaterial
    Unload cableForm

    ' Prepare quantity override
    Dim qtyCol As Long
    qtyCol = tblBOM.ListColumns("Quantity").Index

    ' Store original quantities
    Dim rowCount As Long: rowCount = tblBOM.DataBodyRange.Rows.Count
    Dim originalQuantities() As Variant
    ReDim originalQuantities(1 To rowCount)
    For i = 1 To rowCount
        originalQuantities(i) = tblBOM.DataBodyRange.Cells(i, qtyCol).value
    Next i

    ' --- SIMULATION 1: 0.1m cable + work ---
    For i = 1 To rowCount
        With tblBOM.DataBodyRange.Rows(i)
            If .Cells(tblBOM.ListColumns("Product Number").Index).value = selectedProduct And _
               .Cells(tblBOM.ListColumns("Material").Index).value = cableMaterial Then
                .Cells(qtyCol).value = 0.1
            End If
        End With
    Next i

    application.Calculate
    application.Wait Now + TimeValue("00:00:01")
    labConPrice = GetPriceFromSales(wsSales, selectedProduct, "Selling Price/ Transfer Price Copper Base / 1 piece")

    ' --- SIMULATION 2: 1m cable only ---
    Call ZeroRoutineQuantities(selectedProduct)
    For i = 1 To rowCount
        With tblBOM.DataBodyRange.Rows(i)
            If .Cells(tblBOM.ListColumns("Product Number").Index).value = selectedProduct Then
                If .Cells(tblBOM.ListColumns("Material").Index).value = cableMaterial Then
                    .Cells(qtyCol).value = 1
                Else
                    .Cells(qtyCol).value = 0
                End If
            End If
        End With
    Next i

    application.Calculate
    application.Wait Now + TimeValue("00:00:01")
    cableCostPerM = GetPriceFromSales(wsSales, selectedProduct, "Selling Price/ Transfer Price Copper Base / 1 piece")

    ' Restore routine quantities and BOM quantities
    Call RestoreRoutineQuantities
    For i = 1 To rowCount
        tblBOM.DataBodyRange.Cells(i, qtyCol).value = originalQuantities(i)
    Next i
    application.Calculate

    ' Retrieve weights and identifiers
    copperWeightPerKm = 0: netWeightPerM = 0: otherComponentWeight = 0
    Dim desc As String, erpPN As String
    For Each rowRange In tblBOM.DataBodyRange.Rows
        If Trim(rowRange.Cells(tblBOM.ListColumns("Product Number").Index).value) = selectedProduct Then
            If Trim(rowRange.Cells(tblBOM.ListColumns("Material").Index).value) = cableMaterial Then
                copperWeightPerKm = val(rowRange.Cells(tblBOM.ListColumns("Copper weight [kg/1000m]").Index).value)
                netWeightPerM = val(rowRange.Cells(tblBOM.ListColumns("Net weight (kg/Base unit)").Index).value)
                desc = rowRange.Cells(tblBOM.ListColumns("Product Description").Index).value
                erpPN = rowRange.Cells(tblBOM.ListColumns("ERP Part Number").Index).value
            Else
                otherComponentWeight = otherComponentWeight + val(rowRange.Cells(tblBOM.ListColumns("Net weight (kg/Base unit)").Index).value)
            End If
        End If
    Next rowRange

    ' Clear outputs
    wsServo.Range("A4:I1000").ClearContents
    wsServo.Range("K4:R1000").ClearContents

    ' Output to sheet
    Dim outRow As Long: outRow = 4
    Dim lengthVal As Double
    For lengthVal = 0.5 To 100 Step 0.5
        With wsServo
            .Cells(outRow, 1).value = selectedProduct
            .Cells(outRow, 2).value = desc
            .Cells(outRow, 3).value = erpPN
            .Cells(outRow, 4).value = desc
            .Cells(outRow, 5).value = lengthVal
            .Cells(outRow, 6).value = labConPrice + ((lengthVal) * cableCostPerM)
            .Cells(outRow, 7).value = lengthVal * copperWeightPerKm
            .Cells(outRow, 8).value = 150
            .Cells(outRow, 9).value = lengthVal * netWeightPerM + otherComponentWeight

            .Cells(outRow, 11).value = erpPN
            .Cells(outRow, 12).value = desc
            .Cells(outRow, 13).value = labConPrice
            .Cells(outRow, 14).value = cableCostPerM
            .Cells(outRow, 15).value = cableMaterial
            .Cells(outRow, 16).value = netWeightPerM
            .Cells(outRow, 17).value = otherComponentWeight
            .Cells(outRow, 18).value = copperWeightPerKm
        End With
        outRow = outRow + 1
    Next lengthVal

    wsServo.Activate
End Sub

Function GetPriceFromSales(wsSales As Worksheet, product As String, columnHeader As String) As Double
    Dim headerRow As Range: Set headerRow = wsSales.Rows(14)
    Dim i As Long, colIndex As Long: colIndex = 0
    For i = 1 To headerRow.Columns.Count
        If Trim(CStr(headerRow.Cells(1, i).value)) = columnHeader Then
            colIndex = i
            Exit For
        End If
    Next i
    If colIndex = 0 Then
        MsgBox "Column '" & columnHeader & "' not found.", vbCritical
        GetPriceFromSales = 0
        Exit Function
    End If
    Dim r As Range
    For Each r In wsSales.Range("B15:B2000")
        If Trim(CStr(r.value)) = product Then
            GetPriceFromSales = r.Cells(1, colIndex - 1).value
            Exit Function
        End If
    Next r
    MsgBox "Product '" & product & "' not found in Sales Calculation sheet.", vbExclamation
    GetPriceFromSales = 0
End Function

Sub ZeroRoutineQuantities(product As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long, n As Long
    Set ws = ThisWorkbook.Sheets("2. Routines")
    Set tbl = ws.ListObjects("SelectedRoutines")

    Dim qtyCol As Long
    On Error Resume Next
    qtyCol = tbl.ListColumns("Number of operations").Index
    On Error GoTo 0
    If qtyCol = 0 Then
        MsgBox "Cannot find 'Number of operations' column in SelectedRoutines.", vbCritical
        Exit Sub
    End If

    n = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(tbl.DataBodyRange.Cells(i, tbl.ListColumns("Product Number").Index).value) = product Then
            n = n + 1
        End If
    Next i

    ReDim originalRoutineQuantities(1 To n)
    ReDim routineRowIndexes(1 To n)

    n = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(tbl.DataBodyRange.Cells(i, tbl.ListColumns("Product Number").Index).value) = product Then
            n = n + 1
            routineRowIndexes(n) = i
            originalRoutineQuantities(n) = tbl.DataBodyRange.Cells(i, qtyCol).value
            tbl.DataBodyRange.Cells(i, qtyCol).value = 0
        End If
    Next i
End Sub

Sub RestoreRoutineQuantities()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("2. Routines")
    Set tbl = ws.ListObjects("SelectedRoutines")

    If Not IsEmpty(originalRoutineQuantities) Then
        For i = LBound(routineRowIndexes) To UBound(routineRowIndexes)
            tbl.DataBodyRange.Cells(routineRowIndexes(i), tbl.ListColumns("Number of operations").Index).value = originalRoutineQuantities(i)
        Next i
    End If
End Sub


