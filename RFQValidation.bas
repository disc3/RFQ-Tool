Attribute VB_Name = "RFQValidation"
Public Function ValidateAllComponentsAndProducts() As Boolean
    Dim wsComponents As Worksheet
    Dim wsValidation As Worksheet
    Dim wsRoutines As Worksheet
    Dim tblRoutines As ListObject
    Dim tblComponents As ListObject
    Dim tblFinalProductList As ListObject
    Dim FinalProductList As Collection
    Dim componentRow As ListRow
    Dim productRow As ListRow
    Dim productName As String
    Dim expectedComponentCount As Variant
    Dim actualComponentCount As Long
    Dim allFinalProductsValid As Boolean
    Dim componentQuantityValid As Boolean
    Dim componentCostsValid As Boolean
    Dim componentCountValid As Boolean
    Dim detailedStatus As String
    Dim i As Long
    
    Call FixDecimalSeparatorsInTables

    Set wsComponents = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsValidation = ThisWorkbook.Sheets("3. Clarification Validation")
    Set wsRoutines = ThisWorkbook.Sheets("2. Routines")

    On Error Resume Next
    Set tblComponents = wsComponents.ListObjects("BOMDefinition")
    Set tblFinalProductList = ThisWorkbook.Sheets("Final Products").ListObjects("FinalProductList")
    Set tblRoutines = wsRoutines.ListObjects("SelectedRoutines")
    On Error GoTo 0

    If tblComponents Is Nothing Or tblFinalProductList Is Nothing Or tblRoutines Is Nothing Then
        MsgBox "One or more required tables are missing.", vbCritical
        ValidateAllComponentsAndProducts = False
        Exit Function
    End If

    Set FinalProductList = New Collection
    allFinalProductsValid = True
    componentQuantityValid = True
    componentCostsValid = True
    componentCountValid = True
    detailedStatus = ""

    ' Collect unique products from BOMDefinition (skip rows without a real Product Number)
    On Error Resume Next
    For Each componentRow In tblComponents.ListRows
        productName = Trim(componentRow.Range(tblComponents.ListColumns("Product Number").Index).Value)
        If productName <> "" Then FinalProductList.Add productName, CStr(productName)
    Next componentRow
    On Error GoTo 0

    ' Validation 1: All added final products have components defined
    Dim actualFinalProducts As Long: actualFinalProducts = 0
    For Each productRow In tblFinalProductList.ListRows
        productName = Trim(productRow.Range(tblFinalProductList.ListColumns("Product Number").Index).Value)
        If productName <> "" Then actualFinalProducts = actualFinalProducts + 1
    Next productRow

    If actualFinalProducts = FinalProductList.Count Then
        With wsValidation.Range("O14")
            .Value = "Valid"
            .Interior.Color = RGB(0, 255, 0)
        End With
    Else
        With wsValidation.Range("O14")
            .Value = "Invalid"
            .Interior.Color = RGB(255, 0, 0)
        End With
        allFinalProductsValid = False
        detailedStatus = "Mismatch in product/component count."
    End If

    ' Validation 2: Component count check for each product
'    For i = 1 To FinalProductList.Count
'        productName = FinalProductList(i)
'
'        Do
'            expectedComponentCount = InputBox("Enter the number of components for Product " & productName & ":")
'            If IsNumeric(expectedComponentCount) And val(expectedComponentCount) > 0 Then Exit Do
'            MsgBox "Please enter a valid positive number for Product " & productName & ".", vbExclamation
'        Loop
'
'        actualComponentCount = 0
'        For Each componentRow In tblComponents.ListRows
'            If Trim(componentRow.Range(tblComponents.ListColumns("Product Number").Index).value) = productName Then
'                If application.WorksheetFunction.CountA(componentRow.Range) > 1 Then
'                    actualComponentCount = actualComponentCount + 1
'                End If
'            End If
'        Next componentRow
'
'        If actualComponentCount <> CLng(expectedComponentCount) Then
'            MsgBox "Mismatch for Product " & productName & ": Expected " & expectedComponentCount & ", Found " & actualComponentCount, vbExclamation
'            componentCountValid = False
'            detailedStatus = detailedStatus & vbCrLf & "Component count mismatch for product: " & productName
'        End If
'    Next i

    componentCountValid = True

    ' Validation 3: All components have quantity
    For Each componentRow In tblComponents.ListRows
        If IsEmpty(componentRow.Range(tblComponents.ListColumns("Quantity").Index).Value) Or _
           componentRow.Range(tblComponents.ListColumns("Quantity").Index).Value = 0 Then
            componentQuantityValid = False
            detailedStatus = detailedStatus & vbCrLf & "Missing or invalid quantity for product: " & _
                componentRow.Range(tblComponents.ListColumns("Product Number").Index).Value
        End If
    Next componentRow

    With wsValidation.Range("O20")
        .Value = IIf(componentQuantityValid, "Valid", "Invalid")
        .Interior.Color = IIf(componentQuantityValid, RGB(0, 255, 0), RGB(255, 0, 0))
    End With
    
'        If componentCountValid Then
'        With wsValidation.Range("O17")
'            .value = "Valid"
'            .Interior.Color = RGB(0, 255, 0) ' Green
'        End With
'    Else
'        With wsValidation.Range("O17")
'            .value = "Invalid"
'            .Interior.Color = RGB(255, 0, 0) ' Red
'        End With
'    End If

    ' Validation 4: All components have cost
    For Each componentRow In tblComponents.ListRows
        If IsEmpty(componentRow.Range(tblComponents.ListColumns("Price per 1 unit").Index).Value) Or _
           componentRow.Range(tblComponents.ListColumns("Price per 1 unit").Index).Value <= 0 Then
            componentCostsValid = False
            detailedStatus = detailedStatus & vbCrLf & "Missing or zero cost for product: " & _
                componentRow.Range(tblComponents.ListColumns("Product Number").Index).Value
        End If
    Next componentRow

    With wsValidation.Range("O22")
        .Value = IIf(componentCostsValid, "Valid", "Invalid")
        .Interior.Color = IIf(componentCostsValid, RGB(0, 255, 0), RGB(255, 0, 0))
    End With

    ' Validation 5: All Final Products have routines
    Dim routinesValid As Boolean: routinesValid = True
    Dim routineRange As Range
    On Error Resume Next
    Set routineRange = tblRoutines.ListColumns("Product Number").DataBodyRange
    On Error GoTo 0

    If Not routineRange Is Nothing Then
        For Each productRow In tblFinalProductList.ListRows
            productName = Trim(productRow.Range(tblFinalProductList.ListColumns("Product Number").Index).Value)
            If productName = "" Then GoTo SkipRow

            If application.WorksheetFunction.CountIf(routineRange, productName) = 0 Then
                routinesValid = False
                detailedStatus = detailedStatus & vbCrLf & "No routines defined for product: " & productName
            End If
SkipRow:
        Next productRow
    Else
        routinesValid = False
        detailedStatus = detailedStatus & vbCrLf & "No routines table or routines defined for any product."
    End If

    With wsValidation.Range("O24")
        .Value = IIf(routinesValid, "Valid", "Invalid")
        .Interior.Color = IIf(routinesValid, RGB(0, 255, 0), RGB(255, 0, 0))
    End With

    ' Final decision
    ' If allFinalProductsValid And componentCountValid And componentQuantityValid And componentCostsValid And routinesValid Then
    If allFinalProductsValid And componentQuantityValid And componentCostsValid And routinesValid Then
        With wsValidation.Range("J7")
            .Value = "All Products verified!"
            .Interior.Color = RGB(0, 255, 0)
        End With
        ValidateAllComponentsAndProducts = True
    Else
        With wsValidation.Range("J7")
            .Value = "Validation failed. Details:" & vbCrLf & detailedStatus
            .Interior.Color = RGB(255, 0, 0)
        End With
        ValidateAllComponentsAndProducts = False
    End If
End Function



Private Sub btnValidateRFQ_Click()
    ' Button click event to validate the RFQ
    If ValidateAllComponentsAndProducts() Then
        MsgBox "RFQ validation completed successfully!", vbInformation
    Else
        MsgBox "RFQ validation failed. Please check the Validation Status for details.", vbExclamation
    End If
End Sub


