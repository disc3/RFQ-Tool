VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVariantConfigurator 
   Caption         =   "Create Variant"
   ClientHeight    =   7140
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12465
   OleObjectBlob   =   "Project FilesfrmVariantConfigurator.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmVariantConfigurator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ==== helpers ====
Private Function VariantExistsInBOM(ByVal productNumber As String) As Boolean
    Dim tbl As ListObject, rng As Range, f As Range
    If Len(Trim$(productNumber)) = 0 Then Exit Function
    Set tbl = ThisWorkbook.Sheets("1. BOM Definition").ListObjects("BOMDefinition")
    On Error Resume Next
    Set rng = tbl.ListColumns("Product Number").DataBodyRange
    On Error GoTo 0
    If rng Is Nothing Then Exit Function
    Set f = rng.Find(What:=productNumber, LookIn:=xlValues, LookAt:=xlWhole)
    VariantExistsInBOM = Not f Is Nothing
End Function

Private Function NextFreeVariantName(ByVal baseProduct As String, ByVal tbl As ListObject) As String
    Dim n As Long
    n = GetNextVariantNumber(baseProduct, tbl) ' uses your existing function
    NextFreeVariantName = baseProduct & "-V" & n
    ' Just in case: bump until free
    Do While VariantExistsInBOM(NextFreeVariantName)
        n = n + 1
        NextFreeVariantName = baseProduct & "-V" & n
    Loop
End Function
' ==== /helpers ====

Private Sub UserForm_Initialize()
    Dim ws As Worksheet, tbl As ListObject, row As ListRow
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim prodNum As Variant, prodDesc As Variant
    Dim i As Long: i = 0

    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")

    cmbBaseProduct.Clear
    cmbBaseProduct.ColumnCount = 2
    cmbBaseProduct.ColumnWidths = "100;150"

    For Each row In tbl.ListRows
        prodNum = row.Range(tbl.ListColumns("Product Number").Index).value
        prodDesc = row.Range(tbl.ListColumns("Product Description").Index).value

        If Not IsError(prodNum) And Not IsError(prodDesc) Then
            If Trim$(CStr(prodNum)) <> "" And Not dict.exists(CStr(prodNum)) Then
                dict.Add CStr(prodNum), prodDesc
                cmbBaseProduct.AddItem
                cmbBaseProduct.List(i, 0) = prodNum
                cmbBaseProduct.List(i, 1) = prodDesc
                i = i + 1
            End If
        End If
    Next row
End Sub

Private Sub btnLoadComponents_Click()
    If cmbBaseProduct.ListIndex = -1 Then
        MsgBox "Please select a base product.", vbExclamation
        Exit Sub
    End If

    txtBaseProductDesc.Text = cmbBaseProduct.List(cmbBaseProduct.ListIndex, 1)
    LoadComponentList cmbBaseProduct.List(cmbBaseProduct.ListIndex, 0)
End Sub

Sub LoadComponentList(baseProduct As String)
    Dim ws As Worksheet, tbl As ListObject, row As ListRow
    Dim mat As Variant, desc As Variant, qty As Variant, uom As Variant, manuf As Variant, price As Variant
    Dim item As listItem

    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")

    With lvwComponents
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .ListItems.Clear
        .columnHeaders.Clear

        .columnHeaders.Add , , "Material", 90
        .columnHeaders.Add , , "Material Description", 150
        .columnHeaders.Add , , "Quantity", 60
        .columnHeaders.Add , , "Base unit of component", 70
        .columnHeaders.Add , , "Vendor Name", 100
        .columnHeaders.Add , , "Price per 1 unit", 90

        For Each row In tbl.ListRows
            If row.Range(tbl.ListColumns("Product Number").Index).value = baseProduct Then
                mat = row.Range(tbl.ListColumns("Material").Index).value
                desc = row.Range(tbl.ListColumns("Material Description").Index).value
                qty = row.Range(tbl.ListColumns("Quantity").Index).value
                uom = row.Range(tbl.ListColumns("Base unit of component").Index).value
                manuf = row.Range(tbl.ListColumns("Vendor Name").Index).value
                price = row.Range(tbl.ListColumns("Price per 1 unit").Index).value

                Set item = .ListItems.Add(, , mat)
                item.SubItems(1) = desc
                item.SubItems(2) = qty
                item.SubItems(3) = uom
                item.SubItems(4) = manuf
                item.SubItems(5) = price
            End If
        Next row
    End With
End Sub

Private Sub cmbBaseProduct_Change()
    Dim selectedProduct As String, nextVariant As Long, tbl As ListObject, pn As String

    If cmbBaseProduct.ListIndex = -1 Then Exit Sub

    selectedProduct = cmbBaseProduct.List(cmbBaseProduct.ListIndex, 0)
    txtBaseProductDesc.Text = cmbBaseProduct.List(cmbBaseProduct.ListIndex, 1)

    LoadComponentList selectedProduct

    Set tbl = ThisWorkbook.Sheets("1. BOM Definition").ListObjects("BOMDefinition")
    nextVariant = GetNextVariantNumber(selectedProduct, tbl)
    pn = selectedProduct & "-V" & nextVariant
    ' ensure it's free
    pn = NextFreeVariantName(selectedProduct, tbl)

    txtVariantName.Text = pn
    txtVariantDescription.Text = selectedProduct & " | Modified variant"
End Sub

Function GetNextVariantNumber(baseProduct As String, tbl As ListObject) As Long
    Dim cell As Range, prefix As String, suffix As String
    Dim maxNum As Long, varNum As Long

    prefix = LCase$(baseProduct) & "-v"
    maxNum = 0

    If Not tbl.ListColumns("Product Number").DataBodyRange Is Nothing Then
        For Each cell In tbl.ListColumns("Product Number").DataBodyRange
            If LCase$(Left$(CStr(cell.value), Len(prefix))) = prefix Then
                suffix = Mid$(CStr(cell.value), Len(prefix) + 1)
                If IsNumeric(suffix) Then
                    varNum = CLng(suffix)
                    If varNum > maxNum Then maxNum = varNum
                End If
            End If
        Next cell
    End If

    GetNextVariantNumber = maxNum + 1
End Function

Private Sub txtVariantName_Change()
    ' live feedback if PN already exists in BOM
    If VariantExistsInBOM(txtVariantName.Text) Then
        txtVariantName.BackColor = RGB(255, 230, 230) ' light red
    Else
        txtVariantName.BackColor = vbWhite
    End If
End Sub

Private Sub lvwComponents_DblClick()
    Dim selectedItem As listItem
    Dim newValue As Variant
    Dim decimalSeparator As String

    decimalSeparator = application.International(xlDecimalSeparator)

    If lvwComponents.selectedItem Is Nothing Then
        MsgBox "Please select a component first.", vbExclamation
        Exit Sub
    End If

    Set selectedItem = lvwComponents.selectedItem

    newValue = InputBox("Enter new quantity for material " & selectedItem.Text & ":", _
                        "Edit Quantity", selectedItem.SubItems(2))

    If InStr(newValue, ".") > 0 And decimalSeparator <> "." Then
        newValue = Replace(newValue, ".", decimalSeparator)
    ElseIf InStr(newValue, ",") > 0 And decimalSeparator <> "," Then
        newValue = Replace(newValue, ",", decimalSeparator)
    End If

    If IsNumeric(newValue) Then
        selectedItem.SubItems(2) = newValue
    Else
        MsgBox "Please enter a valid number.", vbExclamation
    End If
End Sub

Private Sub btnCreateVariant_Click()
    Dim tblBOM As ListObject, wsBom As Worksheet
    Dim tblProducts As ListObject, wsProducts As Worksheet
    Dim i As Long, colIndex As Long
    Dim newRow As ListRow, prodRow As ListRow, baseRow As ListRow
    Dim baseProduct As String, VariantName As String, variantDesc As String
    Dim colName As String, qtyVal As Variant
    Dim resp As VbMsgBoxResult, altName As String

    On Error GoTo ErrorHandler

    If cmbBaseProduct.ListIndex = -1 Then
        MsgBox "Select a base product first.", vbExclamation
        Exit Sub
    End If

    baseProduct = cmbBaseProduct.List(cmbBaseProduct.ListIndex, 0)
    VariantName = Trim$(txtVariantName.Text)
    variantDesc = txtVariantDescription.Text

    If VariantName = "" Then
        MsgBox "Please enter a variant product name.", vbExclamation
        Exit Sub
    End If

    Set wsBom = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblBOM = wsBom.ListObjects("BOMDefinition")
    Set wsProducts = ThisWorkbook.Sheets("Final Products")
    Set tblProducts = wsProducts.ListObjects("FinalProductList")

    ' ---- Duplicate PN guard (BOM) ----
    If VariantExistsInBOM(VariantName) Then
        altName = NextFreeVariantName(baseProduct, tblBOM)
        resp = MsgBox("Product Number '" & VariantName & "' already exists in BOM." & vbCrLf & _
                      "Use next available '" & altName & "' instead?", _
                      vbExclamation + vbYesNoCancel, "Duplicate Product Number")
        If resp = vbYes Then
            VariantName = altName
            txtVariantName.Text = altName
        Else
            Exit Sub ' No / Cancel -> let user change it
        End If
    End If

    ' ==== create BOM rows ====
    For i = 1 To lvwComponents.ListItems.Count
        qtyVal = lvwComponents.ListItems(i).SubItems(2)

        If Not IsError(qtyVal) And IsNumeric(qtyVal) And Trim$(CStr(qtyVal)) <> "" And val(qtyVal) <> 0 Then
            Set newRow = tblBOM.ListRows.Add

            For colIndex = 1 To tblBOM.ListColumns.Count
                colName = tblBOM.ListColumns(colIndex).name

                With newRow.Range(colIndex)
                    If Not .HasFormula Then
                        Select Case colName
                            Case "Product Number":              .value = VariantName
                            Case "Product Description":         .value = variantDesc
                            Case "Variant of":                  .value = baseProduct
                            Case "Material":                    .value = lvwComponents.ListItems(i).Text
                            Case "Material Description":        .value = lvwComponents.ListItems(i).SubItems(1)
                            Case "Quantity":                    .value = lvwComponents.ListItems(i).SubItems(2)
                            Case "Base unit of component":      .value = lvwComponents.ListItems(i).SubItems(3)
                            Case "Vendor Name":                 .value = lvwComponents.ListItems(i).SubItems(4)
                            Case "Price per 1 unit":            .value = lvwComponents.ListItems(i).SubItems(5)
                            Case Else
                                For Each baseRow In tblBOM.ListRows
                                    If baseRow.Range(tblBOM.ListColumns("Product Number").Index).value = baseProduct _
                                    And baseRow.Range(tblBOM.ListColumns("Material").Index).value = lvwComponents.ListItems(i).Text Then
                                        .value = baseRow.Range(colIndex).value
                                        Exit For
                                    End If
                                Next baseRow
                        End Select
                    End If
                End With
            Next colIndex
        End If
    Next i

    ' ==== add to Final Products ====
    For Each baseRow In tblProducts.ListRows
        If baseRow.Range(tblProducts.ListColumns("Product Number").Index).value = baseProduct Then Exit For
    Next baseRow

    Set prodRow = tblProducts.ListRows.Add
    For colIndex = 1 To tblProducts.ListColumns.Count
        With prodRow.Range(colIndex)
            If tblProducts.ListColumns(colIndex).name = "Product Number" Then
                .value = VariantName
            ElseIf tblProducts.ListColumns(colIndex).name = "Product Description" Then
                .value = variantDesc
            ElseIf tblProducts.ListColumns(colIndex).name = "Variant of" Then
                .value = baseProduct
            ElseIf Not .HasFormula Then
                .value = baseRow.Range(colIndex).value
            End If
        End With
    Next colIndex

    ' Open routine editor
    Dim frmRoutine As New frmRoutineVariantEditor
    Unload Me
    With frmRoutine
        .baseProduct = baseProduct
        .VariantName = VariantName
        .VariantDescription = variantDesc
        .InitializeForm
        .Show
    End With
    Exit Sub

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.description & vbCrLf & _
           "In line: " & i, vbCritical, "Create Variant"
    Debug.Print "Error " & Err.Number & ": " & Err.description
    Resume Next
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub Label1_Click()
End Sub

Private Sub Label2_Click()
End Sub


