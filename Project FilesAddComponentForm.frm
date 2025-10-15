VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddComponentForm 
   Caption         =   "Add New Component"
   ClientHeight    =   3960
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7860
   OleObjectBlob   =   "Project FilesAddComponentForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "AddComponentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub btnAddComponent_Click()
    Dim wsDestination As Worksheet
    Dim tblDestination As ListObject
    Dim newRow As ListRow
    Dim pieces As String, price As String
    Dim wsMain As Worksheet
    Dim newMaterialName As String
    Dim highestNewIndex As Long
    Dim cell As Range
    Dim materialRange As Range
    Dim wsPlantVariables As Worksheet
    Dim tblPlants As ListObject
    Dim plantCode As String
    Dim plantName As Variant
    Dim productNumber As String
    Dim prefix As String
    Dim txt As String, sufTxt As String, pfxLen As Long
    Dim plantRow As ListRow
    Dim materialColIdx As Long
    Dim foundCell As Range

    Set wsPlantVariables = ThisWorkbook.Sheets("Plant Variables")
    Set tblPlants = wsPlantVariables.ListObjects("PlantVariables")

    ' --- Mandatory fields ---
    If Trim(Me.txtManufacturer.value) = "" Then
        MsgBox "Manufacturer is required. Please fill it in.", vbExclamation
        Me.txtManufacturer.SetFocus
        Exit Sub
    End If

    If Trim(Me.txtManufacturerPartNumber.value) = "" Then
        MsgBox "Manufacturer Part Number is required. Please fill it in.", vbExclamation
        Me.txtManufacturerPartNumber.SetFocus
        Exit Sub
    End If

    ' Pieces
    pieces = InputBox("Enter the number of pieces needed:", "Number of Pieces")
    pieces = FixDecimalSeparator(pieces)
    Debug.Print "Number of pieces " & pieces

    ' Price
    price = Me.txtPrice.value
    price = FixDecimalSeparator(price)

    ' Destination
    Set wsMain = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsDestination = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblDestination = wsDestination.ListObjects("BOMDefinition")

    plantCode = wsMain.Range("C9").value
    plantName = "Unknown"
    For Each plantRow In tblPlants.ListRows
        If Trim$(CStr(plantRow.Range(tblPlants.ListColumns("Plant").Index).value)) = Trim$(CStr(plantCode)) Then
            plantName = plantRow.Range(tblPlants.ListColumns("Plant Name").Index).value
            Exit For
        End If
    Next plantRow

    ' ===================== NAMING: <ProductNumber>-New# =====================
    productNumber = CStr(wsMain.Range("F11").value)
    prefix = productNumber & "-New"
    pfxLen = Len(prefix)

    highestNewIndex = 0
    Set materialRange = tblDestination.ListColumns("Material").DataBodyRange
    If Not materialRange Is Nothing Then
        For Each cell In materialRange
            txt = CStr(cell.value)
            If Len(txt) >= pfxLen Then
                If LCase$(Left$(txt, pfxLen)) = LCase$(prefix) Then
                    sufTxt = Mid$(txt, pfxLen + 1)
                    If IsNumeric(sufTxt) Then
                        If CLng(sufTxt) > highestNewIndex Then highestNewIndex = CLng(sufTxt)
                    End If
                End If
            End If
        Next cell
    End If

    If highestNewIndex = 0 Then
        newMaterialName = prefix & "1"
    Else
        newMaterialName = prefix & CStr(highestNewIndex + 1)
    End If
    ' =======================================================================

    ' Reuse single empty first row or add a new row
    If tblDestination.ListRows.Count = 1 And _
       IsEmpty(tblDestination.ListRows(1).Range(tblDestination.ListColumns("Material").Index).value) Then
        Set newRow = tblDestination.ListRows(1)
    Else
        Set newRow = tblDestination.ListRows.Add(AlwaysInsert:=True)
    End If

    ' Fill row
    newRow.Range.Cells(1, tblDestination.ListColumns("Base unit of component").Index).value = Me.txtBaseUnit.value
    newRow.Range.Cells(1, tblDestination.ListColumns("Price per 1 unit").Index).value = price
    newRow.Range.Cells(1, tblDestination.ListColumns("Condition Currency").Index).value = "EUR"
    newRow.Range.Cells(1, tblDestination.ListColumns("Product Number").Index).value = productNumber
    newRow.Range.Cells(1, tblDestination.ListColumns("Plant").Index).value = wsMain.Range("C9").value
    newRow.Range.Cells(1, tblDestination.ListColumns("Plant name").Index).value = plantName
    newRow.Range.Cells(1, tblDestination.ListColumns("Material Description").Index).value = Me.txtDescription.value
    newRow.Range.Cells(1, tblDestination.ListColumns("Quantity").Index).value = val(pieces)
    newRow.Range.Cells(1, tblDestination.ListColumns("New component").Index).value = "NEW"
    newRow.Range.Cells(1, tblDestination.ListColumns("Manufacturer").Index).value = Me.txtManufacturer.Text
    newRow.Range.Cells(1, tblDestination.ListColumns("Manufacturer Part Number").Index).value = Me.txtManufacturerPartNumber.Text
    
    ' Set Material value
    materialColIdx = tblDestination.ListColumns("Material").Index
    newRow.Range.Cells(1, materialColIdx).value = newMaterialName

    ' === HIGHLIGHT the Material cell in yellow (pre-format) ===
    newRow.Range.Cells(1, materialColIdx).Interior.Color = vbYellow

    ' Apply your standard formatting
    'Utils.ApplyRowFormatting newRow

    ' Sort components by Product
    SortSelectedComponentsByProduct

    ' Re-apply highlight after sorting (find by the generated material name)
    Set foundCell = Nothing
    On Error Resume Next
    Set foundCell = tblDestination.ListColumns("Material").DataBodyRange.Find( _
                        What:=newMaterialName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0
    If Not foundCell Is Nothing Then
        foundCell.Interior.Color = vbYellow
    End If

    ' Close the form
    Unload Me
End Sub


Private Sub btnCancel_Click()
    ' Close the form without doing anything
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

