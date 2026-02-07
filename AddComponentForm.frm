VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddComponentForm 
   Caption         =   "Add New Component"
   ClientHeight    =   3960
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   7860
   OleObjectBlob   =   "AddComponentForm.frx":0000
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
    Dim piecesStr As String
    Dim piecesDbl As Double
    Dim price As Double
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
    If Trim(Me.txtManufacturer.Value) = "" Then
        MsgBox "Manufacturer is required. Please fill it in.", vbExclamation
        Me.txtManufacturer.SetFocus
        Exit Sub
    End If

    If Trim(Me.txtManufacturerPartNumber.Value) = "" Then
        MsgBox "Manufacturer Part Number is required. Please fill it in.", vbExclamation
        Me.txtManufacturerPartNumber.SetFocus
        Exit Sub
    End If

    ' Pieces
    piecesStr = InputBox("Enter the number of pieces needed:", "Number of Pieces")
    If IsNumeric(piecesStr) Then
        piecesDbl = CDbl(piecesStr)
        If piecesDbl <= 0 Then
            MsgBox "Please enter a numeric value greater than 0 for 'Pieces'.", vbExclamation, "Invalid Input"
            Exit Sub
        End If
    Else
        MsgBox "The value entered for 'Pieces' is not a valid number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If
    Debug.Print "Number of pieces " & piecesDbl

    ' Price
    If IsNumeric(Me.txtPrice.Value) Then
        price = CDbl(Me.txtPrice.Value)
        If price <= 0 Then
            MsgBox "Please enter a numeric value greater than 0 for 'Price'.", vbExclamation, "Invalid Input"
            Exit Sub
        End If
    ElseIf Me.txtPrice.Value = "" Then
        ' do later
    Else
        MsgBox "The value entered for 'Price' is not a valid number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If
    Debug.Print "Price " & price

    ' Destination
    Set wsMain = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsDestination = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblDestination = wsDestination.ListObjects("BOMDefinition")

    plantCode = wsMain.Range("C9").Value
    plantName = "Unknown"
    For Each plantRow In tblPlants.ListRows
        If Trim$(CStr(plantRow.Range(tblPlants.ListColumns("Plant").Index).Value)) = Trim$(CStr(plantCode)) Then
            plantName = plantRow.Range(tblPlants.ListColumns("Plant Name").Index).Value
            Exit For
        End If
    Next plantRow

    ' ===================== NAMING: <ProductNumber>-New# =====================
    productNumber = CStr(wsMain.Range("F11").Value)
    prefix = productNumber & "-New"
    pfxLen = Len(prefix)

    highestNewIndex = 0
    Set materialRange = tblDestination.ListColumns("Material").DataBodyRange
    If Not materialRange Is Nothing Then
        For Each cell In materialRange
            txt = CStr(cell.Value)
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
       IsEmpty(tblDestination.ListRows(1).Range(tblDestination.ListColumns("Material").Index).Value) Then
        Set newRow = tblDestination.ListRows(1)
    Else
        Set newRow = tblDestination.ListRows.Add(AlwaysInsert:=True)
    End If

    ' Fill row
    newRow.Range.Cells(1, tblDestination.ListColumns("Base unit of component").Index).Value = Me.txtBaseUnit.Value
    If Me.txtPrice.Value <> "" Then
        newRow.Range.Cells(1, tblDestination.ListColumns("Price").Index).Value = price
        newRow.Range.Cells(1, tblDestination.ListColumns("Price Unit").Index).Value = 1
    End If
    newRow.Range.Cells(1, tblDestination.ListColumns("Condition Currency").Index).Value = "EUR"
    newRow.Range.Cells(1, tblDestination.ListColumns("Product Number").Index).Value = productNumber
    newRow.Range.Cells(1, tblDestination.ListColumns("Plant").Index).Value = wsMain.Range("C9").Text
    newRow.Range.Cells(1, tblDestination.ListColumns("Plant name").Index).Value = plantName
    newRow.Range.Cells(1, tblDestination.ListColumns("Material description").Index).Value = Me.txtDescription.Value
    newRow.Range.Cells(1, tblDestination.ListColumns("Quantity").Index).Value = piecesDbl
    newRow.Range.Cells(1, tblDestination.ListColumns("New component").Index).Value = "NEW"
    newRow.Range.Cells(1, tblDestination.ListColumns("Manufacturer").Index).Value = Me.txtManufacturer.Text
    newRow.Range.Cells(1, tblDestination.ListColumns("Manufacturer Part Number").Index).Value = Me.txtManufacturerPartNumber.Text
    
    ' Set Material value
    materialColIdx = tblDestination.ListColumns("Material").Index
    newRow.Range.Cells(1, materialColIdx).Value = newMaterialName

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
