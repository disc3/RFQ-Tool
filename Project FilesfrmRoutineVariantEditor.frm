VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRoutineVariantEditor 
   Caption         =   "Variant Routines"
   ClientHeight    =   6240
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12390
   OleObjectBlob   =   "Project FilesfrmRoutineVariantEditor.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRoutineVariantEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ======================================
' Variant Routine Editor Form Code (safe copy; only ops changed)
' ======================================

Option Explicit

' ===============================
' Public Properties for Data In
' ===============================
Public baseProduct As String
Public VariantName As String
Public VariantDescription As String

' ---------- helpers ----------
Private Function NormText(v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        NormText = ""
    Else
        NormText = LCase$(Trim$(CStr(v)))
    End If
End Function

Private Function colIdx(tbl As ListObject, ParamArray names() As Variant) As Long
    Dim i As Long, n As Variant
    For Each n In names
        For i = 1 To tbl.ListColumns.Count
            If NormText(tbl.ListColumns(i).name) = NormText(n) Then
                colIdx = i
                Exit Function
            End If
        Next i
    Next n
    colIdx = 0
End Function

Private Function ColumnHasFormula(tbl As ListObject, colIndex As Long) As Boolean
    Dim rng As Range
    Dim hf As Variant

    ColumnHasFormula = False
    If colIndex <= 0 Then Exit Function

    On Error Resume Next
    Set rng = tbl.ListColumns(colIndex).DataBodyRange
    On Error GoTo 0

    If rng Is Nothing Then
        ' Empty table column: treat as not-calculated
        Exit Function
    End If

    hf = rng.HasFormula          ' Can be True/False/Null (mixed)
    If IsNull(hf) Then
        ' Mixed formulas/non-formulas -> treat as calculated to be safe
        ColumnHasFormula = True
    Else
        ColumnHasFormula = CBool(hf)
    End If
End Function


Private Sub SetCellIfWritable(tbl As ListObject, rowObj As ListRow, colIndex As Long, ByVal val As Variant)
    ' Skip if column is calculated or the target cell has a formula
    If colIndex <= 0 Then Exit Sub
    If ColumnHasFormula(tbl, colIndex) Then Exit Sub
    If Not rowObj.Range.Cells(1, colIndex).HasFormula Then
        rowObj.Range.Cells(1, colIndex).value = val
    End If
End Sub

Private Function FindBaseRoutineRow( _
    ByVal tbl As ListObject, _
    ByVal prod As String, _
    ByVal comp As String, _
    ByVal mac As String, _
    ByVal mic As String, _
    ByVal mat As String, _
    ByVal mach As String) As ListRow

    Dim r As ListRow
    Dim cProd As Long, cComp As Long, cMac As Long, cMic As Long, cMat As Long, cMach As Long

    cProd = colIdx(tbl, "Product Number")
    cComp = colIdx(tbl, "Component")
    cMac = colIdx(tbl, "Macrophase")
    cMic = colIdx(tbl, "Microphase")
    cMat = colIdx(tbl, "Material")
    cMach = colIdx(tbl, "Machine")

    For Each r In tbl.ListRows
        If (cProd > 0 And NormText(r.Range.Cells(1, cProd).value) = NormText(prod)) _
        And (cComp = 0 Or NormText(r.Range.Cells(1, cComp).value) = NormText(comp)) _
        And (cMac = 0 Or NormText(r.Range.Cells(1, cMac).value) = NormText(mac)) _
        And (cMic = 0 Or NormText(r.Range.Cells(1, cMic).value) = NormText(mic)) _
        And (cMat = 0 Or NormText(r.Range.Cells(1, cMat).value) = NormText(mat)) _
        And (cMach = 0 Or NormText(r.Range.Cells(1, cMach).value) = NormText(mach)) Then
            Set FindBaseRoutineRow = r
            Exit Function
        End If
    Next r
End Function
' ---------- /helpers ----------

' ======================
' Initialize Routine Form
' ======================
Public Sub InitializeForm()
    txtBaseProduct.Text = baseProduct
    txtVariantName.Text = VariantName
    txtVariantDescription.Text = VariantDescription
    LoadRoutineList baseProduct
End Sub

' ================
' Load Routines
' ================
Private Sub LoadRoutineList(baseProduct As String)
    Dim ws As Worksheet, tbl As ListObject, row As ListRow
    Dim item As listItem
    Dim rowIndex As Long

    On Error GoTo LoadError

    Debug.Print "Loading routines for base product: " & baseProduct

    Set ws = ThisWorkbook.Sheets("2. Routines")
    Set tbl = ws.ListObjects("SelectedRoutines")

    With lvwRoutines
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .ListItems.Clear
        .columnHeaders.Clear

        ' Define columns
        .columnHeaders.Add , , "Component", 80
        .columnHeaders.Add , , "Macrophase", 90
        .columnHeaders.Add , , "Microphase", 90
        .columnHeaders.Add , , "Material", 80
        .columnHeaders.Add , , "Machine", 90
        .columnHeaders.Add , , "Number of Operations", 110

        Dim cProd As Long, cComp As Long, cMac As Long, cMic As Long, cMat As Long, cMach As Long, cOps As Long
        cProd = colIdx(tbl, "Product Number")
        cComp = colIdx(tbl, "Component")
        cMac = colIdx(tbl, "Macrophase")
        cMic = colIdx(tbl, "Microphase")
        cMat = colIdx(tbl, "Material")
        cMach = colIdx(tbl, "Machine")
        cOps = colIdx(tbl, "Number of operations", "Number of Operations")

        rowIndex = 1
        For Each row In tbl.ListRows
            If cProd > 0 And row.Range.Cells(1, cProd).value = baseProduct Then
                Debug.Print "Processing row: " & rowIndex
                Set item = .ListItems.Add(, , row.Range.Cells(1, cComp).value)
                item.SubItems(1) = row.Range.Cells(1, cMac).value
                item.SubItems(2) = row.Range.Cells(1, cMic).value
                item.SubItems(3) = row.Range.Cells(1, cMat).value
                item.SubItems(4) = row.Range.Cells(1, cMach).value
                item.SubItems(5) = IIf(cOps > 0, row.Range.Cells(1, cOps).value, "")
            End If
            rowIndex = rowIndex + 1
        Next row
    End With

    Debug.Print "Routine loading complete."
    Exit Sub

LoadError:
    MsgBox "Error loading routines: " & Err.description, vbCritical, "LoadRoutineList"
    Debug.Print "Error in LoadRoutineList: " & Err.description & " (Row " & rowIndex & ")"
End Sub

Private Sub Label1_Click()
End Sub

' =====================
' Edit Routine Quantity
' =====================
Private Sub lvwRoutines_DblClick()
    Dim selectedItem As listItem
    Dim newValue As Variant
    Dim decimalSeparator As String

    decimalSeparator = application.International(xlDecimalSeparator)

    If lvwRoutines.selectedItem Is Nothing Then
        MsgBox "Please select a routine row first.", vbExclamation
        Exit Sub
    End If

    Set selectedItem = lvwRoutines.selectedItem

    newValue = InputBox("Enter number of operations for " & selectedItem.Text & ":", _
                        "Edit Number of Operations", selectedItem.SubItems(5))

    If InStr(newValue, ".") > 0 And decimalSeparator <> "." Then
        newValue = Replace(newValue, ".", decimalSeparator)
    ElseIf InStr(newValue, ",") > 0 And decimalSeparator <> "," Then
        newValue = Replace(newValue, ",", decimalSeparator)
    End If

    If IsNumeric(newValue) Then
        selectedItem.SubItems(5) = Trim(CStr(CDbl(newValue))) ' force valid number string
    Else
        MsgBox "Please enter a valid number.", vbExclamation
    End If
End Sub

' =====================
' Save Routine Variant
' =====================
Private Sub btnFinalizeVariant_Click()
    Dim ws As Worksheet, tbl As ListObject
    Dim baseRow As ListRow, newRow As ListRow
    Dim colIndex As Long, colName As String
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("2. Routines")
    Set tbl = ws.ListObjects("SelectedRoutines")

    ' cache frequently-used column indices
    Dim cProd As Long, cDesc As Long, cVarOf As Long, cOps As Long
    Dim cComp As Long, cMac As Long, cMic As Long, cMat As Long, cMach As Long

    cProd = colIdx(tbl, "Product Number")
    cDesc = colIdx(tbl, "Product Description")
    cVarOf = colIdx(tbl, "Variant of")
    cOps = colIdx(tbl, "Number of operations", "Number of Operations")
    cComp = colIdx(tbl, "Component")
    cMac = colIdx(tbl, "Macrophase")
    cMic = colIdx(tbl, "Microphase")
    cMat = colIdx(tbl, "Material")
    cMach = colIdx(tbl, "Machine")

    For i = 1 To lvwRoutines.ListItems.Count
        ' 1) locate the exact base row for this operation
        Set baseRow = FindBaseRoutineRow( _
                        tbl, _
                        baseProduct, _
                        lvwRoutines.ListItems(i).Text, _
                        lvwRoutines.ListItems(i).SubItems(1), _
                        lvwRoutines.ListItems(i).SubItems(2), _
                        lvwRoutines.ListItems(i).SubItems(3), _
                        lvwRoutines.ListItems(i).SubItems(4))

        ' 2) add a new variant row
        Set newRow = tbl.ListRows.Add

        ' 3) write core identifiers (skip calculated columns)
        Call SetCellIfWritable(tbl, newRow, cProd, VariantName)
        Call SetCellIfWritable(tbl, newRow, cDesc, VariantDescription)
        Call SetCellIfWritable(tbl, newRow, cVarOf, baseProduct)

        ' keep keys in sync with base row (safer than relying on LV text)
        If Not baseRow Is Nothing Then
            Call SetCellIfWritable(tbl, newRow, cComp, baseRow.Range.Cells(1, cComp).value)
            Call SetCellIfWritable(tbl, newRow, cMac, baseRow.Range.Cells(1, cMac).value)
            Call SetCellIfWritable(tbl, newRow, cMic, baseRow.Range.Cells(1, cMic).value)
            Call SetCellIfWritable(tbl, newRow, cMat, baseRow.Range.Cells(1, cMat).value)
            Call SetCellIfWritable(tbl, newRow, cMach, baseRow.Range.Cells(1, cMach).value)
        Else
            ' fall back to list view values if no base row found
            Call SetCellIfWritable(tbl, newRow, cComp, lvwRoutines.ListItems(i).Text)
            Call SetCellIfWritable(tbl, newRow, cMac, lvwRoutines.ListItems(i).SubItems(1))
            Call SetCellIfWritable(tbl, newRow, cMic, lvwRoutines.ListItems(i).SubItems(2))
            Call SetCellIfWritable(tbl, newRow, cMat, lvwRoutines.ListItems(i).SubItems(3))
            Call SetCellIfWritable(tbl, newRow, cMach, lvwRoutines.ListItems(i).SubItems(4))
        End If

        ' 4) copy every other non-formula column from base row (includes tr/te/etc.)
        If Not baseRow Is Nothing Then
            For colIndex = 1 To tbl.ListColumns.Count
                colName = tbl.ListColumns(colIndex).name
                If colName <> "Product Number" And colName <> "Product Description" And _
                   colName <> "Variant of" And _
                   colName <> "Component" And colName <> "Macrophase" And _
                   colName <> "Microphase" And colName <> "Material" And _
                   colName <> "Machine" And _
                   colName <> "Number of operations" And colName <> "Number of Operations" Then

                    If Not ColumnHasFormula(tbl, colIndex) Then
                        If Not newRow.Range.Cells(1, colIndex).HasFormula Then
                            newRow.Range.Cells(1, colIndex).value = baseRow.Range.Cells(1, colIndex).value
                        End If
                    End If
                End If
            Next colIndex
        End If

        ' 5) finally, override Number of operations with what you edited in the form
        If cOps > 0 Then
            If Not ColumnHasFormula(tbl, cOps) Then
                If Not newRow.Range.Cells(1, cOps).HasFormula Then
                    If IsNumeric(lvwRoutines.ListItems(i).SubItems(5)) Then
                        newRow.Range.Cells(1, cOps).value = CDbl(lvwRoutines.ListItems(i).SubItems(5))
                    Else
                        newRow.Range.Cells(1, cOps).value = lvwRoutines.ListItems(i).SubItems(5)
                    End If
                End If
            End If
        End If
    Next i

    MsgBox "Routine variant '" & VariantName & "' created successfully.", vbInformation
    Unload Me
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


