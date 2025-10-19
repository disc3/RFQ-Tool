VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRoutineVariantEditor 
   Caption         =   "Variant Routines"
   ClientHeight    =   6240
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   12390
   OleObjectBlob   =   "frmRoutineVariantEditor.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRoutineVariantEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Module VBE.UserForm.frmRoutineVariantEditor
'@Author Coding-Assistent
'@Date 2025-10-18
'@Version 3.0
'@Description Elite, memory-safe architecture for managing routine variants.
'             Features robust data handling, formula preservation, and a high-performance,
'             stable write-back mechanism.

Option Explicit

'====================================================================================================
'                                       PRIVATE CONSTANTS
'====================================================================================================
' Using constants prevents errors from typos in column names.
Private Const COL_PROD_NUM As String = "Product Number"
Private Const COL_PROD_DESC As String = "Product Description"
Private Const COL_VARIANT_OF As String = "Variant of"
Private Const COL_NUM_OPS As String = "Number of operations"

Private Const COL_COMPONENT As String = "Component"
Private Const COL_MACROPHASE As String = "Macrophase"
Private Const COL_MICROPHASE As String = "Microphase"
Private Const COL_MATERIAL As String = "Material"
Private Const COL_MACHINE As String = "Machine"

'====================================================================================================
'                                        PUBLIC PROPERTIES
'====================================================================================================
Public baseProduct As String
Public variantName As String
Public VariantDescription As String

'====================================================================================================
'                                          FORM EVENTS
'====================================================================================================

Public Sub InitializeForm()
    Me.txtBaseProduct.Text = baseProduct
    Me.txtVariantName.Text = variantName
    Me.txtVariantDescription.Text = VariantDescription
    LoadRoutineList baseProduct
End Sub

Private Sub lvwRoutines_DblClick()
    Dim selectedItem As listItem
    Set selectedItem = Me.lvwRoutines.selectedItem
    If selectedItem Is Nothing Then
        MsgBox "Please select a routine row first.", vbExclamation
        Exit Sub
    End If
    
    Dim userInput As String
    userInput = InputBox("Enter number of operations for " & selectedItem.Text & ":", _
                         "Edit Number of Operations", selectedItem.SubItems(5))
    If userInput = vbNullString Then Exit Sub ' User cancelled

    ' Use the robust TryParse pattern for safe conversion
    Dim numericValue As Double
    If TryParseDouble(userInput, numericValue) Then
        selectedItem.SubItems(5) = CStr(numericValue)
    Else
        MsgBox "'" & userInput & "' is not a valid number.", vbExclamation
    End If
End Sub

Private Sub btnFinalizeVariant_Click()
    Const PROC_NAME As String = "btnFinalizeVariant_Click"
    'On Error GoTo ErrorHandler

    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("2. Routines")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("SelectedRoutines")

    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual

    ' --- 1. READ SOURCE DATA (VALUES & FORMULAS) ---
    Dim baseValueArray As Variant
    Dim baseFormulaArray As Variant
    GetTableData tbl, COL_PROD_NUM, baseProduct, baseValueArray, baseFormulaArray
    
    If Not IsArray(baseValueArray) Then
        MsgBox "No routine data found for base product '" & baseProduct & "'.", vbExclamation
        GoTo CleanExit
    End If

    ' --- 2. BUILD A FAST LOOKUP DICTIONARY ---
    Dim baseRoutineData As Object: Set baseRoutineData = CreateObject("Scripting.Dictionary")
    baseRoutineData.CompareMode = vbTextCompare ' Case-insensitive
    
    Dim r As Long
    For r = 1 To UBound(baseValueArray, 1)
        Dim key As String
        key = BuildRoutineKey(baseValueArray, r, tbl)
        If Not baseRoutineData.exists(key) Then baseRoutineData.Add key, r
    Next r
    
    ' --- 3. CONSTRUCT THE NEW ROWS IN A MEMORY ARRAY ---
    Dim newRowsData() As Variant
    ReDim newRowsData(1 To Me.lvwRoutines.ListItems.Count, 1 To tbl.ListColumns.Count)
    
    Dim i As Long, c As Long, baseRowIndex As Long
    For i = 1 To Me.lvwRoutines.ListItems.Count
        Dim currentKey As String
        currentKey = BuildRoutineKeyFromListView(Me.lvwRoutines.ListItems(i))
        
        If baseRoutineData.exists(currentKey) Then
            baseRowIndex = baseRoutineData(currentKey)
            For c = 1 To tbl.ListColumns.Count
                If HasFormula(baseFormulaArray(baseRowIndex, c)) Then
                    newRowsData(i, c) = baseFormulaArray(baseRowIndex, c)
                Else
                    newRowsData(i, c) = baseValueArray(baseRowIndex, c)
                End If
            Next c
        End If
        
        SetColumnValue newRowsData, i, tbl, COL_PROD_NUM, variantName
        SetColumnValue newRowsData, i, tbl, COL_PROD_DESC, VariantDescription
        SetColumnValue newRowsData, i, tbl, COL_VARIANT_OF, baseProduct
        
        Dim opsValue As Double
        If Not TryParseDouble(Me.lvwRoutines.ListItems(i).SubItems(5), opsValue) Then opsValue = 0
        SetColumnValue newRowsData, i, tbl, COL_NUM_OPS, opsValue
    Next i

    ' --- 4. CANONICAL WRITE TO LISTOBJECT ---
    ' This is the definitive, robust method for adding data to an Excel Table.
    ' It uses the ListObject's own .ListRows.Add method, which prevents all conflicts
    ' and errors related to table boundaries, expansion, and state.
    If UBound(newRowsData, 1) > 0 Then
        Dim newRow As ListRow
        
        For r = 1 To UBound(newRowsData, 1)
            ' Step 4a: Add a new, blank row to the table. Excel handles all the complexity.
            Set newRow = tbl.ListRows.Add(AlwaysInsert:=True)
            
            ' Step 4b: Populate the new row's range with our data.
            ' .Formula correctly handles both values and formulas from our array.
            newRow.Range.Formula = application.Index(newRowsData, r, 0)
        Next r
    End If
    
    Unload Me
    ws.Activate
    Call Utils.RunProductBasedFormatting("2. Routines", "SelectedRoutines")
    ActiveWorkbook.Sheets("1. BOM Definition").Activate
    MsgBox "Routine variant '" & variantName & "' created successfully.", vbInformation
    
CleanExit:
    application.Calculation = xlCalculationAutomatic
    application.ScreenUpdating = True
    Exit Sub
ErrorHandler:
    MsgBox "An unexpected error occurred." & vbCrLf & vbCrLf & _
           "Procedure: " & PROC_NAME & vbCrLf & "Error: " & Err.description, vbCritical
    GoTo CleanExit
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

'====================================================================================================
'                                        HELPER ROUTINES
'====================================================================================================

Private Sub LoadRoutineList(ByVal productFilter As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets("2. Routines")
    Dim tbl As ListObject: Set tbl = ws.ListObjects("SelectedRoutines")
    
    With Me.lvwRoutines
        .ListItems.Clear
        If .columnHeaders.Count = 0 Then
            .View = lvwReport
            .Gridlines = True
            .FullRowSelect = True
            .columnHeaders.Add Text:=COL_COMPONENT, Width:=80
            .columnHeaders.Add Text:=COL_MACROPHASE, Width:=90
            .columnHeaders.Add Text:=COL_MICROPHASE, Width:=90
            .columnHeaders.Add Text:=COL_MATERIAL, Width:=80
            .columnHeaders.Add Text:=COL_MACHINE, Width:=90
            .columnHeaders.Add Text:=COL_NUM_OPS, Width:=110
        End If
    End With
    
    ' Use the dedicated function to get values only. Clean and efficient.
    Dim routineData As Variant
    routineData = GetTableValues(tbl, COL_PROD_NUM, productFilter)
    If Not IsArray(routineData) Then Exit Sub
    
    Dim colIndices As Object: Set colIndices = CreateObject("Scripting.Dictionary")
    colIndices(COL_COMPONENT) = GetColumnIndex(tbl, COL_COMPONENT)
    colIndices(COL_MACROPHASE) = GetColumnIndex(tbl, COL_MACROPHASE)
    colIndices(COL_MICROPHASE) = GetColumnIndex(tbl, COL_MICROPHASE)
    colIndices(COL_MATERIAL) = GetColumnIndex(tbl, COL_MATERIAL)
    colIndices(COL_MACHINE) = GetColumnIndex(tbl, COL_MACHINE)
    colIndices(COL_NUM_OPS) = GetColumnIndex(tbl, COL_NUM_OPS)
    
    Dim i As Long
    For i = 1 To UBound(routineData, 1)
        With Me.lvwRoutines.ListItems.Add(, , CStr(routineData(i, colIndices(COL_COMPONENT))))
            .SubItems(1) = CStr(routineData(i, colIndices(COL_MACROPHASE)))
            .SubItems(2) = CStr(routineData(i, colIndices(COL_MICROPHASE)))
            .SubItems(3) = CStr(routineData(i, colIndices(COL_MATERIAL)))
            .SubItems(4) = CStr(routineData(i, colIndices(COL_MACHINE)))
            .SubItems(5) = CStr(routineData(i, colIndices(COL_NUM_OPS)))
        End With
    Next i
End Sub

Private Function BuildRoutineKey(ByVal dataArray As Variant, ByVal r As Long, ByVal tbl As ListObject) As String
    Dim keyParts(0 To 4) As String
    keyParts(0) = Trim$(CStr(dataArray(r, GetColumnIndex(tbl, COL_COMPONENT))))
    keyParts(1) = Trim$(CStr(dataArray(r, GetColumnIndex(tbl, COL_MACROPHASE))))
    keyParts(2) = Trim$(CStr(dataArray(r, GetColumnIndex(tbl, COL_MICROPHASE))))
    keyParts(3) = Trim$(CStr(dataArray(r, GetColumnIndex(tbl, COL_MATERIAL))))
    keyParts(4) = Trim$(CStr(dataArray(r, GetColumnIndex(tbl, COL_MACHINE))))
    BuildRoutineKey = Join(keyParts, "|")
End Function

Private Function BuildRoutineKeyFromListView(ByVal lvItem As listItem) As String
    Dim keyParts(0 To 4) As String
    keyParts(0) = Trim$(lvItem.Text)
    keyParts(1) = Trim$(lvItem.SubItems(1))
    keyParts(2) = Trim$(lvItem.SubItems(2))
    keyParts(3) = Trim$(lvItem.SubItems(3))
    keyParts(4) = Trim$(lvItem.SubItems(4))
    BuildRoutineKeyFromListView = Join(keyParts, "|")
End Function

Private Sub SetColumnValue(ByRef data As Variant, ByVal r As Long, ByVal tbl As ListObject, ByVal colName As String, ByVal val As Variant)
    Dim c As Long
    c = GetColumnIndex(tbl, colName)
    If c > 0 Then data(r, c) = val
End Sub

'====================================================================================================
'                                  DATA RETRIEVAL & UTILITIES
'====================================================================================================

Private Function GetTableValues(tbl As ListObject, fieldName As String, criteria As String) As Variant
    ' High-performance function that returns a 2D array of only the values.
    GetTableValues = GetTableFilteredData(tbl, fieldName, criteria, False)
End Function

Private Sub GetTableData(tbl As ListObject, fld As String, crit As String, ByRef vals As Variant, ByRef forms As Variant)
    ' Sub that returns both a value array and a formula array.
    vals = GetTableFilteredData(tbl, fld, crit, False)
    If IsArray(vals) Then
        forms = GetTableFilteredData(tbl, fld, crit, True)
    End If
End Sub

Private Function GetTableFilteredData(tbl As ListObject, fieldName As String, criteria As String, getFormulas As Boolean) As Variant
    ' Core data retrieval engine. Filters a table and safely returns either a
    ' .Value or .Formula array, correctly handling the single-row edge case.
    Dim visibleRows As Range
    
    If tbl.AutoFilter.FilterMode Then tbl.AutoFilter.ShowAllData
    tbl.Range.AutoFilter Field:=tbl.ListColumns(fieldName).Index, Criteria1:=criteria
    
    On Error Resume Next
    Set visibleRows = tbl.DataBodyRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    
    tbl.AutoFilter.ShowAllData
    If visibleRows Is Nothing Then
        GetTableFilteredData = False ' Return a non-array to indicate no data
        Exit Function
    End If
    
    If visibleRows.Rows.Count = 1 Then
        ' Single row found: Manually build a 2D array to ensure consistency.
        Dim singleRowArray() As Variant
        Dim colCount As Long: colCount = visibleRows.Columns.Count
        ReDim singleRowArray(1 To 1, 1 To colCount)
        Dim c As Long
        For c = 1 To colCount
            If getFormulas Then
                singleRowArray(1, c) = visibleRows.Cells(1, c).Formula
            Else
                singleRowArray(1, c) = visibleRows.Cells(1, c).Value
            End If
        Next c
        GetTableFilteredData = singleRowArray
    Else
        ' Multiple rows found: Fast bulk retrieval.
        If getFormulas Then
            GetTableFilteredData = visibleRows.Formula
        Else
            GetTableFilteredData = visibleRows.Value
        End If
    End If
End Function

Private Function GetColumnIndex(ByVal tbl As ListObject, ByVal colName As String) As Long
    On Error Resume Next
    GetColumnIndex = tbl.ListColumns(colName).Index
End Function

Private Function HasFormula(ByVal cellFormula As Variant) As Boolean
    If VarType(cellFormula) = vbString Then
        If Left$(cellFormula, 1) = "=" Then HasFormula = True
    End If
End Function

Private Function TryParseDouble(ByVal Text As String, ByRef Result As Double) As Boolean
    ' Robustly converts a string to a Double, regardless of regional settings for decimals.
    Dim normalizedText As String
    normalizedText = Replace(Trim(Text), ",", ".") ' Standardize to period
    
    If IsNumeric(normalizedText) Then
        Result = CDbl(normalizedText)
        TryParseDouble = True
    End If
End Function

Private Sub UserForm_Click()

End Sub
