Attribute VB_Name = "ManualOverrides"
Option Explicit

'================================================================================
' MODULE: ManualOverrides
' PURPOSE: Tracks manual edits to protected columns in BOMDefinition so that
'          "Update Components" (RefreshBOMData) does not overwrite them.
'
' HOW IT WORKS:
'   - A hidden sheet "ManualOverrides" stores a table with columns:
'     Material | Plant | ProductNumber | ColumnHeader | OverrideValue
'   - When a user manually edits a protected column, the Worksheet_Change event
'     on the BOM Definition sheet calls RecordOverride to log the change.
'   - When RefreshBOMData runs, it loads the overrides into a Dictionary and
'     skips writing to any cell that has a recorded override.
'   - The SuppressChangeTracking flag is set to True before any programmatic
'     writes, ensuring that only genuine manual user edits are tracked.
'================================================================================

' --- Module-level suppression flag ---
' Set this to True before any macro/programmatic writes to BOMDefinition,
' and back to False afterwards. This ensures the Worksheet_Change event
' only records genuinely manual edits.
Public SuppressChangeTracking As Boolean

Private Const OVERRIDES_SHEET_NAME As String = "ManualOverrides"
Private Const OVERRIDES_TABLE_NAME As String = "ManualOverridesTable"

'================================================================================
' CONFIGURABLE LIST OF PROTECTED COLUMNS
' Add or remove column names here to control which columns are tracked.
' These are the columns that Purchasing may edit manually and that should
' NOT be overwritten by "Update Components".
'================================================================================
Private Function GetProtectedColumns() As Variant
    GetProtectedColumns = Array( _
        "Copper Weight [kg/1000m]", _
        "MOQ", _
        "Planned delivery time", _
        "Price", _
        "Price Unit" _
    )
End Function

'''
' Checks whether a given column name is in the protected list.
' @param {String} colName The column header name to check.
' @return {Boolean} True if the column is protected.
'''
Public Function IsProtectedColumn(ByVal colName As String) As Boolean
    Dim cols As Variant
    Dim i As Long
    cols = GetProtectedColumns()
    For i = LBound(cols) To UBound(cols)
        If cols(i) = colName Then
            IsProtectedColumn = True
            Exit Function
        End If
    Next i
    IsProtectedColumn = False
End Function

'''
' Ensures the ManualOverrides hidden sheet and table exist.
' Creates them if they are missing.
' @return {ListObject} The ManualOverridesTable ListObject.
'''
Public Function EnsureOverridesTable() As ListObject
    Dim ws As Worksheet
    Dim lo As ListObject

    ' Try to get existing sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(OVERRIDES_SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        ' Create the sheet
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.name = OVERRIDES_SHEET_NAME
        ws.Visible = xlSheetVeryHidden

        ' Set up headers (5 columns: Material, Plant, ProductNumber, ColumnHeader, OverrideValue)
        ws.Range("A1").Value = "Material"
        ws.Range("B1").Value = "Plant"
        ws.Range("C1").Value = "ProductNumber"
        ws.Range("D1").Value = "ColumnHeader"
        ws.Range("E1").Value = "OverrideValue"

        ' Create table from headers
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E1"), , xlYes)
        lo.name = OVERRIDES_TABLE_NAME
    Else
        ' Ensure the sheet is very hidden
        If ws.Visible <> xlSheetVeryHidden Then ws.Visible = xlSheetVeryHidden

        ' Get existing table
        On Error Resume Next
        Set lo = ws.ListObjects(OVERRIDES_TABLE_NAME)
        On Error GoTo 0

        ' Check if table exists and has the correct schema (5 columns with ProductNumber)
        Dim needsRecreate As Boolean
        needsRecreate = False
        If lo Is Nothing Then
            needsRecreate = True
        Else
            ' Migrate from old 4-column format to new 5-column format
            On Error Resume Next
            Dim testCol As ListColumn
            Set testCol = lo.ListColumns("ProductNumber")
            On Error GoTo 0
            If testCol Is Nothing Then needsRecreate = True
        End If

        If needsRecreate Then
            ' Clear and recreate with the new 5-column schema
            If Not lo Is Nothing Then
                If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
                lo.Delete
            End If
            ws.Cells.Clear
            ws.Range("A1").Value = "Material"
            ws.Range("B1").Value = "Plant"
            ws.Range("C1").Value = "ProductNumber"
            ws.Range("D1").Value = "ColumnHeader"
            ws.Range("E1").Value = "OverrideValue"
            Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:E1"), , xlYes)
            lo.name = OVERRIDES_TABLE_NAME
        End If
    End If

    Set EnsureOverridesTable = lo
End Function

'''
' Records (upserts) a manual override for a specific Material + Plant + ProductNumber + Column.
' If an entry already exists for the same key, its value is updated.
' @param {String} material The material number.
' @param {String} plant The plant code.
' @param {String} productNumber The product number (row-level identifier).
' @param {String} colHeader The column header name that was changed.
' @param {Variant} overrideValue The new value entered by the user.
'''
Public Sub RecordOverride(ByVal material As String, ByVal plant As String, _
                          ByVal productNumber As String, ByVal colHeader As String, _
                          ByVal overrideValue As Variant)
    Dim lo As ListObject
    Dim rw As ListRow
    Dim i As Long

    Set lo = EnsureOverridesTable()

    ' Check for existing entry (upsert)
    If Not lo.DataBodyRange Is Nothing Then
        For i = 1 To lo.ListRows.Count
            Set rw = lo.ListRows(i)
            If CStr(rw.Range(1, 1).Value) = material And _
               CStr(rw.Range(1, 2).Value) = plant And _
               CStr(rw.Range(1, 3).Value) = productNumber And _
               CStr(rw.Range(1, 4).Value) = colHeader Then
                ' Update existing
                rw.Range(1, 5).Value = overrideValue
                Exit Sub
            End If
        Next i
    End If

    ' Insert new row
    Set rw = lo.ListRows.Add
    rw.Range(1, 1).Value = material
    rw.Range(1, 2).Value = plant
    rw.Range(1, 3).Value = productNumber
    rw.Range(1, 4).Value = colHeader
    rw.Range(1, 5).Value = overrideValue
End Sub

'''
' Removes a manual override for a specific Material + Plant + ProductNumber + Column.
' Called when the user clears a protected cell.
' @param {String} material The material number.
' @param {String} plant The plant code.
' @param {String} productNumber The product number (row-level identifier).
' @param {String} colHeader The column header name.
'''
Public Sub RemoveOverride(ByVal material As String, ByVal plant As String, _
                          ByVal productNumber As String, ByVal colHeader As String)
    Dim lo As ListObject
    Dim rw As ListRow
    Dim i As Long

    Set lo = EnsureOverridesTable()

    If lo.DataBodyRange Is Nothing Then Exit Sub

    ' Search from bottom to top to safely delete
    For i = lo.ListRows.Count To 1 Step -1
        Set rw = lo.ListRows(i)
        If CStr(rw.Range(1, 1).Value) = material And _
           CStr(rw.Range(1, 2).Value) = plant And _
           CStr(rw.Range(1, 3).Value) = productNumber And _
           CStr(rw.Range(1, 4).Value) = colHeader Then
            rw.Delete
            Exit Sub
        End If
    Next i
End Sub

'''
' Loads all overrides into a Dictionary for fast O(1) lookups.
' Key format: "material|plant|productNumber|columnHeader"
' Value: the override value
' @return {Object} A Scripting.Dictionary with all overrides.
'''
Public Function LoadOverridesDict() As Object
    Dim dict As Object
    Dim lo As ListObject
    Dim rw As ListRow
    Dim dictKey As String

    Set dict = CreateObject("Scripting.Dictionary")
    Set lo = EnsureOverridesTable()

    If Not lo.DataBodyRange Is Nothing Then
        Dim i As Long
        For i = 1 To lo.ListRows.Count
            Set rw = lo.ListRows(i)
            dictKey = CStr(rw.Range(1, 1).Value) & "|" & _
                      CStr(rw.Range(1, 2).Value) & "|" & _
                      CStr(rw.Range(1, 3).Value) & "|" & _
                      CStr(rw.Range(1, 4).Value)
            If Not dict.exists(dictKey) Then
                dict.Add dictKey, rw.Range(1, 5).Value
            End If
        Next i
    End If

    Set LoadOverridesDict = dict
End Function

'''
' Deletes all data rows from the ManualOverrides table.
' Called by DeleteAllProducts (ClearAllSheets).
'''
Public Sub ClearAllOverrides()
    Dim lo As ListObject
    Set lo = EnsureOverridesTable()

    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.Delete
    End If
End Sub
