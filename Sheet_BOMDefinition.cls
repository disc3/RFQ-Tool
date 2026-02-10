VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet_BOMDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'================================================================================
' SHEET MODULE: "1. BOM Definition"
'
' INSTALLATION: Copy this code into the sheet module of "1. BOM Definition".
'   In the VBA editor, double-click the sheet "1. BOM Definition" in the
'   Project Explorer and paste the Worksheet_Change sub below.
'
' PURPOSE: Automatically tracks manual edits to protected columns (Price,
'   Price Unit, MOQ, etc.) in the ManualOverrides changelog table.
'   This ensures that "Update Components" does not overwrite purchasing edits.
'
' IMPORTANT: Only genuine manual edits are tracked. All macro/programmatic
'   writes set ManualOverrides.SuppressChangeTracking = True before writing,
'   which causes this event to exit immediately.
'================================================================================
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    ' --- Exit if programmatic writes are in progress ---
    If ManualOverrides.SuppressChangeTracking Then Exit Sub

    ' --- Setup: get the BOMDefinition table ---
    Dim loBom As ListObject
    On Error Resume Next
    Set loBom = Me.ListObjects("BOMDefinition")
    On Error GoTo 0
    If loBom Is Nothing Then Exit Sub

    ' --- Exit if table has no data ---
    If loBom.DataBodyRange Is Nothing Then Exit Sub

    ' --- Check if Target intersects the table's data area ---
    Dim isect As Range
    Set isect = Application.Intersect(Target, loBom.DataBodyRange)
    If isect Is Nothing Then Exit Sub

    ' --- Process each changed cell ---
    Dim cell As Range
    Dim colName As String
    Dim material As String, plant As String, productNumber As String
    Dim matColIdx As Long, plantColIdx As Long, prodColIdx As Long
    Dim rowNum As Long, colIdx As Long, totalRows As Long
    Dim r As Long
    Dim otherMat As String, otherPlant As String, otherProd As String
    Dim otherCell As Range
    Dim numFmt As String
    Dim cellCleared As Boolean

    On Error Resume Next
    matColIdx = loBom.ListColumns("Material").Index
    plantColIdx = loBom.ListColumns("Plant").Index
    prodColIdx = loBom.ListColumns("Product Number").Index
    On Error GoTo 0

    If matColIdx = 0 Or plantColIdx = 0 Or prodColIdx = 0 Then Exit Sub

    ' Temporarily suppress events to avoid recursion
    Application.EnableEvents = False

    totalRows = loBom.ListRows.Count

    For Each cell In isect.Cells
        ' Determine which column this cell is in
        colIdx = cell.Column - loBom.DataBodyRange.Column + 1
        colName = loBom.HeaderRowRange.Cells(1, colIdx).Value

        ' Only track protected columns
        If ManualOverrides.IsProtectedColumn(colName) Then
            ' Get the row number relative to the table
            rowNum = cell.Row - loBom.DataBodyRange.Row + 1

            ' Read Material, Plant and Product Number from the same row
            material = CStr(loBom.DataBodyRange.Cells(rowNum, matColIdx).Value)
            plant = CStr(loBom.DataBodyRange.Cells(rowNum, plantColIdx).Value)
            productNumber = CStr(loBom.DataBodyRange.Cells(rowNum, prodColIdx).Value)

            ' Get the column's NumberFormat for consistent formatting
            numFmt = cell.numberFormat

            ' Skip if Material is empty (incomplete row)
            If Len(Trim$(material)) > 0 Then
                cellCleared = IsEmpty(cell.Value) Or Len(Trim$(CStr(cell.Value))) = 0

                If cellCleared Then
                    ' Cell was cleared -> remove the override
                    ManualOverrides.RemoveOverride material, plant, productNumber, colName
                Else
                    ' Cell was edited -> record the override with format
                    ManualOverrides.RecordOverride material, plant, productNumber, colName, cell.Value, numFmt
                End If

                ' --- Propagate to all other BOM rows with the same Material+Plant ---
                For r = 1 To totalRows
                    If r <> rowNum Then
                        otherMat = CStr(loBom.DataBodyRange.Cells(r, matColIdx).Value)
                        otherPlant = CStr(loBom.DataBodyRange.Cells(r, plantColIdx).Value)

                        If otherMat = material And otherPlant = plant Then
                            Set otherCell = loBom.DataBodyRange.Cells(r, colIdx)
                            If Not otherCell.HasFormula Then
                                otherProd = CStr(loBom.DataBodyRange.Cells(r, prodColIdx).Value)

                                If cellCleared Then
                                    otherCell.ClearContents
                                    ManualOverrides.RemoveOverride otherMat, otherPlant, otherProd, colName
                                Else
                                    otherCell.Value = cell.Value
                                    otherCell.numberFormat = numFmt
                                    ManualOverrides.RecordOverride otherMat, otherPlant, otherProd, colName, cell.Value, numFmt
                                End If
                            End If
                        End If
                    End If
                Next r
            End If
        End If
    Next cell

    Application.EnableEvents = True
End Sub
