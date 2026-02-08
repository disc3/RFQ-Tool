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
    Dim material As String
    Dim plant As String
    Dim matColIdx As Long, plantColIdx As Long

    On Error Resume Next
    matColIdx = loBom.ListColumns("Material").Index
    plantColIdx = loBom.ListColumns("Plant").Index
    On Error GoTo 0

    If matColIdx = 0 Or plantColIdx = 0 Then Exit Sub

    ' Temporarily suppress events to avoid recursion
    Application.EnableEvents = False

    Dim rowNum As Long
    For Each cell In isect.Cells
        ' Determine which column this cell is in
        colName = loBom.HeaderRowRange.Cells(1, cell.Column - loBom.DataBodyRange.Column + 1).Value

        ' Only track protected columns
        If ManualOverrides.IsProtectedColumn(colName) Then
            ' Get the row number relative to the table
            rowNum = cell.Row - loBom.DataBodyRange.Row + 1

            ' Read Material and Plant from the same row
            material = CStr(loBom.DataBodyRange.Cells(rowNum, matColIdx).Value)
            plant = CStr(loBom.DataBodyRange.Cells(rowNum, plantColIdx).Value)

            ' Skip if Material is empty (incomplete row)
            If Len(Trim$(material)) > 0 Then
                If IsEmpty(cell.Value) Or Len(Trim$(CStr(cell.Value))) = 0 Then
                    ' Cell was cleared -> remove the override
                    ManualOverrides.RemoveOverride material, plant, colName
                Else
                    ' Cell was edited -> record the override
                    ManualOverrides.RecordOverride material, plant, colName, cell.Value
                End If
            End If
        End If
    Next cell

    Application.EnableEvents = True
End Sub
