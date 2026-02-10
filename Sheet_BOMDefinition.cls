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
' SECTIONS:
'   1. Selected Plant (C9) synchronization to other sheets
'   2. API Triggers (C10:C12) for PowerAutomate
'   3. Manual Override tracking + propagation for protected columns
'================================================================================
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    ' ---------------------------------------------------------
    ' 1. Handle "Selected Plant" (C9) synchronization
    ' ---------------------------------------------------------
    If Not Intersect(Target, Me.Range("C9")) Is Nothing Then
        Application.EnableEvents = False
        On Error Resume Next
        ThisWorkbook.Sheets("4. Sales Calculation (Internal)").Range("E12").Value = Target.Text
        ThisWorkbook.Sheets("4.1 Customer service offer").Range("E12").Value = Target.Text
        On Error GoTo 0
        Application.EnableEvents = True

        Call ToggleMovingPriceSheetsBasedOnPlant
    End If

    ' ---------------------------------------------------------
    ' 2. Handle API Triggers (C10:C12)
    ' ---------------------------------------------------------
    Dim KeyCells As Range
    Set KeyCells = Me.Range("C10:C12")

    If Not Application.Intersect(KeyCells, Target) Is Nothing Then
        On Error GoTo ErrorHandler

        ' --- Logic A: C12 (Extended Workbench Checkbox) Changed ---
        If Not Intersect(Target, Me.Range("C12")) Is Nothing Then
            Dim isWorkbench As Boolean
            If IsError(Me.Range("C12").Value) Then Exit Sub
            If IsEmpty(Me.Range("C12").Value) Then
                isWorkbench = False
            Else
                isWorkbench = CBool(Me.Range("C12").Value)
            End If

            Call PowerAutomateAPI.SendExtendedWorkbenchUpdate(isWorkbench)
        End If

        ' --- Logic B: C10 or C11 Changed (General Updates) ---
        If Not Intersect(Target, Me.Range("C10:C11")) Is Nothing Then
            If Application.WorksheetFunction.CountA(Target) > 0 Then
                Call PowerAutomateAPI.SendGeneralStatusUpdate
            End If
        End If
    End If

    ' ---------------------------------------------------------
    ' 3. Manual Override Tracking + Propagation
    ' ---------------------------------------------------------
    ' Exit if programmatic writes are in progress
    If ManualOverrides.SuppressChangeTracking Then Exit Sub

    Dim loBom As ListObject
    On Error Resume Next
    Set loBom = Me.ListObjects("BOMDefinition")
    On Error GoTo 0
    If loBom Is Nothing Then Exit Sub

    If loBom.DataBodyRange Is Nothing Then Exit Sub

    Dim isect As Range
    Set isect = Application.Intersect(Target, loBom.DataBodyRange)
    If isect Is Nothing Then Exit Sub

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
        colIdx = cell.Column - loBom.DataBodyRange.Column + 1
        colName = loBom.HeaderRowRange.Cells(1, colIdx).Value

        If ManualOverrides.IsProtectedColumn(colName) Then
            rowNum = cell.Row - loBom.DataBodyRange.Row + 1

            material = CStr(loBom.DataBodyRange.Cells(rowNum, matColIdx).Value)
            plant = CStr(loBom.DataBodyRange.Cells(rowNum, plantColIdx).Value)
            productNumber = CStr(loBom.DataBodyRange.Cells(rowNum, prodColIdx).Value)

            numFmt = cell.NumberFormat

            If Len(Trim$(material)) > 0 Then
                cellCleared = IsEmpty(cell.Value) Or Len(Trim$(CStr(cell.Value))) = 0

                If cellCleared Then
                    ManualOverrides.RemoveOverride material, plant, productNumber, colName
                Else
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
                                    otherCell.NumberFormat = numFmt
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

    Exit Sub

ErrorHandler:
    MsgBox "An error occurred in the Worksheet_Change event: " & Err.Description, vbCritical
    Application.EnableEvents = True
End Sub
