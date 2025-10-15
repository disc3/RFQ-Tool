Attribute VB_Name = "ClearTables"
Public Sub ClearSelectedRoutinesTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim confirmClear As VbMsgBoxResult
    
    ' Prompt the user for confirmation before clearing
    confirmClear = MsgBox("Are you sure you want to clear all rows in the SelectedRoutines table?", vbYesNo + vbQuestion, "Confirm Clear")
    
    ' Exit if the user selects "No"
    If confirmClear = vbNo Then Exit Sub
    
    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets("2. Routines") ' Replace with the actual sheet name
    Set tbl = ws.ListObjects("SelectedRoutines") ' Replace with the actual table name
    
    ' Delete all rows in the DataBodyRange if the table has rows
    If Not tbl.DataBodyRange Is Nothing Then
        With tbl
            On Error Resume Next
            .DataBodyRange.offset(1).Resize(.DataBodyRange.Rows.Count - 1, .DataBodyRange.Columns.Count).Rows.Delete
            .DataBodyRange.SpecialCells(xlCellTypeConstants).ClearContents
            On Error GoTo 0
        End With
        
        MsgBox "All rows deleted from SelectedRoutines table", vbInformation
    Else
        MsgBox "The SelectedRoutines table is already empty.", vbExclamation
    End If
End Sub
Public Sub ClearSelectedComponentsTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim confirmClear As VbMsgBoxResult
    
    ' Prompt the user for confirmation before clearing
    confirmClear = MsgBox("Are you sure you want to clear data from the Selected Components table, keeping only formulas in the first row?", vbYesNo + vbQuestion, "Confirm Clear")
    
    ' Exit if the user selects "No"
    If confirmClear = vbNo Then Exit Sub
    
    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")
     
    ' Delete all rows except the first row in the table's DataBodyRange
    If tbl.DataBodyRange.Rows.Count > 1 Then
        tbl.DataBodyRange.offset(1, 0).Resize(tbl.DataBodyRange.Rows.Count - 1).Rows.Delete
    End If
    
    ' Loop through each column in the table, clearing contents but keeping only formulas in the first row
    For Each col In tbl.ListColumns
        For Each cell In col.DataBodyRange
            ' Clear cell content if it's static (not a formula), but leave formulas intact
            If Not cell.HasFormula And cell.row = tbl.DataBodyRange.Rows(1).row Then
                cell.ClearContents
            End If
        Next cell
    Next col
   
    MsgBox "Data cleared from SelectedComponents table, keeping formulas only in the first row", vbInformation
End Sub

Sub ClearProjectDataColumns()
    Dim wsProjectData As Worksheet
    Dim tblProjectData As ListObject
    Dim columnHeaders As Variant
    Dim header As Variant
    Dim columnIndex As Long

    ' Define the worksheet and table
    Set wsProjectData = ThisWorkbook.Sheets("0. ProjectData")
    On Error Resume Next
    Set tblProjectData = wsProjectData.ListObjects("ProjectData")
    On Error GoTo 0

    ' Ensure the table exists
    If tblProjectData Is Nothing Then
        MsgBox "The ProjectData table does not exist.", vbExclamation
        Exit Sub
    End If

    ' Define the column headers to clear
    columnHeaders = Array("Status", "RFQ Initialized", "BOM & Routing added", "Components validated by Purchasing", "Full RFQ Validated")

    ' Loop through each header and clear the corresponding column
    For Each header In columnHeaders
        On Error Resume Next
        columnIndex = tblProjectData.ListColumns(header).Index
        If columnIndex > 0 Then
            tblProjectData.ListColumns(columnIndex).DataBodyRange.ClearContents
        End If
        On Error GoTo 0
    Next header

    MsgBox "The specified columns in the ProjectData table have been cleared.", vbInformation
End Sub
Public Sub ClearMassUploadTable()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim confirmClear As VbMsgBoxResult
    
    ' Prompt the user for confirmation before clearing
    confirmClear = MsgBox("Are you sure you want to clear all rows in the Mass Upload table?", vbYesNo + vbQuestion, "Confirm Clear")
    
    ' Exit if the user selects "No"
    If confirmClear = vbNo Then Exit Sub
    
    ' Error handling if sheet or table does not exist
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("MassUpload") ' Replace with the actual sheet name
    Set tbl = ws.ListObjects("MassUploadTable") ' Replace with the actual table name
    On Error GoTo 0

    ' Check if table exists
    If tbl Is Nothing Then
        MsgBox "Error: The Mass Upload table does not exist.", vbExclamation
        Exit Sub
    End If
    
    ' Check if the table has data before deleting
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If

    ' Reset formatting - check if DataBodyRange exists first
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Interior.ColorIndex = xlNone ' Reset background color
        tbl.DataBodyRange.Font.ColorIndex = xlAutomatic ' Reset font color
    Else
        ' If the table is empty, reset formatting for the whole table range
        tbl.Range.Interior.ColorIndex = xlNone
        tbl.Range.Font.ColorIndex = xlAutomatic
    End If

    ' Remove conditional formatting
    tbl.Range.FormatConditions.Delete
    
    MsgBox "All rows deleted in the Mass Upload table.", vbInformation
End Sub

Public Sub ClearPurchasingInput()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim confirmClear As VbMsgBoxResult
    
    ' Prompt the user for confirmation before clearing
    ' confirmClear = MsgBox("Are you sure you want to clear all rows in the Purhcasing Input table?", vbYesNo + vbQuestion, "Confirm Clear")
    
    ' Exit if the user selects "No"
    If confirmClear = vbNo Then Exit Sub
    
    ' Error handling if sheet or table does not exist
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("1.1. Purchasing Input") ' Replace with the actual sheet name
    Set tbl = ws.ListObjects("PurchasingInput") ' Replace with the actual table name
    On Error GoTo 0

    ' Check if table exists
    If tbl Is Nothing Then
        MsgBox "Error: The PurchasingInput table does not exist.", vbExclamation
        Exit Sub
    End If
    
    

    ' Reset formatting - check if DataBodyRange exists first
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Interior.ColorIndex = xlNone ' Reset background color
        tbl.DataBodyRange.Font.ColorIndex = xlAutomatic ' Reset font color
    Else
        ' If the table is empty, reset formatting for the whole table range
        tbl.Range.Interior.ColorIndex = xlNone
        tbl.Range.Font.ColorIndex = xlAutomatic
    End If
    
    ' Check if the table has data before deleting
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.Delete
    End If
    ' Remove conditional formatting
    tbl.Range.FormatConditions.Delete
    
    MsgBox "All rows deleted in the Purchasing Input table.", vbInformation
End Sub

Public Sub ClearComponentDatabase(Optional isClosingExcel As Boolean)
    Dim ws, wsResults As Worksheet
    Dim tbl As ListObject
    
    On Error Resume Next
    ' Set the worksheet and table
    Set ws = ThisWorkbook.Sheets("Purchasing Info Records") ' Replace with the actual sheet name
    Set tbl = ws.ListObjects("LoadedData") ' Replace with the actual table name
    
    ' Delete all rows in the DataBodyRange if the table has rows
    If Not tbl.DataBodyRange Is Nothing Then
        With tbl
            .Range.AutoFilter
            If .ListRows.Count > 1 Then
                .DataBodyRange.offset(1).Resize(.DataBodyRange.Rows.Count - 1, .DataBodyRange.Columns.Count).Rows.Delete
                application.StatusBar = "Component Database has been cleared. Auto-save has been turned on again."
            Else
                application.StatusBar = "Component Database is not loaded. Auto-save is active."
            End If
            On Error Resume Next
            .DataBodyRange.SpecialCells(xlCellTypeConstants).ClearContents
            On Error GoTo 0
        End With
    End If
    On Error Resume Next
    
    FilterComponents.SetCloudAutoSave True
    ThisWorkbook.Save
    On Error GoTo 0
    
    If Not isClosingExcel Then
        ClearStatusBar
    End If
End Sub

Public Sub ClearStatusBar(Optional ClearStatusBar As Boolean)
    If ClearStatusBar Then
        application.StatusBar = False
    Else
        application.OnTime Now + TimeValue("00:00:10"), "'ClearStatusBar True'"
    End If
End Sub
