Attribute VB_Name = "Modul1"
Sub ShowSplitUserForm()
    UserForm1.Show
End Sub
' TRUE if the cell has a manual fill (Interior) — not CF
Public Function HasManualFill(c As Range) As Boolean
    HasManualFill = (c.Interior.ColorIndex <> xlColorIndexNone)
End Function


Public Sub ToggleMovingPriceSheetsBasedOnPlant()
    Dim selectedPlant As String
    Dim wsMain As Worksheet

    Set wsMain = ThisWorkbook.Sheets("1. BOM Definition")
    selectedPlant = Trim(CStr(wsMain.Range("C9").Value))

    On Error Resume Next
    With ThisWorkbook
        If selectedPlant = "PL10" Then
            .Sheets("4.2 Product Moving Price (PL)").Visible = xlSheetVisible
            .Sheets("4.3 HALB Moving Price (PL)").Visible = xlSheetVisible
        Else
            .Sheets("4.2 Product Moving Price (PL)").Visible = xlSheetVeryHidden
            .Sheets("4.3 HALB Moving Price (PL)").Visible = xlSheetVeryHidden
        End If
    End With
    On Error GoTo 0
End Sub

Sub Debug_CheckColumnHeaders()
    Dim tbl As ListObject
    Set tbl = ThisWorkbook.Sheets("1. BOM Definition").ListObjects("BOMDefinition")
    
    Dim col As ListColumn
    Debug.Print "--- Checking Headers in 'BOMDefinition' ---"
    
    For Each col In tbl.ListColumns
        If IsError(col.Range.Cells(1, 1).Value) Then
            Debug.Print "COLUMN ERROR: Header in cell " & col.Range.Cells(1, 1).Address & " is an error value."
        Else
            Debug.Print "OK: Column " & col.Index & " -> '" & col.name & "'"
        End If
    Next col
    
    Debug.Print "--- Check Complete ---"
End Sub
