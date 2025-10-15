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
    selectedPlant = Trim(CStr(wsMain.Range("C9").value))

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

