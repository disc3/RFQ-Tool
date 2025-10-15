VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal target As Range)
    ' Check if the changed cell is either the Plant dropdown (D12)

    If Not Intersect(target, Me.Range("E12")) Is Nothing Then
        ' Synchronize with "1. BOM Definition" sheet C9
        application.EnableEvents = False
        ThisWorkbook.Sheets("1. BOM Definition").Range("C9  ").value = target.value
        ThisWorkbook.Sheets("4.1 Customer service offer").Range("E12  ").value = target.value
        application.EnableEvents = True
        Call ToggleMovingPriceSheetsBasedOnPlant
    End If
End Sub

