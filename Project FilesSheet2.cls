VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal target As Range)
    ' Check if the changed cell is F11 (Product selection)
    If Not Intersect(target, Me.Range("F11")) Is Nothing Then
        Dim ws As Worksheet
        Dim tbl As ListObject
        Dim foundRow As Range
        Dim ArtNumber As String

        ' Get the selected Product Number
        ArtNumber = Me.Range("F11").value

        ' Reference the Products sheet and FinalProductList table
        Set ws = ThisWorkbook.Sheets("Final Products")
        Set tbl = ws.ListObjects("FinalProductList")

        ' Ensure DataBodyRange exists before attempting to find
        If Not tbl.DataBodyRange Is Nothing Then
            ' Find the selected Product in the FinalProductList
            Set foundRow = tbl.ListColumns("Product Number").DataBodyRange.Find(ArtNumber, LookIn:=xlValues, LookAt:=xlWhole)

            ' If found, update F12 and F13
            If Not foundRow Is Nothing Then
                Me.Range("F13").value = foundRow.offset(0, 3).value ' Product Type
                Me.Range("F12").value = foundRow.offset(0, 2).value ' Product Description
            Else
                ' Optional: Clear F12 and F13 if product not found
                Me.Range("F13").ClearContents
                Me.Range("F12").ClearContents
            End If
        Else
            ' Optional: Clear F7 and F8 if FinalProductList is empty
            Me.Range("F12").ClearContents
            Me.Range("F13").ClearContents
        End If
    End If
    
     ' Check if the changed cell is C9 - Selected Plant
    If Not Intersect(target, Me.Range("C9")) Is Nothing Then
        ' Synchronize with "4. Sales Calculation (Internal)" sheet E12
        application.EnableEvents = False
        ThisWorkbook.Sheets("4. Sales Calculation (Internal)").Range("E12").value = target.value
        ThisWorkbook.Sheets("4.1 Customer service offer").Range("E12  ").value = target.value
        application.EnableEvents = True
        Call ToggleMovingPriceSheetsBasedOnPlant
    End If
End Sub

    

