VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal target As Range)
    ' Check if the changed cell is D6 (Product selection)
    If Not Intersect(target, Me.Range("D6")) Is Nothing Then
        Dim ws As Worksheet
        Dim tbl As ListObject
        Dim foundRow As Range
        Dim ArtNumber As String

        ' Get the selected Product Number
        ArtNumber = Me.Range("D6").value

        ' Reference the Products sheet and FinalProductList table
        Set ws = ThisWorkbook.Sheets("Final Products")
        Set tbl = ws.ListObjects("FinalProductList")

        ' Ensure DataBodyRange exists before attempting to find
        If Not tbl.DataBodyRange Is Nothing Then
            ' Find the selected Product in the FinalProductList
            Set foundRow = tbl.ListColumns("Product Number").DataBodyRange.Find(ArtNumber, LookIn:=xlValues, LookAt:=xlWhole)

            ' If found, update D8 and D9
            If Not foundRow Is Nothing Then
                Me.Range("D8").value = foundRow.offset(0, 1).value ' Product Description
                Me.Range("D9").value = foundRow.offset(0, 2).value ' Product Type
            Else
                ' Optional: Clear D8 and D9 if product not found
                Me.Range("D8").ClearContents
                Me.Range("D9").ClearContents
            End If
        Else
            ' Optional: Clear D8 and D9 if FinalProductList is empty
            Me.Range("D8").ClearContents
            Me.Range("D9").ClearContents
        End If
    End If
End Sub
