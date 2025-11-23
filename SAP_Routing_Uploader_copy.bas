Attribute VB_Name = "SAP_Routing_Uploader_copy"
Sub ExportRoutingToSAPTemplate()
    Call GenerateERPRoutine
    Dim wsSource As Worksheet, wsOut As Worksheet
    Dim tbl As ListObject
    Dim outRow As Long
    Dim currentProduct As String, lastProduct As String, plant As String
    Dim i As Long, opNum As Long

    Set wsSource = ThisWorkbook.Sheets("6. Routine uploaders")
    Set wsOut = ThisWorkbook.Sheets("Template_Routing_Connect")
    Set tbl = wsSource.ListObjects("ERPRouting")

    ' Clear existing content except headers
    wsOut.Range("A4:Z10000").EntireRow.ClearContents
    wsOut.Visible = xlSheetVisible
    
    outRow = 4
    lastProduct = ""
    opNum = 10

    For i = 1 To tbl.DataBodyRange.Rows.Count
        currentProduct = tbl.DataBodyRange.Cells(i, 1).Value ' "Product"
        plant = ThisWorkbook.Sheets("1. BOM Definition").Range("C9").Value

        ' Get setup and machine times
        Dim setupTime As Variant, machineTime As Variant
        setupTime = tbl.DataBodyRange.Cells(i, 12).Value ' "Setup"
        machineTime = tbl.DataBodyRange.Cells(i, 15).Value ' "Machine"

        ' Skip if both are zero or empty
        If (Nz(setupTime) = 0 And Nz(machineTime) = 0) Then
            GoTo SkipRow
        End If

        ' Write new header if material changes
        If currentProduct <> lastProduct Then
            With wsOut
                .Cells(outRow, 1).Value = "H"
                .Cells(outRow, 2).Value = currentProduct
                .Cells(outRow, 3).Value = plant
                .Cells(outRow, 4).Value = 1
                .Cells(outRow, 5).Value = 4
                .Cells(outRow, 6).Value = 142
                .Cells(outRow, 7).Value = Date ' today's date
            End With
            outRow = outRow + 1
            opNum = 10
            lastProduct = currentProduct
        End If

        ' Write operation
        With wsOut
            .Cells(outRow, 1).Value = "O"
            .Cells(outRow, 8).Value = opNum
            .Cells(outRow, 9).Value = "" ' blank col
            .Cells(outRow, 10).Value = tbl.DataBodyRange.Cells(i, 4).Value  ' Work center
            .Cells(outRow, 11).Value = tbl.DataBodyRange.Cells(i, 5).Value  ' Plant
            .Cells(outRow, 12).Value = tbl.DataBodyRange.Cells(i, 6).Value  ' Control key
            .Cells(outRow, 13).Value = "" ' blank col
            .Cells(outRow, 14).Value = tbl.DataBodyRange.Cells(i, 8).Value  ' Description
            .Cells(outRow, 15).Value = tbl.DataBodyRange.Cells(i, 10).Value ' Base qty
            .Cells(outRow, 16).Value = tbl.DataBodyRange.Cells(i, 11).Value ' Unit
            .Cells(outRow, 17).Value = setupTime
            .Cells(outRow, 18).Value = tbl.DataBodyRange.Cells(i, 13).Value ' Unit2
            .Cells(outRow, 19).Value = machineTime
            .Cells(outRow, 20).Value = tbl.DataBodyRange.Cells(i, 16).Value ' Unit3
            .Cells(outRow, 21).Value = "" ' Personal time
            .Cells(outRow, 22).Value = tbl.DataBodyRange.Cells(i, 16).Value ' Re-use machine time unit
        End With

        opNum = opNum + 10
        outRow = outRow + 1
SkipRow:
    Next i

    'MsgBox "Routing exported to SAP template successfully!", vbInformation
    wsOut.Activate
End Sub

Private Function Nz(val As Variant) As Double
    If IsEmpty(val) Or IsNull(val) Or val = "" Then
        Nz = 0
    Else
        Nz = val
    End If
End Function



