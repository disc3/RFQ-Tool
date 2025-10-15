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
        currentProduct = tbl.DataBodyRange.Cells(i, 1).value ' "Product"
        plant = ThisWorkbook.Sheets("1. BOM Definition").Range("C9").value

        ' Get setup and machine times
        Dim setupTime As Variant, machineTime As Variant
        setupTime = tbl.DataBodyRange.Cells(i, 12).value ' "Setup"
        machineTime = tbl.DataBodyRange.Cells(i, 15).value ' "Machine"

        ' Skip if both are zero or empty
        If (Nz(setupTime) = 0 And Nz(machineTime) = 0) Then
            GoTo SkipRow
        End If

        ' Write new header if material changes
        If currentProduct <> lastProduct Then
            With wsOut
                .Cells(outRow, 1).value = "H"
                .Cells(outRow, 2).value = currentProduct
                .Cells(outRow, 3).value = plant
                .Cells(outRow, 4).value = 1
                .Cells(outRow, 5).value = 4
                .Cells(outRow, 6).value = 142
                .Cells(outRow, 7).value = Date ' today's date
            End With
            outRow = outRow + 1
            opNum = 10
            lastProduct = currentProduct
        End If

        ' Write operation
        With wsOut
            .Cells(outRow, 1).value = "O"
            .Cells(outRow, 8).value = opNum
            .Cells(outRow, 9).value = "" ' blank col
            .Cells(outRow, 10).value = tbl.DataBodyRange.Cells(i, 4).value  ' Work center
            .Cells(outRow, 11).value = tbl.DataBodyRange.Cells(i, 5).value  ' Plant
            .Cells(outRow, 12).value = tbl.DataBodyRange.Cells(i, 6).value  ' Control key
            .Cells(outRow, 13).value = "" ' blank col
            .Cells(outRow, 14).value = tbl.DataBodyRange.Cells(i, 8).value  ' Description
            .Cells(outRow, 15).value = tbl.DataBodyRange.Cells(i, 10).value ' Base qty
            .Cells(outRow, 16).value = tbl.DataBodyRange.Cells(i, 11).value ' Unit
            .Cells(outRow, 17).value = setupTime
            .Cells(outRow, 18).value = tbl.DataBodyRange.Cells(i, 13).value ' Unit2
            .Cells(outRow, 19).value = machineTime
            .Cells(outRow, 20).value = tbl.DataBodyRange.Cells(i, 16).value ' Unit3
            .Cells(outRow, 21).value = "" ' Personal time
            .Cells(outRow, 22).value = tbl.DataBodyRange.Cells(i, 16).value ' Re-use machine time unit
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



