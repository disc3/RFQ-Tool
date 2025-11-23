Attribute VB_Name = "SAP_BOM_Uploader_copy"
Sub ExportBOMToSAPTemplate()
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim tbl As ListObject
    Dim row As ListRow
    Dim outRow As Long
    Dim plant As String
    Dim currentProd As String, prevProd As String
    Dim itemNumber As Long
    Dim baseQtyRow As Long
    Dim matCol As Long, descCol As Long, qtyCol As Long, unitCol As Long
    Dim decimalSeparator As String

    Set wsSource = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsTarget = ThisWorkbook.Sheets("Template_BOM_Connect")
    Set tbl = wsSource.ListObjects("BOMDefinition")

    wsTarget.Visible = xlSheetVisible

    ' Source column indexes
    matCol = tbl.ListColumns("Material").Index
    descCol = tbl.ListColumns("Material Description").Index
    qtyCol = tbl.ListColumns("Quantity").Index
    unitCol = tbl.ListColumns("Base unit of component").Index

    ' Clear target sheet except headers
    wsTarget.Range("A3:T10000").ClearContents
    outRow = 3
    prevProd = ""
    itemNumber = 0

    ' Get the correct decimal separator for the system
    decimalSeparator = application.International(xlDecimalSeparator)

    For Each row In tbl.ListRows
        currentProd = row.Range(tbl.ListColumns("ERP Part Number").Index).Value
        plant = ThisWorkbook.Sheets("1. BOM Definition").Range("C9").Value
        
        ' Insert header row if new product
        If currentProd <> prevProd Then
            With wsTarget
                .Cells(outRow, 1).Value = "H"
                .Cells(outRow, 2).Value = currentProd
                .Cells(outRow, 3).Value = plant
                .Cells(outRow, 4).Value = Date
                .Cells(outRow, 5).Value = "1"
                .Cells(outRow, 6).Value = "Pc"
            End With
            baseQtyRow = outRow
            outRow = outRow + 1
            prevProd = currentProd
            itemNumber = 10
        End If

        ' Insert item row
        With wsTarget
            .Cells(outRow, 1).Value = "I"
            .Cells(outRow, 8).Value = itemNumber
            ' Force Material as string to preserve leading zeros
            .Cells(outRow, 9).Value = "'" & CStr(row.Range(matCol).Value)
            .Cells(outRow, 10).Value = row.Range(descCol).Value
            
            Dim qtyValue As String
            qtyValue = CStr(row.Range(qtyCol).Value)
            
            If decimalSeparator = "," Then
                qtyValue = Replace(qtyValue, ".", ",")
            Else
                qtyValue = Replace(qtyValue, ",", ".")
            End If
            
            .Cells(outRow, 11).FormulaLocal = "=$E$" & baseQtyRow & "*" & qtyValue


            .Cells(outRow, 12).Value = row.Range(unitCol).Value
            .Cells(outRow, 18).Value = "X"

            ' Formula for individual length
            Dim qtyCell As String, cableCell As String, baseQtyCell As String
            qtyCell = .Cells(outRow, 11).Address(False, False)
            cableCell = .Cells(outRow, 14).Address(False, False)
            baseQtyCell = "$E$" & baseQtyRow
            .Cells(outRow, 15).Formula = "=IF(" & cableCell & "=" & Chr(34) & "YES" & Chr(34) & "," & qtyCell & "/" & baseQtyCell & ","""")"
        End With

        itemNumber = itemNumber + 10
        outRow = outRow + 1
    Next row

    wsTarget.Activate
    MsgBox "BOM export complete.", vbInformation
End Sub

