Attribute VB_Name = "PurchasingInfo"
Sub CopyUnpricedComponentsToPurchasingInput()
    Dim wsBom As Worksheet, wsPI As Worksheet
    Dim tblBOM As ListObject, tblPI As ListObject
    Dim rowBOM As ListRow, rowPI As ListRow
    Dim colBOM As Range, colPI As Range
    Dim i As Long

    Set wsBom = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsPI = ThisWorkbook.Sheets("1.1. Purchasing Input")
    Set tblBOM = wsBom.ListObjects("BOMDefinition")
    Set tblPI = wsPI.ListObjects("PurchasingInput")

    application.ScreenUpdating = False
    For Each rowBOM In tblBOM.ListRows
        ' Check if price is empty or zero (use "Price" column, the editable source)
        If Trim(rowBOM.Range.Columns(tblBOM.ListColumns("Price").Index).Value) = "" _
           Or rowBOM.Range.Columns(tblBOM.ListColumns("Price").Index).Value = 0 Then

            Set rowPI = tblPI.ListRows.Add
            For i = 1 To tblBOM.ListColumns.Count
                On Error Resume Next
                Set colBOM = tblBOM.ListColumns(i).Range.Cells(1, 1)
                Set colPI = tblPI.ListColumns(colBOM.Value).Range.Cells(1, 1)
                If Not colPI Is Nothing Then
                    rowPI.Range(1, colPI.column - tblPI.Range.Cells(1, 1).column + 1).Value = _
                        rowBOM.Range(1, i).Value
                End If
                Set colPI = Nothing
                On Error GoTo 0
            Next i
        End If
    Next rowBOM
    application.ScreenUpdating = True

    MsgBox "Unpriced components copied to PurchasingInput.", vbInformation
End Sub


Sub CopyAllComponentsToPurchasingInput()
    Dim wsBom As Worksheet, wsPI As Worksheet
    Dim tblBOM As ListObject, tblPI As ListObject
    Dim rowBOM As ListRow, rowPI As ListRow
    Dim colBOM As Range, colPI As Range
    Dim i As Long

    Set wsBom = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsPI = ThisWorkbook.Sheets("1.1. Purchasing Input")
    Set tblBOM = wsBom.ListObjects("BOMDefinition")
    Set tblPI = wsPI.ListObjects("PurchasingInput")

    application.ScreenUpdating = False
    For Each rowBOM In tblBOM.ListRows
        Set rowPI = tblPI.ListRows.Add
        For i = 1 To tblBOM.ListColumns.Count
            On Error Resume Next
            Set colBOM = tblBOM.ListColumns(i).Range.Cells(1, 1)
            Set colPI = tblPI.ListColumns(colBOM.Value).Range.Cells(1, 1)
            If Not colPI Is Nothing Then
                rowPI.Range(1, colPI.column - tblPI.Range.Cells(1, 1).column + 1).Value = _
                    rowBOM.Range(1, i).Value
            End If
            Set colPI = Nothing
            On Error GoTo 0
        Next i
    Next rowBOM
    application.ScreenUpdating = True

    MsgBox "All components copied to PurchasingInput.", vbInformation
End Sub
Sub UpdateBOMDefinitionFromPurchasingInput()
    Dim wsBom As Worksheet, wsPI As Worksheet
    Dim tblBOM As ListObject, tblPI As ListObject
    Dim rowBOM As ListRow, rowPI As ListRow
    Dim colMap As Object
    Dim i As Long, bomColIndex As Long
    Dim key As String
    Dim dictPI As Object
    Set dictPI = CreateObject("Scripting.Dictionary")
    Set colMap = CreateObject("Scripting.Dictionary")

    Set wsBom = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsPI = ThisWorkbook.Sheets("1.1. Purchasing Input")
    Set tblBOM = wsBom.ListObjects("BOMDefinition")
    Set tblPI = wsPI.ListObjects("PurchasingInput")

    ' Normalize BOM column names and build a mapping
    Dim normName As String
    For i = 1 To tblBOM.ListColumns.Count
        normName = NormalizeHeader(tblBOM.ListColumns(i).name)
        colMap(normName) = i
    Next i

    ' Build dictionary of BOM rows by Product Number + Material
    For Each rowBOM In tblBOM.ListRows
        key = Trim(rowBOM.Range(1, tblBOM.ListColumns("Product Number").Index).Value) & "|" & _
              Trim(rowBOM.Range(1, tblBOM.ListColumns("Material").Index).Value)
        dictPI(key) = rowBOM.Index
    Next rowBOM

    application.ScreenUpdating = False
    For Each rowPI In tblPI.ListRows
        key = Trim(rowPI.Range(1, tblPI.ListColumns("Product Number").Index).Value) & "|" & _
              Trim(rowPI.Range(1, tblPI.ListColumns("Material").Index).Value)

        If dictPI.exists(key) Then
            Set rowBOM = tblBOM.ListRows(dictPI(key))

            ' Loop through PI columns and update only matching, non-formula BOM cells
            For i = 1 To tblPI.ListColumns.Count
                normName = NormalizeHeader(tblPI.ListColumns(i).name)
                If colMap.exists(normName) Then
                    bomColIndex = colMap(normName)
                    If Not rowBOM.Range(1, bomColIndex).HasFormula Then
                        Dim rawValue As Variant
                        rawValue = rowPI.Range(1, i).Value
                    
                        If IsNumeric(rawValue) Then
                            ' Convert to string, normalize the decimal separator
                            Dim rawStr As String
                            rawStr = CStr(rawValue)
                    
                            Dim decSep As String
                            decSep = application.International(xlDecimalSeparator)
                    
                            ' Fix wrong separator
                            If decSep = "," And InStr(rawStr, ".") > 0 Then
                                rawStr = Replace(rawStr, ".", ",")
                            ElseIf decSep = "." And InStr(rawStr, ",") > 0 Then
                                rawStr = Replace(rawStr, ",", ".")
                            End If
                    
                            ' Convert back to numeric with proper separator
                            If Trim(rawStr) <> "" Then
                                On Error Resume Next
                                Dim parsedValue As Double
                                parsedValue = CDbl(rawStr)
                                If Err.Number = 0 Then
                                    rowBOM.Range(1, bomColIndex).Value = parsedValue
                                Else
                                    ' Fallback if conversion fails
                                    rowBOM.Range(1, bomColIndex).Value = rawStr
                                End If
                                On Error GoTo 0
                            Else
                                rowBOM.Range(1, bomColIndex).Value = ""
                            End If

                        Else
                            rowBOM.Range(1, bomColIndex).Value = rawValue
                        End If
                    End If

                End If
            Next i
        End If
    Next rowPI
    application.ScreenUpdating = True

    MsgBox "BOMDefinition safely updated from PurchasingInput (columns matched by name).", vbInformation
End Sub

Private Function NormalizeHeader(header As String) As String
    ' Remove spaces, brackets, lowercase everything
    NormalizeHeader = LCase(Replace(Replace(Replace(header, " ", ""), "[", ""), "]", ""))
End Function






