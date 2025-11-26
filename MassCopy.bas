Attribute VB_Name = "MassCopy"
Option Explicit

' =========================================================================================
'  1) PrepareForMassUpload - same as your original code
' =========================================================================================
Public Sub PrepareForMassUpload()
    Dim wsMain As Worksheet
    Dim qt As QueryTable
    Dim ws As Worksheet
    Dim searchTerm As String

    ' Get the search term from the TextBox (left blank here)
    searchTerm = ""

    ' Update the query parameter dynamically
    UpdateQueryParameterMass "SearchTerm", searchTerm

    ' Refresh the query to apply the new search term
    ' On Error Resume Next
    ' ThisWorkbook.Connections("GetExtendedMMD").Refresh

    ' Reference Product Specification sheet and clear B9
    ' Set wsMain = ThisWorkbook.Sheets("Product Specification")
    ' wsMain.Range("B9").Value = ""

    ' Refresh GetExtendedMMD query
    ' Reference the table named "GetExtendedMMD" on the "ExtendedMMD" sheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("ExtendedMMD")
    Set qt = ws.ListObjects("GetExtendedMMD").QueryTable
    On Error GoTo 0
    
    If qt Is Nothing Then
        MsgBox "Table 'GetExtendedMMD' not found on the 'ExtendedMMD' sheet.", vbCritical
        Exit Sub
    End If
    
    ' Refresh the data synchronously
    ' qt.Refresh BackgroundQuery:=False
End Sub


' =========================================================================================
'  2) UpdateQueryParameterMass - same as your original code
' =========================================================================================
Private Sub UpdateQueryParameterMass(ByVal paramName As String, ByVal paramValue As String)
    Dim queryName As String
    Dim queryText As String
    Dim updatedQueryText As String
    Dim searchPattern As String
    Dim replacePattern As String
    Dim connectionName As String
    Dim conn As WorkbookConnection

    ' Specify the name of the query to update
    queryName = "GetExtendedMMD"

    ' Get the current query formula
    On Error Resume Next
    queryText = ThisWorkbook.Queries(queryName).Formula
    On Error GoTo 0

    ' Define the search pattern to match any value assigned to SearchTerm
    searchPattern = "SearchTerm = "".*?""" ' Match "SearchTerm = " followed by any value
    replacePattern = "SearchTerm = """ & paramValue & """" ' New replacement value

    ' Use Regular Expressions to dynamically update SearchTerm
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = True
        .IgnoreCase = True
        .pattern = searchPattern
    End With

    ' Replace the SearchTerm in the query formula
    If regex.Test(queryText) Then
        updatedQueryText = regex.Replace(queryText, replacePattern)
    Else
        MsgBox "The query formula does not contain 'SearchTerm'.", vbCritical
        Exit Sub
    End If

    ' Update the query formula
    ThisWorkbook.Queries(queryName).Formula = updatedQueryText

    ' Find the connection associated with the query
    connectionName = ""
    For Each conn In ThisWorkbook.Connections
        If InStr(1, conn.name, queryName, vbTextCompare) > 0 Then
            connectionName = conn.name
            Exit For
        End If
    Next conn

    If connectionName = "" Then
        MsgBox "No connection found for query '" & queryName & "'.", vbCritical
        Exit Sub
    End If

    ' Refresh the connection without background
    With ThisWorkbook.Connections(connectionName)
        .OLEDBConnection.BackgroundQuery = False ' Disable background refresh
        .Refresh ' Synchronously refresh
    End With
End Sub
Public Sub ProcessMassUpload()
    Dim wsMassUpload As Worksheet
    Dim tblMassUpload As ListObject
    Dim rw As ListRow
    Dim materialNumber As String
    Dim quantity As String
    Dim productNumber As String
    Dim netPriceUnit As Double
    Dim matchingRows As Range
    Dim correctedQuantity As Double
    Dim decimalSeparator As String
    Dim notFoundMaterials As String
    Dim lastPN As String

    EnsureImmediateVisible
    Dbg "=== ProcessMassUpload START ==="

    Set wsMassUpload = ThisWorkbook.Sheets("MassUpload")
    Set tblMassUpload = wsMassUpload.ListObjects("MassUploadTable")

    If tblMassUpload.DataBodyRange Is Nothing Then
        MsgBox "No data available in the MassUpload table.", vbInformation
        Dbg "No data in MassUploadTable."
        Exit Sub
    End If

    ' Clear highlights
    Dim c As Range
    For Each rw In tblMassUpload.ListRows
        For Each c In rw.Range
            c.Interior.ColorIndex = xlNone
        Next c
    Next rw

    PrepareForMassUpload

    decimalSeparator = application.International(xlDecimalSeparator)
    notFoundMaterials = ""

    For Each rw In tblMassUpload.ListRows
        materialNumber = rw.Range(tblMassUpload.ListColumns("Component").Index).Value
        quantity = rw.Range(tblMassUpload.ListColumns("Quantity").Index).Value
        productNumber = rw.Range(tblMassUpload.ListColumns("Product Number").Index).Value
        netPriceUnit = rw.Range(tblMassUpload.ListColumns("Price per 1 unit").Index).Value

        ' Fallbacks for PN (so placeholders always get prefixed)
        If Len(Trim$(productNumber)) = 0 Then productNumber = lastPN
        If Len(Trim$(productNumber)) = 0 Then productNumber = CStr(ThisWorkbook.Sheets("1. BOM Definition").Range("F11").Value)
        If Len(Trim$(productNumber)) > 0 Then lastPN = productNumber

        Dbg "Row#" & rw.Index & _
            " | PN='" & Trim$(CStr(productNumber)) & "'" & _
            " | Comp='" & Trim$(CStr(materialNumber)) & "'" & _
            " | QtyRaw='" & CStr(quantity) & "'" & _
            " | PriceRaw='" & CStr(netPriceUnit) & "'"

        ' normalize decimals
        If InStr(quantity, ".") > 0 And decimalSeparator <> "." Then
            quantity = Replace(quantity, ".", decimalSeparator)
        ElseIf InStr(quantity, ",") > 0 And decimalSeparator <> "," Then
            quantity = Replace(quantity, ",", decimalSeparator)
        End If

        correctedQuantity = IIf(IsNumeric(quantity), CDbl(quantity), 0)
        Dbg "  QtyCorrected=" & correctedQuantity

        If correctedQuantity <= 0 Then
            rw.Range.Interior.Color = RGB(255, 255, 0)
            Dbg "  Invalid quantity -> highlight + skip."
            MsgBox "Invalid quantity for Component: " & materialNumber, vbExclamation
            GoTo NextRow
        End If

        Set matchingRows = FindMaterialRowsInExtendedMMD(materialNumber)

        If matchingRows Is Nothing Then
            Dbg "  No DB match for '" & materialNumber & "' -> AddPlaceholderComponentToBOM"
            notFoundMaterials = notFoundMaterials & materialNumber & vbCrLf
            rw.Range.Interior.Color = RGB(255, 0, 0)
            AddPlaceholderComponentToBOM materialNumber, correctedQuantity, productNumber, netPriceUnit
            GoTo NextRow

        Else
            Dim totalRowCount As Long, area As Range
            totalRowCount = 0
            For Each area In matchingRows.Areas
                totalRowCount = totalRowCount + area.Rows.Count
            Next area
            Dbg "  Found " & totalRowCount & " DB match(es) for '" & materialNumber & "'."

            If totalRowCount = 1 Then
                Dbg "  Auto-copy single match."
                CopyRowToDestination _
                    ThisWorkbook.Sheets("ExtendedMMD").Rows(matchingRows.row), _
                    correctedQuantity, _
                    productNumber, _
                    netPriceUnit
            Else
                Dbg "  Multiple matches -> selection form."
                LoadedDataSelectionForm.InitializeForm matchingRows
                LoadedDataSelectionForm.Show vbModal
                If LoadedDataSelectionForm.selectedRowIndex > 0 Then
                    Dbg "  User picked row " & LoadedDataSelectionForm.selectedRowIndex & " -> copy."
                    CopyRowToDestination _
                        ThisWorkbook.Sheets("ExtendedMMD").Rows(LoadedDataSelectionForm.selectedRowIndex), _
                        correctedQuantity, _
                        productNumber, _
                        netPriceUnit
                Else
                    Dbg "  User cancelled selection."
                End If
            End If
        End If

NextRow:
    Next rw

    If notFoundMaterials <> "" Then
        MsgBox "Process completed with placeholders. These components weren’t found in DB and were added to BOM with highlight:" & _
               vbCrLf & notFoundMaterials, vbExclamation
        Dbg "Placeholders added for: " & Replace(notFoundMaterials, vbCrLf, ", ")
    Else
        MsgBox "Mass upload process completed successfully.", vbInformation
        Dbg "Process completed successfully."
    End If

    SortSelectedComponentsByProduct
    Call Utils.RunProductBasedFormatting("1. BOM Definition", "BOMDefinition", "Helper Format BOMs")
    Dbg "=== ProcessMassUpload END ==="
End Sub


' =========================================================================================
'  4) CopyRowToDestination - unchanged from your original code
' =========================================================================================
Public Sub CopyRowToDestination(sourceRow As Range, quantity As Double, productNumber As String, netPriceUnit As Variant)
    Dim wsDestination As Worksheet
    Dim tblDestinationMMD As ListObject
    Dim tblSource As ListObject
    Dim newRowMMD As ListRow
    Dim headerCell As Range
    Dim destColIndex As Long

    ' Set the source and destination tables
    Set wsDestination = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblDestinationMMD = wsDestination.ListObjects("BOMDefinition")
    Set tblSource = ThisWorkbook.Sheets("ExtendedMMD").ListObjects("GetExtendedMMD")

    ' Add a new row to the destination table
    Set newRowMMD = tblDestinationMMD.ListRows.Add

    ' Loop through each header in the source table and copy data to the destination by matching column names
    For Each headerCell In tblSource.headerRowRange
        ' Find the corresponding column in the destination table based on the header
        On Error Resume Next
        destColIndex = tblDestinationMMD.ListColumns(headerCell.Value).Index
        On Error GoTo 0

        ' If a matching column exists in the destination table, copy the data
        If destColIndex > 0 Then
            newRowMMD.Range.Cells(1, destColIndex).Value = _
                sourceRow.Cells(1, headerCell.column - tblSource.headerRowRange.column + 1).Value
        End If
    Next headerCell

    ' Copy specific fields like Quantity, Product Number, and Price per 1 unit from MassUploadTable to the destination
    On Error Resume Next
    newRowMMD.Range.Cells(1, tblDestinationMMD.ListColumns("Quantity").Index).Value = quantity
    newRowMMD.Range.Cells(1, tblDestinationMMD.ListColumns("Product Number").Index).Value = productNumber
    
    ' Only set Price per 1 unit if it is not empty
    If netPriceUnit <> 0 Then
        newRowMMD.Range.Cells(1, tblDestinationMMD.ListColumns("Price per 1 unit").Index).Value = netPriceUnit
    End If
    On Error GoTo 0
End Sub


' =========================================================================================
'  5) FindMaterialRowsInExtendedMMD - with multi-area debug
' =========================================================================================
Private Function FindMaterialRowsInExtendedMMD(materialNumber As String) As Range
    Dim tblExtendedMMD As ListObject
    Dim foundCell As Range
    Dim firstAddress As String
    Dim matchingRows As Range
    
    Dim totalAreasCount As Long
    Dim totalRowsCount As Long
    Dim area As Range

    ' Set reference to GetExtendedMMD table
    Set tblExtendedMMD = ThisWorkbook.Sheets("ExtendedMMD").ListObjects("GetExtendedMMD")

    Debug.Print "Searching for materialNumber: " & materialNumber

    ' Find the first occurrence of the material number
    Set foundCell = tblExtendedMMD.ListColumns("Material").DataBodyRange.Find( _
        What:=materialNumber, _
        LookIn:=xlValues, _
        LookAt:=xlWhole) ' Change to xlPart if partial matches are needed
    
    If Not foundCell Is Nothing Then
        firstAddress = foundCell.Address
        Debug.Print "  Found first match at " & foundCell.Address & _
                    ", Row=" & foundCell.row & _
                    ", Value='" & foundCell.Value & "'"

        ' Start collecting rows
        Set matchingRows = foundCell.EntireRow

        Do
            Set foundCell = tblExtendedMMD.ListColumns("Material").DataBodyRange.FindNext(foundCell)
            If Not foundCell Is Nothing And foundCell.Address <> firstAddress Then
                Debug.Print "  Found additional match at " & foundCell.Address & _
                            ", Row=" & foundCell.row & _
                            ", Value='" & foundCell.Value & "'"
                
                ' Union the entire row
                Set matchingRows = Union(matchingRows, foundCell.EntireRow)
            Else
                Exit Do
            End If
        Loop

        ' Count total rows across all areas (debug)
        For Each area In matchingRows.Areas
            totalRowsCount = totalRowsCount + area.Rows.Count
        Next area
        Debug.Print "  Total matched row(s) across all areas: " & totalRowsCount
    Else
        Debug.Print "  No match found for materialNumber: " & materialNumber
    End If

    ' Return the union of all matching rows
    Set FindMaterialRowsInExtendedMMD = matchingRows
End Function


' =========================================================================================
'  6) ShowLoadedDataSelectionForm - from your original code (if still needed)
' =========================================================================================
Public Function ShowLoadedDataSelectionForm(matches As Range) As Long
    ' Initialize the form with the matching rows
    LoadedDataSelectionForm.InitializeForm matches

    ' Show the form and wait for the user’s selection
    LoadedDataSelectionForm.Show

    ' Return the selected row index (or -1 if canceled)
    ShowLoadedDataSelectionForm = LoadedDataSelectionForm.selectedRowIndex
End Function


' =========================================================================================
'  8) NEW: ProcessMassUploadSelectedPlant
'     Use this sub for a button that only searches for the plant in "1. BOM Definition"!C9.
' =========================================================================================
Public Sub ProcessMassUploadSelectedPlant()
    Dim wsMassUpload As Worksheet
    Dim tblMassUpload As ListObject
    Dim rowObj As ListRow
    Dim materialNumber As String
    Dim quantity As String
    Dim productNumber As String
    Dim netPriceUnit As Double
    Dim matchingRows As Range
    Dim correctedQuantity As Double
    Dim decimalSeparator As String
    Dim notFoundMaterials As String

    ' Read the selected plant from cell C9 in "1. BOM Definition"
    Dim selectedPlant As String
    selectedPlant = ThisWorkbook.Sheets("1. BOM Definition").Range("C9").Value
    
    Set wsMassUpload = ThisWorkbook.Sheets("MassUpload")
    Set tblMassUpload = wsMassUpload.ListObjects("MassUploadTable")

    If tblMassUpload.DataBodyRange Is Nothing Then
        MsgBox "No data available in the MassUpload table.", vbInformation
        Exit Sub
    End If

    ' Clear existing highlights
    Dim cell As Range
    For Each rowObj In tblMassUpload.ListRows
        For Each cell In rowObj.Range
            cell.Interior.ColorIndex = xlNone
        Next cell
    Next rowObj

    ' Prepare ExtendedMMD
    PrepareForMassUpload

    decimalSeparator = application.International(xlDecimalSeparator)
    notFoundMaterials = ""

    ' Loop each row in MassUpload
    For Each rowObj In tblMassUpload.ListRows
        materialNumber = rowObj.Range(tblMassUpload.ListColumns("Component").Index).Value
        quantity = rowObj.Range(tblMassUpload.ListColumns("Quantity").Index).Value
        productNumber = rowObj.Range(tblMassUpload.ListColumns("Product Number").Index).Value
        netPriceUnit = rowObj.Range(tblMassUpload.ListColumns("Price per 1 unit").Index).Value

        ' Fix decimals
        If InStr(quantity, ".") > 0 And decimalSeparator <> "." Then
            quantity = Replace(quantity, ".", decimalSeparator)
        ElseIf InStr(quantity, ",") > 0 And decimalSeparator <> "," Then
            quantity = Replace(quantity, ",", decimalSeparator)
        End If

        If IsNumeric(quantity) Then
            correctedQuantity = CDbl(quantity)
        Else
            correctedQuantity = 0
        End If

        If correctedQuantity <= 0 Then
            rowObj.Range.Interior.Color = RGB(255, 255, 0)
            MsgBox "Invalid quantity for Component: " & materialNumber, vbExclamation
            GoTo NextRow
        End If

        ' === Find rows that match BOTH Material & selectedPlant
        Set matchingRows = FindRowsByMaterialAndPlant(materialNumber, selectedPlant)

        If matchingRows Is Nothing Then
            notFoundMaterials = notFoundMaterials & materialNumber & vbCrLf
            rowObj.Range.Interior.Color = RGB(255, 0, 0)
            GoTo NextRow
        Else
            Dim totalRowCount As Long
            totalRowCount = 0
            Dim area As Range
            For Each area In matchingRows.Areas
                totalRowCount = totalRowCount + area.Rows.Count
            Next area

            Debug.Print "Material " & materialNumber & ", Plant " & selectedPlant & _
                        " => totalRowCount: " & totalRowCount

            If totalRowCount = 1 Then
                Debug.Print "  Only 1 row found (Material + Plant). Copying automatically..."
                CopyRowToDestination _
                    ThisWorkbook.Sheets("ExtendedMMD").Rows(matchingRows.row), _
                    correctedQuantity, _
                    productNumber, _
                    netPriceUnit
            Else
                Debug.Print "  Multiple Material+Plant matches found. Opening selection form..."
                LoadedDataSelectionForm.InitializeForm matchingRows
                LoadedDataSelectionForm.Show vbModal
                If LoadedDataSelectionForm.selectedRowIndex > 0 Then
                    CopyRowToDestination _
                        ThisWorkbook.Sheets("ExtendedMMD").Rows(LoadedDataSelectionForm.selectedRowIndex), _
                        correctedQuantity, _
                        productNumber, _
                        netPriceUnit
                End If
            End If
        End If
NextRow:
    Next rowObj

    If notFoundMaterials <> "" Then
        MsgBox "Process completed with errors. The following components+plants were not found:" & _
               vbCrLf & notFoundMaterials, vbExclamation
    Else
        MsgBox "Mass upload by Plant completed successfully.", vbInformation
    End If

    SortSelectedComponentsByProduct
End Sub

' =========================================================================================
'  9) NEW: FindRowsByMaterialAndPlant - returns only rows matching both "Material" & "Plant"
' =========================================================================================
Private Function FindRowsByMaterialAndPlant(materialNumber As String, selectedPlant As String) As Range
    Dim allMaterialRows As Range
    Dim plantColIndex As Long
    Dim matchedRows As Range
    Dim r As Range

    ' 1) First, get all rows that match the material (just like your original function).
    Set allMaterialRows = FindMaterialRowsInExtendedMMD(materialNumber)
    If allMaterialRows Is Nothing Then
        ' If no row matched the Material, no need to filter further
        Exit Function
    End If

    ' 2) Identify which column in "GetExtendedMMD" is named "Plant"
    Dim tblExtendedMMD As ListObject
    Set tblExtendedMMD = ThisWorkbook.Sheets("ExtendedMMD").ListObjects("GetExtendedMMD")
    On Error Resume Next
    plantColIndex = tblExtendedMMD.ListColumns("Plant").Index
    On Error GoTo 0

    If plantColIndex = 0 Then
        ' No "Plant" column found
        MsgBox "Could not locate the 'Plant' column in the GetExtendedMMD table!", vbExclamation
        Exit Function
    End If

    ' 3) Among those rows, keep only the ones whose "Plant" cell matches selectedPlant
    '    We'll build a new union of rows that have the matching Plant
    Dim area As Range, rowCheck As Range
    For Each area In allMaterialRows.Areas
        For Each rowCheck In area.Rows
            ' Check the "Plant" cell in this row
            ' We subtract the table's header row offset to get the correct row in DataBodyRange
            Dim relativeRow As Long
            relativeRow = rowCheck.row - tblExtendedMMD.headerRowRange.row

            ' Safeguard: must be within table's data
            If relativeRow >= 1 And relativeRow <= tblExtendedMMD.DataBodyRange.Rows.Count Then
                ' The actual cell for "Plant" in this row:
                Dim plantCell As Range
                Set plantCell = tblExtendedMMD.DataBodyRange.Rows(relativeRow).Cells(1, plantColIndex)

                If Trim(CStr(plantCell.Value)) = Trim(CStr(selectedPlant)) Then
                    ' This row matches both the material & plant
                    If matchedRows Is Nothing Then
                        Set matchedRows = rowCheck
                    Else
                        Set matchedRows = Union(matchedRows, rowCheck)
                    End If
                End If
            End If
        Next rowCheck
    Next area

    Set FindRowsByMaterialAndPlant = matchedRows
End Function
Public Sub AddPlaceholderComponentToBOM( _
    ByVal materialPN As String, _
    ByVal qty As Double, _
    ByVal productNumber As String, _
    Optional ByVal netPriceUnit As Variant)

    Dim wsDest As Worksheet
    Dim tblDest As ListObject
    Dim newRow As ListRow

    Dim cMat As Long, cQty As Long, cProd As Long, cPrice As Long
    Dim cProdType As Long, cPlant As Long, cNewComp As Long, cDescr As Long

    Dim selectedProdType As Variant
    Dim selectedPlant As Variant
    Dim materialName As String

    ' --- Context ---
    EnsureImmediateVisible
    selectedPlant = ThisWorkbook.Sheets("1. BOM Definition").Range("C9").Value
    If Len(Trim$(productNumber)) = 0 Then productNumber = CStr(ThisWorkbook.Sheets("1. BOM Definition").Range("F11").Value)
    selectedProdType = ThisWorkbook.Sheets("1. BOM Definition").Range("D8").Value

    Dbg "AddPlaceholder | PN='" & Trim$(CStr(productNumber)) & "' | Typed='" & Trim$(CStr(materialPN)) & _
        "' | Qty=" & qty & " | Price=" & IIf(IsMissing(netPriceUnit), "(missing)", CStr(netPriceUnit)) & _
        " | Plant(C9)='" & Trim$(CStr(selectedPlant)) & "'"

    ' Destination
    Set wsDest = ThisWorkbook.Sheets("1. BOM Definition")
    Set tblDest = wsDest.ListObjects("BOMDefinition")
    Set newRow = tblDest.ListRows.Add
    Dbg "  New BOM row -> table index " & newRow.Index & ", sheet row " & newRow.Range.row

    ' Column indexes
    On Error Resume Next
    cProd = tblDest.ListColumns("Product Number").Index
    cProdType = tblDest.ListColumns("Product Type").Index
    cPlant = tblDest.ListColumns("Plant").Index
    cMat = tblDest.ListColumns("Material").Index
    cQty = tblDest.ListColumns("Quantity").Index
    cPrice = tblDest.ListColumns("Price per 1 unit").Index
    cNewComp = tblDest.ListColumns("New component").Index
    cDescr = tblDest.ListColumns("Material Description").Index
    On Error GoTo 0

    Dbg "  ColIdx: Prod=" & cProd & ", ProdType=" & cProdType & ", Plant=" & cPlant & _
        ", Material=" & cMat & ", Qty=" & cQty & ", Price=" & cPrice & ", NewComp=" & cNewComp & ", Descr=" & cDescr

    ' Values
    If cProd > 0 Then newRow.Range.Cells(1, cProd).Value = productNumber
    If cProdType > 0 Then newRow.Range.Cells(1, cProdType).Value = selectedProdType

    ' ALWAYS set Plant from C9
    If cPlant > 0 Then
        newRow.Range.Cells(1, cPlant).Value = selectedPlant
        Dbg "  Wrote Plant '" & selectedPlant & "' -> " & newRow.Range.Cells(1, cPlant).Address
    End If

    ' Material as "<PN>-<typed>"
    materialName = EnsurePNPrefix(productNumber, materialPN)
    If cMat > 0 Then
        newRow.Range.Cells(1, cMat).Value = materialName
        newRow.Range.Cells(1, cMat).Interior.Color = RGB(255, 255, 0)
        Dbg "  Wrote Material '" & materialName & "' -> " & newRow.Range.Cells(1, cMat).Address
    End If

    ' Optional description
    If cDescr > 0 Then
        If Len(Trim$(CStr(newRow.Range.Cells(1, cDescr).Value))) = 0 Then
            newRow.Range.Cells(1, cDescr).Value = "Mass upload placeholder for: " & CStr(materialPN)
            Dbg "  Wrote description trace."
        End If
    End If

    If cNewComp > 0 Then newRow.Range.Cells(1, cNewComp).Value = "NEW"
    If cQty > 0 Then newRow.Range.Cells(1, cQty).Value = qty
    If cPrice > 0 Then
        If Not IsMissing(netPriceUnit) Then
            If IsNumeric(netPriceUnit) And CDbl(netPriceUnit) <> 0 Then
                newRow.Range.Cells(1, cPrice).Value = CDbl(netPriceUnit)
            End If
        End If
    End If

    ' Read back confirm
    On Error Resume Next
    Dbg "  Confirm -> Plant='" & IIf(cPlant > 0, CStr(newRow.Range.Cells(1, cPlant).Value), "(n/a)") & _
        "', Material='" & IIf(cMat > 0, CStr(newRow.Range.Cells(1, cMat).Value), "(n/a)") & "'"
    On Error GoTo 0
End Sub


' Build "<PN>-<typed>" and avoid double-prefixing if user already typed PN-...
Private Function EnsurePNPrefix(ByVal productNumber As String, ByVal typed As String) As String
    Dim pn As String, t As String, outVal As String
    pn = Trim$(CStr(productNumber))
    t = Trim$(CStr(typed))

    If Len(pn) = 0 Then
        outVal = t
    ElseIf LCase$(Left$(t, Len(pn) + 1)) = LCase$(pn & "-") Then
        outVal = t
    Else
        outVal = pn & "-" & t
    End If

    Dbg "EnsurePNPrefix | PN='" & pn & "' | Typed='" & t & "' | Out='" & outVal & "'"
    EnsurePNPrefix = outVal
End Function


' Returns the next "<ProductNumber>-New#" by scanning BOMDefinition[Material]
' Strictly matches: ^<ProductNumber>-New(\d+)$ (case-insensitive)
Private Function NextNewMaterialName(ByVal productNumber As String) As String
    Dim tbl As ListObject, rng As Range, cell As Range
    Dim regex As Object
    Dim maxIdx As Long, m As Object
    
    If Len(Trim$(productNumber)) = 0 Then
        NextNewMaterialName = "New1"  ' fallback, but AddPlaceholder will try to fix PN first
        Exit Function
    End If
    
    Set tbl = ThisWorkbook.Sheets("1. BOM Definition").ListObjects("BOMDefinition")
    On Error Resume Next
    Set rng = tbl.ListColumns("Material").DataBodyRange
    On Error GoTo 0

    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .Global = False
        .IgnoreCase = True
        .pattern = "^" & Replace(productNumber, "-", "\-") & "-New(\d+)$"
    End With
    
    maxIdx = 0
    If Not rng Is Nothing Then
        For Each cell In rng
            If regex.Test(CStr(cell.Value)) Then
                Set m = regex.Execute(CStr(cell.Value))(0)
                If CLng(m.SubMatches(0)) > maxIdx Then maxIdx = CLng(m.SubMatches(0))
            End If
        Next cell
    End If
    
    NextNewMaterialName = productNumber & "-New" & (maxIdx + 1)
End Function

' Lightweight logger to Immediate Window (Ctrl+G)
Private Sub Dbg(ByVal msg As String)
    On Error Resume Next
    Debug.Print Format(Now, "hh:nn:ss") & " [MassUpload] " & msg
End Sub

' Try to show the Immediate window (needs Trust access to VBOM enabled; otherwise silently ignored)
Private Sub EnsureImmediateVisible()
    On Error Resume Next
    application.VBE.Windows.item("Immediate").Visible = True
End Sub

