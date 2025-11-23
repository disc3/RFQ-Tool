Attribute VB_Name = "Sort"
Sub SortSelectedComponentsByProduct()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim sortColumn As ListColumn
    Dim lastRow As ListRow
    Dim cell As Range
    Dim isRowEmpty As Boolean
    
    ' Set the worksheet where the table is located
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    
    ' Set the table (ListObject)
    Set tbl = ws.ListObjects("BOMDefinition")
    
    ' Set the column to sort by "Product Number"
    Set sortColumn = tbl.ListColumns("Product Number")
    
    ' Check if the table has any data
    If tbl.ListRows.Count > 0 Then
        ' Sort the table by the "Product Number" column in ascending order
        With tbl.Sort
            .SortFields.Clear ' Clear any previous sort fields
            .SortFields.Add key:=sortColumn.Range, Order:=xlAscending ' Add the sort field
            .header = xlYes ' Indicate that the first row contains headers
            .Apply ' Apply the sort
        End With

        ' Check if the last row is empty (excluding formulas) and delete if necessary
        Do While tbl.ListRows.Count > 0
            Set lastRow = tbl.ListRows(tbl.ListRows.Count) ' Get the last row
            isRowEmpty = True ' Assume the row is empty
            
            ' Loop through each cell in the row to check if it's empty (excluding formulas)
            For Each cell In lastRow.Range
                If Not IsEmpty(cell.Value) Then
                    If cell.HasFormula Then
                        isRowEmpty = True
                    Else
                        isRowEmpty = False ' If any cell has a value, the row is not empty
                    Exit For
                    End If
                End If
            Next cell
            
            ' Delete the row if it is empty (excluding formulas)
            If isRowEmpty Then
                lastRow.Delete
            Else
                Exit Do ' Exit the loop if the last row is not empty
            End If
        Loop
        
        ' MsgBox "Table sorted by Product Number (in order of definition) and empty rows removed.", vbInformation
    Else
        ' MsgBox "No data to sort in the SelectedComponents table.", vbExclamation
    End If
    On Error Resume Next
    Call Utils.RunProductBasedFormatting(ws.name, tbl.name)
    On Error GoTo 0
    ' Call SortSelectedRoutingByProduct after this function
    SortSelectedRoutingByProduct
End Sub



Sub SortSelectedRoutingByProduct()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim sortColumnProduct As ListColumn
    Dim sortColumnSortOrder As ListColumn

    ' Set the worksheet where the table is located
    Set ws = ThisWorkbook.Sheets("2. Routines") ' Replace with your actual sheet name

    ' Set the table (ListObject)
    Set tbl = ws.ListObjects("SelectedRoutines") ' Ensure this is the correct name of your table

    ' Set the columns to sort by
    Set sortColumnProduct = tbl.ListColumns("Product Number") ' Ensure this is the correct name of your Product column
    Set sortColumnSortOrder = tbl.ListColumns("Sort Order") ' Ensure this is the correct name of your Sort Order column

    ' Check if the table has any data
    If tbl.ListRows.Count > 0 Then
        ' Sort the table by Product Number first and then by Sort Order
        With tbl.Sort
            .SortFields.Clear ' Clear any previous sort fields
            
            ' Sort by Product Number (ascending)
            .SortFields.Add key:=sortColumnProduct.Range, Order:=xlAscending
            
            ' Then sort by Sort Order (ascending)
            .SortFields.Add key:=sortColumnSortOrder.Range, Order:=xlAscending
            
            .header = xlYes ' Indicate that the first row contains headers
            .Apply ' Apply the sort
        End With
        
        ' Optionally, you can uncomment this line to display a message box after sorting
        ' MsgBox "Table sorted by Product Number and Sort Order (ascending).", vbInformation
    Else
        ' Display a message if there is no data to sort
        ' MsgBox "No data to sort in the SelectedRoutines table.", vbExclamation
    End If
    On Error Resume Next
    Call Utils.RunProductBasedFormatting(ws.name, tbl.name)
    On Error GoTo 0
End Sub

