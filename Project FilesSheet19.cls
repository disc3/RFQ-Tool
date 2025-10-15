VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet19"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub tlaèítko5_Kliknutí_bom_import()


    

    Dim rng As Range, cell As Range, copyToCell As Range
    
    ActiveWorkbook.Sheets("Components costs").Select
    Set rng = Sheet.Range("E3:E30") 'This is the range you want to copy from
    ActiveWorkbook.Sheets("Template_BOM_Connect").Select
    Set copyToCell = ThisWorkbook.Sheets("Template_BOM_Connect").Range("E3") 'This is the first cell in the range you want to copy to
    Set copyToCell_name = ThisWorkbook.Sheets("Template_BOM_Connect").Range("B3") 'This is the first cell in the range you want to copy to
    Set copyToCell_quantity = ThisWorkbook.Sheets("Template_BOM_Connect").Range("G3") 'This is the first cell in the range you want to copy to
    
    For Each cell In rng
            copyToCell.value = cell.value
            Set copyToCell = copyToCell.offset(1, 0) 'Move the copy to cell over by one column (or by row if you prefer)
    Next cell

    For Each cell In rng
            copyToCell_name.value = cell.offset(0, -1).value
            Set copyToCell_name = copyToCell_name.offset(1, 0) 'Move the copy to cell over by one column (or by row if you prefer)
    Next cell

    For Each cell In rng
            copyToCell_quantity.value = cell.offset(0, 3).value
            Set copyToCell_quantity = copyToCell_quantity.offset(1, 0) 'Move the copy to cell over by one column (or by row if you prefer)
    Next cell
    
End Sub
