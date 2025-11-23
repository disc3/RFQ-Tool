Attribute VB_Name = "CopyDataFromBM1Sheet"
Sub Makro2()
'
' Sub CopyFilteredComponentsWithoutSpecialCells()
    Dim srcWorkbook As Workbook
    Dim srcSheet As Worksheet
    Dim tgtSheet As Worksheet
    Dim lastRowSrc As Long
    Dim lastRowTgt As Long
    Dim currentRowTgt As Long
    Dim productInfo As String
    Dim i As Long
    Dim workbookFound As Boolean
    Dim tgtWorkbook As Workbook
    Dim wb As Workbook
    Dim sheetFound As Boolean
    Dim firstRowWritten As Boolean

    ' Hledání sešitu s listem BM1 mezi otevøenými sešity
    workbookFound = False
    For Each srcWorkbook In application.Workbooks
        On Error Resume Next
        Set srcSheet = srcWorkbook.Sheets("BM1")
        On Error GoTo 0
        If Not srcSheet Is Nothing Then
            workbookFound = True
            Exit For
        End If
    Next srcWorkbook

    ' Pokud sešit nebyl nalezen, vyžádat vstup od uživatele
    If Not workbookFound Then
        MsgBox "Sešit obsahující list 'BM1' nebyl nalezen. Vyberte soubor ruènì.", vbExclamation
        Dim filePath As Variant
        filePath = application.GetOpenFilename("Excel Files (*.xlsm), *.xlsm")
        If filePath = False Then
            MsgBox "Soubor nebyl vybrán. Operace byla zrušena.", vbCritical
            Exit Sub
        End If
        Set srcWorkbook = Workbooks.Open(filePath)
        On Error Resume Next
        Set srcSheet = srcWorkbook.Sheets("BM1")
        On Error GoTo 0
        If srcSheet Is Nothing Then
            MsgBox "Vybraný soubor neobsahuje list 'BM1'. Operace byla zrušena.", vbCritical
            Exit Sub
        End If
    End If

    ' Nastavení cílového sešitu a listu
    'Set tgtWorkbook = Workbooks("Import.xlsm")
    'Set tgtSheet = tgtWorkbook.Sheets(1) ' Nastavte správný list cílového souboru
    
    sheetFound = False
    ' Projdeme všechny otevøené sešity
    For Each wb In application.Workbooks
     On Error Resume Next ' Zabrání chybì, pokud list neexistuje
        Set tgtSheet = wb.Sheets("Template_BOM_Connect")
        On Error GoTo 0
        If Not tgtSheet Is Nothing Then
            Set tgtWorkbook = wb
            sheetFound = True
            Exit For
        End If
    Next wb

    ' Nalezení posledního øádku ve zdrojovém souboru (sloupec F)
    lastRowSrc = srcSheet.Cells(srcSheet.Rows.Count, "F").End(xlUp).row
    ' Nalezení posledního øádku v cílovém souboru (sloupec I)
    lastRowTgt = tgtSheet.Cells(tgtSheet.Rows.Count, "I").End(xlUp).row
    currentRowTgt = lastRowTgt + 1

firstRowWritten = False ' Indikuje, zda byl první øádek již vyplnìn

' Iterace pøes všechny øádky ve zdrojovém souboru (od øádku 11)
For i = 11 To lastRowSrc
    ' Naètení informace o produktu (sloupec B, odpovídající aktuálnímu øádku)
    productInfo = srcSheet.Cells(i, "B").Value
    
    ' Kontrola, zda je hodnota ve sloupci F neprázdná
    If srcSheet.Cells(i, "F").Value <> "" Then
    
        ' Zápis øádku s kompletními daty
        tgtSheet.Cells(currentRowTgt, "B").Value = productInfo ' Informace o produktu
        ' Posun na další øádek v cílovém souboru
        currentRowTgt = currentRowTgt + 1
        tgtSheet.Cells(currentRowTgt, "H").Value = srcSheet.Cells(i, "D").Value ' Number of Item
        tgtSheet.Cells(currentRowTgt, "I").Value = srcSheet.Cells(i, "F").Value ' Art. Number
        tgtSheet.Cells(currentRowTgt, "J").Value = srcSheet.Cells(i, "G").Value ' Description
        tgtSheet.Cells(currentRowTgt, "K").Value = srcSheet.Cells(i, "H").Value ' Quantity
        tgtSheet.Cells(currentRowTgt, "L").Value = srcSheet.Cells(i, "I").Value ' Unit
    End If
    
    ' Kontrola, zda je nìkterá buòka (D, F, G, H, I) prázdná
    If srcSheet.Cells(i, "D").Value = "" Then
            ' Posun na další øádek v cílovém souboru
        currentRowTgt = currentRowTgt + 1
        ' Zápis prázdného øádku s pouze informací o produktu
        'tgtSheet.Cells(currentRowTgt, "B").Value = productInfo ' Zápis pouze do sloupce B
    End If
Next i

    ' Hotovo
    MsgBox "Kopírování s filtrováním dokonèeno!", vbInformation

End Sub

