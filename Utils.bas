Attribute VB_Name = "Utils"
Public Function FixDecimalSeparator(ByVal inputValue As String) As String
    Dim decimalSeparator As String
    decimalSeparator = application.International(xlDecimalSeparator)
    
    If InStr(inputValue, ".") > 0 And decimalSeparator = "," Then
        inputValue = Replace(inputValue, ".", ",")
    ElseIf InStr(inputValue, ",") > 0 And decimalSeparator = "." Then
        inputValue = Replace(inputValue, ",", ".")
    End If
    
    FixDecimalSeparator = inputValue
End Function
Public Sub FixDecimalSeparatorsInTables()
    Call FixBOMDefinitionDecimalSeparators
    Call FixSelectedRoutinesDecimalSeparators
    ' MsgBox "Decimal separators fixed in BOMDefinition and SelectedRoutines.", vbInformation
End Sub

Private Sub FixBOMDefinitionDecimalSeparators()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim columnNames As Variant
    Dim name As Variant
    
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")
    
    columnNames = Array("Quantity", "Price per 1 unit", "Net weight [kg/Base unit]", "Copper weight [kg/1000m]")
    
    For Each name In columnNames
        On Error Resume Next
        Set col = tbl.ListColumns(name)
        On Error GoTo 0
        
        If Not col Is Nothing Then
            For Each cell In col.DataBodyRange
                If Not IsEmpty(cell.value) And VarType(cell.value) = vbString Then
                    cell.value = FixDecimalSeparator(cell.value)
                End If
            Next cell
        End If
    Next name
End Sub
Private Sub FixSelectedRoutinesDecimalSeparators()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim col As ListColumn
    Dim cell As Range
    Dim columnNames As Variant
    Dim name As Variant

    Set ws = ThisWorkbook.Sheets("2. Routines")
    Set tbl = ws.ListObjects("SelectedRoutines")

    columnNames = Array("tr", "te", "Number of Operations", "Number of Setups")

    For Each name In columnNames
        On Error Resume Next
        Set col = tbl.ListColumns(name)
        On Error GoTo 0

        If Not col Is Nothing Then
            If Not col.DataBodyRange Is Nothing Then
                For Each cell In col.DataBodyRange
                    If Not IsEmpty(cell.value) And VarType(cell.value) = vbString Then
                        cell.value = FixDecimalSeparator(cell.value)
                    End If
                Next cell
            End If
        End If
    Next name
End Sub

Sub btnOpenChainForm_Click()
    ThisWorkbook.Sheets("Page 1 - Chain RFQ Form").Activate
End Sub

Sub ExportAllVBACode()
    ' ?? Requires "Trust access to the VBA project object model" to be enabled in Macro settings
    ' Tools > References > Enable "Microsoft Visual Basic for Applications Extensibility"
    
    Dim vbComp As VBIDE.VBComponent
    Dim exportPath As String
    Dim fileName As String

    ' Set export directory (adjust path as needed)
    exportPath = "C:\Users\Adrian D\OneDrive - SC Dreams of the Future SRL\Desktop\VBAd"
    
    ' Create directory if it doesn't exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    ' Loop through all components
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type
            Case vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_Document
                fileName = exportPath & vbComp.name & ".cls"
            Case vbext_ct_MSForm
                fileName = exportPath & vbComp.name & ".frm"
            Case Else
                fileName = exportPath & vbComp.name & ".txt"
        End Select
        
        vbComp.Export fileName
    Next vbComp

    MsgBox "VBA code exported to: " & exportPath, vbInformation
End Sub

'''
' Applies the correct number formatting to an entire row based on column names.
' @param {ListRow} targetRow The ListRow object in the table to be formatted.
'''
'''
' Applies the correct number formatting to an entire row based on column names.
' @param {ListRow} targetRow The ListRow object in the BOMDefinition table to be formatted.
'''
Public Sub ApplyRowFormatting(ByVal targetRow As ListRow)
    If targetRow Is Nothing Then Exit Sub

    Dim loBom As ListObject
    Dim col As ListColumn
    Dim targetCell As Range

    On Error Resume Next
    Set loBom = targetRow.Parent
    If loBom Is Nothing Then Exit Sub
    On Error GoTo 0

    ' Loop through each column of the table to format the cell in the target row
    For Each col In loBom.ListColumns
        Set targetCell = targetRow.Range(col.Index)
        
        ' Leave cells with formulas as is
        If Not targetCell.HasFormula Then
            ' Apply formatting rules based on the column name
            Select Case col.name
                Case "Price per 1 unit"
                    targetCell.NumberFormat = "0.0000"
                Case "Copper weight [kg/1000m]", "Net weight (kg/Base unit)"
                    targetCell.NumberFormat = "0.000"
                Case "Quantity", "te"
                    ' Use a standard number format for quantity
                    targetCell.NumberFormat = "0.00"
                    
                Case "tr", "Number of operations", "Number of Setups", "Batch", "AOQ"
                    ' Leave key identifier columns with their default "General" format
                    targetCell.NumberFormat = "0"
                Case Else
                    ' Set all other data columns to Text format to preserve leading zeros, etc.
                    targetCell.NumberFormat = "@"
            End Select
            
        End If
    Next col
End Sub


Sub ShowSourceDataOfBoms()
    ThisWorkbook.Sheets("Purchasing Info Records").Visible = True
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsNumericRegex
' Author    : Gemini
' Date      : 12.06.2025
' Purpose   : Checks if a string is a numeric value using a regular expression.
'             It accepts an optional leading sign ("-") and either a single "."
'             or "," as a decimal separator.
'
' Parameter : inputText (String) - The string value to check.
'
' Returns   : Boolean - True if the string matches the numeric pattern, otherwise False.
'---------------------------------------------------------------------------------------
Public Function IsNumericRegex(ByVal inputText As String) As Boolean
    ' This function uses a regular expression to validate if a string is a number.
    ' This method gives very precise control over the accepted format.

    ' We use late binding here with CreateObject("VBScript.RegExp").
    ' This avoids needing to set a reference to the "Microsoft VBScript Regular Expressions"
    ' library, which makes the code more portable across different machines.
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    ' The pattern is defined as follows:
    ' ^      - Asserts the start of the string.
    ' -?     - Allows an optional hyphen (negative sign) at the beginning.
    ' \d+    - Matches one or more digits (the integer part).
    ' (      - Starts an optional group for the decimal part.
    ' [.,]   - Matches a single character that is either a comma or a period.
    ' \d+    - Matches one or more digits (the fractional part).
    ' )?     - Closes the optional group, making the whole decimal part optional.
    ' $      - Asserts the end of the string.
    ' This ensures the *entire* string must match the number format.
    regex.pattern = "^-?\d+([.,]\d+)?$"

    ' The .Test() method of the RegExp object returns True if the pattern
    ' is found in the input string, and False otherwise.
    IsNumericRegex = regex.Test(inputText)

End Function



' ===================================================================
' Hauptprozedur zum Aufrufen der finalen, dynamischen Formatierung
' ===================================================================
Sub RunProductBasedFormatting(associatedSheetname As String, associatedTableName As String)
    application.ScreenUpdating = False

    Dim wsProducts As Worksheet
    Dim wsData As Worksheet
    Dim tblFinalProducts As ListObject
    Dim tblAssociatedData As ListObject
    Dim productColumn As ListColumn
    Dim productCell As Range
    
    Dim baseColor1 As Long, baseColor2 As Long
    Dim currentColor As Long, shadeColor As Long
    Dim productIndex As Long
    
    ' *** ANPASSEN ***
    ' Gib hier die Namen der beiden Arbeitsblätter an.
    Set wsProducts = ThisWorkbook.Sheets("Final Products") ' <-- Name des Blattes mit der Produktliste
    Set tblFinalProducts = wsProducts.ListObjects("FinalProductList")
    
    Set wsData = ThisWorkbook.Sheets(associatedSheetname)         ' <-- Name des Blattes mit der Datentabelle
    Set tblAssociatedData = wsData.ListObjects("BOMDefinition") ' <-- Name der zu formatierenden Tabelle
    
    ' Definiere deine ZWEI abwechselnden Grundfarben
    baseColor1 = RGB(235, 241, 250)    ' Farbe für ungerade Produkte (1, 3, ...)
    baseColor2 = RGB(250, 243, 233)   ' Farbe für gerade Produkte (2, 4, ...)
    
    ' Gib den Namen der Spalte in der Zieltabelle an, die die Produktnamen enthält.
    Set productColumn = tblAssociatedData.ListColumns("ProductNumberText") ' <-- Spaltenname anpassen
    ' *** ENDE ANPASSUNG ***

    ' 1. Alle bestehenden Farben aus der Zieltabelle entfernen
    tblAssociatedData.DataBodyRange.Interior.Color = xlNone
    If wsData.AutoFilterMode Then wsData.AutoFilter.ShowAllData
    
    ' Zähler für die Produktposition initialisieren
    productIndex = 1

    ' 2. Schleife durch jedes Endprodukt in deiner Haupttabelle
    For Each productCell In tblFinalProducts.ListColumns(2).DataBodyRange
        
        ' 3. Wähle die Grundfarbe basierend auf der geraden/ungeraden Position des Produkts
        If productIndex Mod 2 <> 0 Then
            currentColor = baseColor1
        Else
            currentColor = baseColor2
        End If
        
        ' Berechne die hellere Schattenfarbe
        shadeColor = LightenColor(currentColor, 0.6)
        
        ' 4. Formatiere die zugehörigen Zeilen in der Zieltabelle
        FormatRowsForProduct tblAssociatedData, productColumn.Index, productCell.value, currentColor, shadeColor
        
        ' Zähler für das nächste Produkt erhöhen
        productIndex = productIndex + 1
        
    Next productCell
    
    ' Filter am Ende auf dem Datenblatt sicherheitshalber entfernen
    On Error Resume Next
    wsData.AutoFilter.ShowAllData
    On Error GoTo 0
    application.ScreenUpdating = True
End Sub


' ===================================================================
' Formatiert die Zeilen für ein bestimmtes Produkt in einer Tabelle (unverändert)
' ===================================================================
Private Sub FormatRowsForProduct(ByVal targetTable As ListObject, ByVal colIndex As Long, ByVal criteria As String, ByVal color1 As Long, ByVal color2 As Long)
    targetTable.Range.AutoFilter Field:=colIndex, Criteria1:=criteria
    
    Dim visibleRowCounter As Long
    Dim lr As ListRow
    visibleRowCounter = 1
    
    For Each lr In targetTable.ListRows
        If Not lr.Range.EntireRow.Hidden Then
            If visibleRowCounter Mod 2 <> 0 Then
                lr.Range.Interior.Color = color1
            Else
                lr.Range.Interior.Color = color2
            End If
            visibleRowCounter = visibleRowCounter + 1
        End If
    Next lr
    
End Sub

' ===================================================================
' HILFSFUNKTION: Hellt eine gegebene Farbe auf (unverändert)
' ===================================================================
Public Function LightenColor(ByVal baseColor As Long, ByVal factor As Double) As Long
    Dim r As Long, g As Long, b As Long
    r = baseColor Mod 256
    g = (baseColor \ 256) Mod 256
    b = (baseColor \ 65536) Mod 256
    r = r + (255 - r) * factor
    g = g + (255 - g) * factor
    b = b + (255 - b) * factor
    If r > 255 Then r = 255
    If g > 255 Then g = 255
    If b > 255 Then b = 255
    LightenColor = RGB(r, g, b)
End Function
