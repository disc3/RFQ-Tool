Attribute VB_Name = "Utils"
Option Explicit
Private lCalcSave As Long

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
                If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
                    cell.Value = FixDecimalSeparator(cell.Value)
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
                    If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
                        cell.Value = FixDecimalSeparator(cell.Value)
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
                    
                Case "Number of operations"
                    targetCell.NumberFormat = "0.##"
                Case "tr", "Number of Setups", "Batch", "AOQ"
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

Option Explicit

' ===================================================================
' OPTIMIZED SUBROUTINE: RunProductBasedFormatting
' ===================================================================
' PURPOSE: Formats a target table based on Product grouping.
' IMPROVEMENTS:
' 1. Uses Arrays for data reading (Instant).
' 2. Uses Dictionary for Product Index lookup (No nested loops).
' 3. Uses Union Ranges to apply color in one batch (No row-by-row painting).
' 4. Accepts 'helperColName' to distinguish between BOM and Routing logic.
' ===================================================================
Sub RunProductBasedFormatting(associatedSheetname As String, associatedTableName As String, helperColName As String)
    On Error GoTo ErrorHandler
    
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual

    ' --- 1. SETUP WORKSHEETS & TABLES ---
    Dim wsProducts As Worksheet: Set wsProducts = ThisWorkbook.Sheets("Final Products")
    Dim tblFinalProducts As ListObject: Set tblFinalProducts = wsProducts.ListObjects("FinalProductList")
    
    Dim wsData As Worksheet: Set wsData = ThisWorkbook.Sheets(associatedSheetname)
    Dim tblAssociatedData As ListObject: Set tblAssociatedData = wsData.ListObjects(associatedTableName)
    
    ' --- CHECK: Is the table "visually" empty? ---
    If tblAssociatedData.DataBodyRange Is Nothing Then GoTo CleanExit
    If tblFinalProducts.DataBodyRange Is Nothing Then GoTo CleanExit
    
    ' (Contains 1 row, but the Product Number cell is blank/empty)
    If tblAssociatedData.ListRows.Count = 1 Then
        Dim firstVal As Variant
        ' Note: Using "ProductNumberText" as per your previous target table logic
        firstVal = tblAssociatedData.DataBodyRange(1, tblAssociatedData.ListColumns("ProductNumberText").Index).Value
        
        If IsEmpty(firstVal) Or Trim(CStr(firstVal)) = "" Then
            ' Table contains only a placeholder row. Exit.
            ' Reset colors just in case
            tblAssociatedData.DataBodyRange.Interior.Color = xlNone
            GoTo CleanExit
        End If
    End If
    
    ' --- 2. DEFINE COLORS ---
    Dim baseColor1 As Long, baseColor2 As Long
    Dim shadeColor1 As Long, shadeColor2 As Long
    
    baseColor1 = RGB(235, 241, 250)   ' Odd Products
    baseColor2 = RGB(250, 243, 233)   ' Even Products
    shadeColor1 = LightenColor(baseColor1, 0.6)
    shadeColor2 = LightenColor(baseColor2, 0.6)
    
    ' --- 3. RESET FORMATTING ---
    tblAssociatedData.DataBodyRange.Interior.Color = xlNone
    If tblAssociatedData.AutoFilter.FilterMode Then tblAssociatedData.AutoFilter.ShowAllData

    ' --- 4. BUILD LOOKUP DICTIONARY (Source Data) ---
    Dim dictProdIndex As Object
    Set dictProdIndex = CreateObject("Scripting.Dictionary")
    
    Dim arrSourceProd As Variant
    Dim arrSourceHelper As Variant
    Dim rngSourceProd As Range, rngSourceHelper As Range
    
    ' Define Ranges
    Set rngSourceProd = tblFinalProducts.ListColumns("Product Number").DataBodyRange
    Set rngSourceHelper = tblFinalProducts.ListColumns(helperColName).DataBodyRange
    
    ' *** FIX: Handle Single Row Case for Source ***
    If rngSourceProd.Cells.Count = 1 Then
        ReDim arrSourceProd(1 To 1, 1 To 1)
        ReDim arrSourceHelper(1 To 1, 1 To 1)
        arrSourceProd(1, 1) = rngSourceProd.Value
        arrSourceHelper(1, 1) = rngSourceHelper.Value
    Else
        arrSourceProd = rngSourceProd.Value
        arrSourceHelper = rngSourceHelper.Value
    End If
    
    Dim i As Long
    Dim pKey As String, pIndex As Variant
    
    For i = 1 To UBound(arrSourceProd, 1)
        pKey = CStr(arrSourceProd(i, 1))
        pIndex = arrSourceHelper(i, 1)
        
        If Not IsEmpty(pIndex) And Not IsError(pIndex) And pKey <> "" Then
            If Not dictProdIndex.exists(pKey) Then
                dictProdIndex.Add pKey, CLng(pIndex)
            End If
        End If
    Next i
    
    ' --- 5. PROCESS TARGET TABLE ---
    Dim arrTargetProd As Variant
    Dim rngTargetProd As Range
    Set rngTargetProd = tblAssociatedData.ListColumns("ProductNumberText").DataBodyRange
    
    ' *** FIX: Handle Single Row Case for Target ***
    If rngTargetProd.Cells.Count = 1 Then
        ReDim arrTargetProd(1 To 1, 1 To 1)
        arrTargetProd(1, 1) = rngTargetProd.Value
    Else
        arrTargetProd = rngTargetProd.Value
    End If
    
    Dim rngC1_Dark As Range, rngC1_Light As Range
    Dim rngC2_Dark As Range, rngC2_Light As Range
    Dim currentRange As Range
    
    Dim lastProd As String: lastProd = ""
    Dim currentRowInBlock As Long: currentRowInBlock = 0
    Dim prodIdx As Long
    Dim isProductOdd As Boolean
    
    For i = 1 To UBound(arrTargetProd, 1)
        pKey = CStr(arrTargetProd(i, 1))
        
        Set currentRange = tblAssociatedData.DataBodyRange.Rows(i)
        
        If dictProdIndex.exists(pKey) Then
            prodIdx = dictProdIndex(pKey)
            
            If pKey <> lastProd Then
                currentRowInBlock = 1
                lastProd = pKey
            Else
                currentRowInBlock = currentRowInBlock + 1
            End If
            
            isProductOdd = (prodIdx Mod 2 <> 0)
            
            ' Add to appropriate Union Range
            If isProductOdd Then
                If currentRowInBlock Mod 2 <> 0 Then
                    If rngC1_Dark Is Nothing Then Set rngC1_Dark = currentRange Else Set rngC1_Dark = Union(rngC1_Dark, currentRange)
                Else
                    If rngC1_Light Is Nothing Then Set rngC1_Light = currentRange Else Set rngC1_Light = Union(rngC1_Light, currentRange)
                End If
            Else
                If currentRowInBlock Mod 2 <> 0 Then
                    If rngC2_Dark Is Nothing Then Set rngC2_Dark = currentRange Else Set rngC2_Dark = Union(rngC2_Dark, currentRange)
                Else
                    If rngC2_Light Is Nothing Then Set rngC2_Light = currentRange Else Set rngC2_Light = Union(rngC2_Light, currentRange)
                End If
            End If
        End If
    Next i
    
    ' --- 6. APPLY COLORS (Bulk Operation) ---
    If Not rngC1_Dark Is Nothing Then rngC1_Dark.Interior.Color = baseColor1
    If Not rngC1_Light Is Nothing Then rngC1_Light.Interior.Color = shadeColor1
    
    If Not rngC2_Dark Is Nothing Then rngC2_Dark.Interior.Color = baseColor2
    If Not rngC2_Light Is Nothing Then rngC2_Light.Interior.Color = shadeColor2

CleanExit:
    application.ScreenUpdating = True
    application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    MsgBox "Error in Product Based Formatting (color coding): " & Err.description, vbInformation
    Resume CleanExit
End Sub

' ===================================================================
' HELPER FUNCTION: LightenColor (Kept exactly as you had it)
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

Public Sub RunProductBasedFormattingBOM()
    Call RunProductBasedFormatting("1. BOM Definition", "BOMDefinition", "Helper Format BOMs")
End Sub

Public Sub RunProductBasedFormattingRouting()
    Call RunProductBasedFormatting("2. Routines", "SelectedRoutines", "Helper Format Routings")
End Sub

Sub SpeedOn()
    application.ScreenUpdating = False
    application.EnableEvents = False
    lCalcSave = application.Calculation
    application.Calculation = xlCalculationManual
End Sub

Sub SpeedOff()
    application.ScreenUpdating = True
    application.EnableEvents = True
    application.Calculation = lCalcSave
    ' Force a recalculation to update the new rows
    If lCalcSave = xlCalculationAutomatic Then application.Calculate
End Sub
