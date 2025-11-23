Attribute VB_Name = "ServoCalculation"
' --- OPTIMIERTES HAUPTMODUL ---

' Behält die ursprünglichen Mengen für die Routinen bei
Dim originalRoutineQuantities() As Variant
Dim routineRowIndexes() As Long

Public Sub RunServoCalculation()
    Dim startTime As Double
    startTime = Timer ' Zeitmessung starten

    ' --- Optimierungs-Boilerplate ---
    ' Excel-Funktionen deaktivieren, um die Geschwindigkeit zu maximieren
    application.ScreenUpdating = False
    application.EnableEvents = False
    application.Calculation = xlCalculationManual

    'On Error GoTo ErrorHandler ' Stellt sicher, dass die Einstellungen auch bei einem Fehler zurückgesetzt werden

    Dim wsBom As Worksheet, wsSales As Worksheet, wsServo As Worksheet
    Set wsBom = ThisWorkbook.Sheets("1. BOM Definition")
    Set wsSales = ThisWorkbook.Sheets("4. Sales Calculation (Internal)")
    Set wsServo = ThisWorkbook.Sheets("8. Servo calculation")
    
    Dim tblBOM As ListObject
    Set tblBOM = wsBom.ListObjects("BOMDefinition")

    Dim selectedProduct As String
    selectedProduct = Trim(wsBom.Range("F11").Value)
    If selectedProduct = "" Then
        MsgBox "Please select a product in cell F11 first.", vbExclamation
        
        GoTo CleanExit
    End If

    ' --- Kabelauswahl (UserForm bleibt gleich, aber der Aufruf ist hier) ---
    Dim cableForm As ServoCableSelectForm
    Set cableForm = New ServoCableSelectForm
    cableForm.Show
    
    Dim cableMaterial As String
    cableMaterial = cableForm.SelectedCableMaterial
    Unload cableForm
    
    If cableMaterial = "" Then
        MsgBox "No cable component selected.", vbExclamation
        GoTo CleanExit
    End If
    
    ' --- Daten in Arrays einlesen (Einmaliger Lesevorgang) ---
    Dim bomData As Variant
    bomData = tblBOM.DataBodyRange.Value ' Gesamte Tabelle in ein Array lesen

    ' Spaltenindizes finden, um flexibel zu bleiben
    Dim prodCol As Long, matCol As Long, qtyCol As Long, copperCol As Long, netWeightCol As Long, descCol As Long, erpCol As Long
    prodCol = tblBOM.ListColumns("Product Number").Index
    matCol = tblBOM.ListColumns("Material").Index
    qtyCol = tblBOM.ListColumns("Quantity").Index
    copperCol = tblBOM.ListColumns("Copper weight [kg/1000m]").Index
    netWeightCol = tblBOM.ListColumns("Net weight (kg/Base unit)").Index
    descCol = tblBOM.ListColumns("Product Description").Index
    erpCol = tblBOM.ListColumns("ERP Part Number").Index
    
    ' --- Benötigte Daten in einer einzigen Schleife sammeln ---
    Dim i As Variant
    Dim originalQuantities As Object ' Dictionary zum Speichern der Originalwerte
    Set originalQuantities = CreateObject("Scripting.Dictionary")
    
    Dim cableRowIndex As Long, otherComponentWeight As Double
    Dim copperWeightPerKm As Double, netWeightPerM As Double
    Dim desc As String, erpPN As String

    For i = 1 To UBound(bomData, 1)
        If Trim(bomData(i, prodCol)) = selectedProduct Then
            ' Originalmenge speichern
            originalQuantities(i) = bomData(i, qtyCol)
            
            If Trim(bomData(i, matCol)) = cableMaterial Then
                ' Daten für das ausgewählte Kabel merken
                cableRowIndex = i
                copperWeightPerKm = CDbl(bomData(i, copperCol))
                netWeightPerM = CDbl(bomData(i, netWeightCol))
                desc = bomData(i, descCol)
                erpPN = bomData(i, erpCol)
            Else
                ' Gewicht der anderen Komponenten aufsummieren
                otherComponentWeight = otherComponentWeight + CDbl(bomData(i, netWeightCol))
            End If
        End If
    Next i
    
    If cableRowIndex = 0 Then
        MsgBox "Selected cable component could not be found in the BOM.", vbCritical
        GoTo CleanExit
    End If

    ' --- SIMULATION 1: 0.1m Kabel + Arbeit ---
    ' Nur die eine benötigte Zelle ändern
    tblBOM.DataBodyRange.Cells(cableRowIndex, qtyCol).Value = 0.1
    application.Calculate ' Neuberechnung durchführen
    Dim labConPrice As Double
    labConPrice = GetPriceFromSales(wsSales, selectedProduct, "Selling Price/ Transfer Price Copper Base / 1 piece")

    ' --- SIMULATION 2: 1m Kabel only ---
    ' Andere Mengen auf 0 setzen
    Call ZeroRoutineQuantities(selectedProduct)
    For Each i In originalQuantities.Keys
        If i <> cableRowIndex Then
            tblBOM.DataBodyRange.Cells(i, qtyCol).Value = 0
        End If
    Next i
    tblBOM.DataBodyRange.Cells(cableRowIndex, qtyCol).Value = 1
    
    application.Calculate ' Neuberechnung durchführen
    Dim cableCostPerM As Double
    cableCostPerM = GetPriceFromSales(wsSales, selectedProduct, "Selling Price/ Transfer Price Copper Base / 1 piece")

    ' --- Wiederherstellung ---
    Call RestoreRoutineQuantities
    For Each i In originalQuantities.Keys
        tblBOM.DataBodyRange.Cells(i, qtyCol).Value = originalQuantities(i)
    Next i
    application.Calculate ' Endgültigen Zustand wiederherstellen

    ' --- Output-Daten in einem Array vorbereiten ---
    wsServo.Range("A4:I9000,K4:R9000").ClearContents ' Bereiche auf einmal löschen
    
    Dim outputData() As Variant
    ReDim outputData(1 To 200, 1 To 18) ' Array für 200 Zeilen und 18 Spalten
    
    Dim outRow As Long: outRow = 1
    Dim lengthVal As Double
    
    For lengthVal = 0.5 To 100 Step 0.5
        ' Daten ins Array schreiben (viel schneller als auf das Blatt)
        outputData(outRow, 1) = selectedProduct
        outputData(outRow, 2) = desc
        outputData(outRow, 3) = erpPN
        outputData(outRow, 4) = desc
        outputData(outRow, 5) = lengthVal
        outputData(outRow, 6) = labConPrice + (lengthVal * cableCostPerM)
        outputData(outRow, 7) = lengthVal * copperWeightPerKm
        outputData(outRow, 8) = 150
        outputData(outRow, 9) = lengthVal * netWeightPerM + otherComponentWeight

        outputData(outRow, 11) = erpPN
        outputData(outRow, 12) = desc
        outputData(outRow, 13) = labConPrice
        outputData(outRow, 14) = cableCostPerM
        outputData(outRow, 15) = cableMaterial
        outputData(outRow, 16) = netWeightPerM
        outputData(outRow, 17) = otherComponentWeight
        outputData(outRow, 18) = copperWeightPerKm
        
        outRow = outRow + 1
    Next lengthVal

    ' --- Array in einer einzigen Operation auf das Blatt schreiben ---
    wsServo.Range("A4").Resize(UBound(outputData, 1), 9).Value = outputData
    ' Da die Spalten nicht zusammenhängend sind, schreiben wir den zweiten Teil separat
    ' Dafür müssen wir einen Trick anwenden und die Daten in ein zweites Array kopieren
    Dim outputDataPart2() As Variant
    ReDim outputDataPart2(1 To 200, 1 To 8)
    For i = 1 To 200
        outputDataPart2(i, 1) = outputData(i, 11)
        outputDataPart2(i, 2) = outputData(i, 12)
        outputDataPart2(i, 3) = outputData(i, 13)
        outputDataPart2(i, 4) = outputData(i, 14)
        outputDataPart2(i, 5) = outputData(i, 15)
        outputDataPart2(i, 6) = outputData(i, 16)
        outputDataPart2(i, 7) = outputData(i, 17)
        outputDataPart2(i, 8) = outputData(i, 18)
    Next i
    wsServo.Range("K4").Resize(UBound(outputDataPart2, 1), UBound(outputDataPart2, 2)).Value = outputDataPart2
    
    wsServo.Activate
    
    Debug.Print "Makro-Laufzeit: " & Timer - startTime & " Sekunden"

CleanExit:
    ' --- Aufräumen ---
    ' Excel-Einstellungen wiederherstellen
    application.ScreenUpdating = True
    application.EnableEvents = True
    application.Calculation = xlCalculationAutomatic
    Exit Sub

ErrorHandler:
    ' Fehlerbehandlung: Stellt sicher, dass die Einstellungen immer zurückgesetzt werden
    MsgBox "Ein Fehler ist aufgetreten: " & vbCrLf & Err.description
    Resume CleanExit
End Sub

' --- OPTIMIERTE GetPriceFromSales FUNKTION ---
Function GetPriceFromSales(wsSales As Worksheet, product As String, columnHeader As String) As Double
    Dim headerRowRange As Range: Set headerRowRange = wsSales.Rows(14)
    Dim productColRange As Range: Set productColRange = wsSales.Range("B15:B2000")
    
    Dim colIndex As Variant, rowIndex As Variant
    
    ' Finde die Spalte mit Application.Match (viel schneller als eine Schleife)
    colIndex = application.Match(columnHeader, headerRowRange, 0)
    
    If IsError(colIndex) Then
        MsgBox "Column '" & columnHeader & "' not found.", vbCritical
        GetPriceFromSales = 0
        Exit Function
    End If
    
    ' Finde die Zeile mit Application.Match
    rowIndex = application.Match(product, productColRange, 0)
    
    If IsError(rowIndex) Then
        MsgBox "Product '" & product & "' not found in Sales Calculation sheet.", vbExclamation
        GetPriceFromSales = 0
        Exit Function
    End If
    Debug.Print wsSales.Cells(14 + rowIndex, colIndex).Value
    ' Wert direkt aus der Zelle auslesen
    GetPriceFromSales = CDbl(wsSales.Cells(14 + rowIndex, colIndex).Value)
End Function

Sub ZeroRoutineQuantities(product As String)
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long, n As Long
    Set ws = ThisWorkbook.Sheets("2. Routines")
    Set tbl = ws.ListObjects("SelectedRoutines")

    Dim qtyCol As Long
    On Error Resume Next
    qtyCol = tbl.ListColumns("Number of operations").Index
    On Error GoTo 0
    If qtyCol = 0 Then
        MsgBox "Cannot find 'Number of operations' column in SelectedRoutines.", vbCritical
        Exit Sub
    End If

    n = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(tbl.DataBodyRange.Cells(i, tbl.ListColumns("Product Number").Index).Value) = product Then
            n = n + 1
        End If
    Next i

    ReDim originalRoutineQuantities(1 To n)
    ReDim routineRowIndexes(1 To n)

    n = 0
    For i = 1 To tbl.DataBodyRange.Rows.Count
        If Trim(tbl.DataBodyRange.Cells(i, tbl.ListColumns("Product Number").Index).Value) = product Then
            n = n + 1
            routineRowIndexes(n) = i
            originalRoutineQuantities(n) = tbl.DataBodyRange.Cells(i, qtyCol).Value
            tbl.DataBodyRange.Cells(i, qtyCol).Value = 0
        End If
    Next i
End Sub

Sub RestoreRoutineQuantities()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim i As Long
    Set ws = ThisWorkbook.Sheets("2. Routines")
    Set tbl = ws.ListObjects("SelectedRoutines")

    If Not IsEmpty(originalRoutineQuantities) Then
        For i = LBound(routineRowIndexes) To UBound(routineRowIndexes)
            tbl.DataBodyRange.Cells(routineRowIndexes(i), tbl.ListColumns("Number of operations").Index).Value = originalRoutineQuantities(i)
        Next i
    End If
End Sub


