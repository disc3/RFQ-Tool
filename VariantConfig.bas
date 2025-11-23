Attribute VB_Name = "VariantConfig"

Sub CreateVariants()

    
    ' Clear public variables to ensure no stale data
    selectedMacrophase = vbNullString
    selectedMicrophase = vbNullString
    selectedMaterial = vbNullString
    selectedMachine = vbNullString
    selectedOperations = 0

    VariantCreationForm.Show
End Sub

Sub CreateVariants_V2()
' --- 1. PRE-FLIGHT CHECK ---
    ' First, verify that the necessary data exists before even trying to load the form.
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim isTableEmpty As Boolean
    
    On Error Resume Next ' In case sheet or table doesn't exist
    Set ws = ThisWorkbook.Sheets("1. BOM Definition")
    Set tbl = ws.ListObjects("BOMDefinition")
    On Error GoTo 0
    
    If tbl Is Nothing Then
        MsgBox "The 'BOMDefinition' table could not be found. The tool cannot be started.", vbCritical, "Prerequisite Missing"
        Exit Sub
    End If
    
    If tbl.DataBodyRange Is Nothing Then
        isTableEmpty = True
    ElseIf tbl.ListRows.Count = 1 And _
           Len(Trim(CStr(tbl.ListColumns("Product Number").DataBodyRange.Cells(1).Value))) = 0 Then
        isTableEmpty = True
    End If
    
    If isTableEmpty Then
        MsgBox "The BOM Definition table is empty." & vbCrLf & vbCrLf & _
               "Please add at least one base product before creating a variant.", _
               vbExclamation, "No Base BOM Data"
        Exit Sub ' Stop right here.
    End If

    ' --- 2. LAUNCH THE FORM ---
    ' If we get this far, we know the data is valid.
    ' We no longer need the complex "WasCancelled" check.
    frmVariantConfigurator.Show
End Sub

Sub GenerateWorkCenterSummary()
    Dim wsData As Worksheet, wsOut As Worksheet
    Dim tbl As ListObject
    Dim lastRow As Long, outRow As Long
    Dim dict As Object, key As Variant
    Dim i As Variant
    Dim prodNum As Variant, variantOf As String, macrophase As String
    Dim tr As Double, te As Double, batch As Double
    Dim summaryKey As String
    Dim prodDict As Object

    Set wsData = ThisWorkbook.Sheets("2. Routines")
    Set tbl = wsData.ListObjects("SelectedRoutines")

    ' Create output sheet
    On Error Resume Next
    application.DisplayAlerts = False
    ThisWorkbook.Sheets("RoutineSummary").Delete
    application.DisplayAlerts = True
    On Error GoTo 0
    Set wsOut = ThisWorkbook.Sheets.Add
    wsOut.name = "RoutineSummary"

    ' Output headers
    wsOut.Range("A1:F1").Value = Array("Product Number", "Macrophase", "Sum Tr", "Sum Te", "TR / Piece", "TE / Piece")
    outRow = 2

    Set dict = CreateObject("Scripting.Dictionary")

    ' First, build a map of each product's rows
    Set prodDict = CreateObject("Scripting.Dictionary")
    For i = 1 To tbl.ListRows.Count
        prodNum = tbl.DataBodyRange(i, tbl.ListColumns("Product Number").Index).Value
        If Not prodDict.exists(prodNum) Then
            Set prodDict(prodNum) = New Collection
        End If
        prodDict(prodNum).Add i ' Store row index
    Next i

    ' Process all product numbers individually
    For Each prodNum In prodDict.Keys
        ' Check if it's a variant and get base
        variantOf = ""
        i = prodDict(prodNum).item(1)
        On Error Resume Next
        variantOf = tbl.DataBodyRange(i, tbl.ListColumns("Variant of").Index).Value
        On Error GoTo 0

        ' Create a collection of relevant rows: base + self
        Dim rowIndexes As Collection
        Set rowIndexes = New Collection

        ' Add own rows
        For Each i In prodDict(prodNum)
            rowIndexes.Add i
        Next i

        ' Add base rows if variant
        If variantOf <> "" And prodDict.exists(variantOf) Then
            For Each i In prodDict(variantOf)
                rowIndexes.Add i
            Next i
        End If

        ' Group by Macrophase
        Dim macDict As Object
        Set macDict = CreateObject("Scripting.Dictionary")
        
        Dim trSum As Double, teSum As Double, batchSize As Double
        
        For Each i In rowIndexes
            macrophase = tbl.ListColumns("Macrophase").DataBodyRange.Cells(i, 1).Value
            tr = tbl.ListColumns("Total Tr").DataBodyRange.Cells(i, 1).Value
            te = tbl.ListColumns("Total Te").DataBodyRange.Cells(i, 1).Value
            batch = tbl.ListColumns("Batch").DataBodyRange.Cells(i, 1).Value
        
            If Not macDict.exists(macrophase) Then
                macDict(macrophase) = Array(0#, 0#, batch)
            End If
        
            trSum = macDict(macrophase)(0) + tr
            teSum = macDict(macrophase)(1) + te
            batchSize = macDict(macrophase)(2) ' assume consistent batch size
        
            macDict(macrophase) = Array(trSum, teSum, batchSize)
        Next i


        ' Output to sheet
        For Each key In macDict.Keys
            trSum = macDict(key)(0)
            teSum = macDict(key)(1)
            batchSize = macDict(key)(2)
        
            wsOut.Cells(outRow, 1).Value = prodNum
            wsOut.Cells(outRow, 2).Value = key
            wsOut.Cells(outRow, 3).Value = trSum
            wsOut.Cells(outRow, 4).Value = teSum
        
            If batchSize <> 0 Then
                wsOut.Cells(outRow, 5).Value = trSum / batchSize
                wsOut.Cells(outRow, 6).Value = teSum / batchSize
            End If
        
            outRow = outRow + 1
        Next key

    Next prodNum

    MsgBox "Routine summary by work center generated on 'RoutineSummary' sheet.", vbInformation
End Sub

