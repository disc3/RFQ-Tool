Attribute VB_Name = "Module2"
Sub TransferData()
    Dim sourceWB As Workbook
    Dim targetWB As Workbook
    Dim sourceWS As Worksheet
    Dim targetWS As Worksheet
    Dim sourceRow As Long, targetRow As Long
    Dim sourceCol As Long, targetCol As Long
    Dim lastSourceRow As Long
    Dim listFound As Boolean
    Dim wb As Workbook
    Dim headerValue As String
    Dim sourceE As Integer
    Dim targetB As Integer
    Dim sourceA As Integer
    Dim targetH As Integer
    Dim incrementE As Integer
    Dim incrementB As Integer
    Dim incrementA As Integer
    Dim incrementH As Integer


' Inicializace promìnné listFound
listFound = False

' Procházení všech otevøených workbookù pro hledání požadovaných listù
For Each wb In application.Workbooks
    On Error Resume Next ' Pokud list neexistuje, pøejde na další
    
    ' Hledání listu "Pøevodní tabulka Routings" v workbooku
    Set sourceWS = wb.Sheets("Pøevodní tabulka Routings")
    If Not sourceWS Is Nothing Then
        Set sourceWB = wb
        listFound = True
    End If
    
    ' Hledání listu "Template_Routing_Connect" v workbooku
    Set targetWS = wb.Sheets("Template_Routing_Connect")
    If Not targetWS Is Nothing Then
        Set targetWB = wb
        listFound = True
    End If
    
    On Error GoTo 0 ' Obnovení bìžné chybové kontroly
Next wb

' Pokud nebyly nalezeny požadované workbooky a listy, ukonèí se makro
If sourceWB Is Nothing Or targetWB Is Nothing Then
    MsgBox "Požadované workbooky nebo listy nejsou otevøeny."
    Exit Sub
End If

' Define initial values for source and target rows and columns
    sourceRow = 14 ' Starting row in the source sheet
    targetRow = 4  ' Starting row in the target sheet
    sourceCol = 1  ' Starting column in the source sheet
    targetCol = 2  ' Starting column in the target sheet
    
    ' Define increments for each iteration
    incrementB = 7   ' Increment for target column B
    incrementE = 20  ' Increment for source column F
    incrementA = 20  ' Increment for source column A
    incrementH = 7   ' Increment for target column H

    ' Start the loop that will run up to 100 times
    i = 1
    sourceE = 1  ' Starting row for column F in the source sheet
    targetB = 4  ' Starting row for column B in the target sheet
    sourceA = 14 ' Starting row for column A in the source sheet
    targetH = 5  ' Starting row for column H in the target sheet

    Do Until i > 100
        ' Copy F1 from source to B4:B10 in target
        targetWS.Range("B" & targetB & ":B" & targetB + 6).value = sourceWS.Range("E" & sourceE).value

        ' Copy A14:M19 from source to H5:T9 in target
        targetWS.Range("H" & targetH & ":T" & targetH + 5).value = sourceWS.Range("A" & sourceA & ":M" & sourceA + 5).value

        ' Copy A34:M39 from source to H12:T17 in target
        targetWS.Range("H" & targetH + incrementH & ":T" & targetH + incrementH + 5).value = sourceWS.Range("A" & sourceA + incrementA & ":M" & sourceA + incrementA + 5).value

        ' Second iteration: copy F21 from source to B11:B17 in target
        targetWS.Range("B" & (targetB + incrementB) & ":B" & (targetB + incrementB + 6)).value = sourceWS.Range("E" & (sourceE + incrementE)).value

        ' Copy A34:M39 from source to H12:T17 in target
        targetWS.Range("H" & targetH + incrementH & ":T" & targetH + incrementH + 5).value = sourceWS.Range("A" & sourceA + incrementA & ":M" & sourceA + incrementA + 5).value

        ' Update the target and source values for the next iteration
        targetB = targetB + incrementB  ' Increment target B
        sourceE = sourceE + incrementE  ' Increment source E
        sourceA = sourceA + incrementA ' Increment source A
        targetH = targetH + incrementH ' Increment target H
        
        ' Increment the counter for the loop
        i = i + 1
    Loop

    ' Save and close workbooks
    ' targetWB.Save

    MsgBox "Data transfer complete!"
End Sub

