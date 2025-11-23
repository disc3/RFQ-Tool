Attribute VB_Name = "FilterComponents"
Option Explicit

' --- KONSTANTEN FÜR DIE FILTERFUNKTION ---
Private Const LIST_OBJECT_NAME As String = "LoadedData"
Private Const SEARCH_COLUMN_NAME As String = "SearchColumn"
Private Const SOURCE_SHEET_NAME As String = "Purchasing Info Records"
Private Const PLANT_DATA_COLUMN_NAME As String = "Source"

'##################################################################################
'# NAME:          GetFilteredData
'# ZWECK:         Filtert Daten aus einer Tabelle basierend auf einem Suchbegriff
'#                und einer Werksauswahl und gibt die Ergebnisse als Collection zurück.
'# PARAMETER:
'#   userInput:     Der Suchbegriff des Benutzers. Wildcard * wird unterstützt.
'#   plantsToInclude: Eine Collection mit den Werksnamen, die berücksichtigt werden sollen.
'#                    Wenn die Collection leer ist, werden alle Werke durchsucht.
'# RÜCKGABE:      Eine Collection, bei der jedes Element ein Array
'#                (eine Ergebniszeile) ist.
'##################################################################################
Function GetFilteredData(userInput As String, plantsToInclude As Collection, searchColumnName As String) As Collection
    Dim lo As ListObject
    Dim sourceSheet As Worksheet
    Dim searchColumnObj As ListColumn
    Dim dataArr As Variant
    Dim results As Collection
    Dim i As Long, c As Long
    Dim regexPattern As String
    Dim regex As Object
    Dim tempRowValues As Variant
    Dim numSourceColumns As Long
    Dim searchColumnIndex As Long
    Dim plantColumnIndex As Long
    Dim preFilterActive As Boolean

    application.ScreenUpdating = False

    ' --- 1. Quelldaten lokalisieren ---
    On Error Resume Next
    Set sourceSheet = ThisWorkbook.Sheets(SOURCE_SHEET_NAME)
    If sourceSheet Is Nothing Then
        MsgBox "Quell-Tabellenblatt '" & SOURCE_SHEET_NAME & "' nicht gefunden.", vbCritical
        Exit Function
    End If

    Set lo = sourceSheet.ListObjects(LIST_OBJECT_NAME)
    If lo Is Nothing Then
        MsgBox "Tabelle (ListObject) '" & LIST_OBJECT_NAME & "' nicht auf Blatt '" & SOURCE_SHEET_NAME & "' gefunden.", vbCritical
        Exit Function
    End If
    On Error GoTo 0

    If lo.ListRows.Count = 0 Or lo.DataBodyRange Is Nothing Then
        ' Leise beenden, die UserForm kann eine Nachricht anzeigen
        Exit Function
    End If
    
    ' --- 2. Spaltenindizes bestimmen ---
    numSourceColumns = lo.ListColumns.Count
    
    On Error Resume Next
    Set searchColumnObj = lo.ListColumns(SEARCH_COLUMN_NAME)
    If searchColumnObj Is Nothing Then
        MsgBox "Spalte '" & SEARCH_COLUMN_NAME & "' in Tabelle '" & LIST_OBJECT_NAME & "' nicht gefunden.", vbExclamation
        Exit Function
    End If
    searchColumnIndex = searchColumnObj.Index
    
    ' Index der Werksspalte finden
    plantColumnIndex = 0
    plantColumnIndex = lo.ListColumns(searchColumnName).Index
    On Error GoTo 0 ' Fehlerbehandlung zurücksetzen

    ' --- 3. Filter-Setup ---
    Set results = New Collection
    preFilterActive = (plantsToInclude.Count > 0 And plantColumnIndex > 0)

    ' Regex-Vorbereitung
    regexPattern = BuildRegexPatternForSearch(userInput)
    Set regex = CreateObject("VBScript.RegExp")
    With regex
        .pattern = regexPattern
        .Global = False
        .IgnoreCase = True
    End With

    ' Daten in ein Array laden für maximale Geschwindigkeit
    dataArr = lo.DataBodyRange.Value2

    ' --- 4. Daten filtern ---
    Dim currentCellData As Variant
    Dim stringToTest As String
    Dim currentPlantInRow As String
    Dim plantMatch As Boolean

    For i = 1 To UBound(dataArr, 1) ' Zeilen durchlaufen
        
        ' --- Werks-Vorfilterung anwenden ---
        If preFilterActive Then
            plantMatch = False
            currentPlantInRow = Trim(CStr(dataArr(i, plantColumnIndex)))
            Dim plantToCompare As Variant
            For Each plantToCompare In plantsToInclude
                If LCase(currentPlantInRow) = LCase(CStr(plantToCompare)) Then
                    plantMatch = True
                    Exit For ' Werk gefunden, innere Schleife verlassen
                End If
            Next plantToCompare
            
            If Not plantMatch Then GoTo NextRow ' Wenn Werk nicht passt, nächste Zeile
        End If

        ' --- Haupt-Filterung mit Regex anwenden ---
        currentCellData = dataArr(i, searchColumnIndex)
        If IsError(currentCellData) Then
            stringToTest = "[Error]"
        ElseIf IsNull(currentCellData) Or IsEmpty(currentCellData) Or Trim(CStr(currentCellData)) = "" Then
            stringToTest = ""
        Else
            stringToTest = CStr(currentCellData)
        End If

        If regex.Test(stringToTest) Then
            ' Ganze Zeile zum Ergebnis hinzufügen (ohne "SearchColumn")
            ReDim tempRowValues(1 To numSourceColumns - 1)
            Dim outputColIdx As Long
            outputColIdx = 0
            
            For c = 1 To numSourceColumns
                ' Die Suchspalte wird nicht in die Ausgabe übernommen
                If c <> searchColumnIndex Then
                    outputColIdx = outputColIdx + 1
                    Dim valueToWrite As Variant
                    valueToWrite = dataArr(i, c)
                    
                    ' Schutz vor Excel-Formel-Interpretation
                    If VarType(valueToWrite) = vbString And Len(valueToWrite) > 0 Then
                        If Left(valueToWrite, 1) = "=" Or Left(valueToWrite, 1) = "+" Or Left(valueToWrite, 1) = "-" Or Left(valueToWrite, 1) = "@" Then
                            tempRowValues(outputColIdx) = "'" & valueToWrite
                        Else
                            tempRowValues(outputColIdx) = valueToWrite
                        End If
                    Else
                        tempRowValues(outputColIdx) = valueToWrite
                    End If
                End If
            Next c
            results.Add tempRowValues
        End If
        
NextRow:
    Next i

    ' --- 5. Aufräumen und Ergebnisse zurückgeben ---
    Set GetFilteredData = results ' Das Collection-Objekt wird zurückgegeben

    Set regex = Nothing
    Set lo = Nothing
    Set searchColumnObj = Nothing
    Set sourceSheet = Nothing
    application.ScreenUpdating = True
End Function

'----------------------------------------------------------------------------------
' HELPER: Baut ein RegEx-Muster aus einer Benutzereingabe mit Wildcard *
'----------------------------------------------------------------------------------
Private Function BuildRegexPatternForSearch(userInput As String) As String
    If Trim(userInput) = "" Then
        BuildRegexPatternForSearch = ".*?" ' Alles finden, wenn die Suche leer ist
        Exit Function
    End If
    
    Dim pattern As String
    Dim i As Long
    Dim char As String
    pattern = ""
    For i = 1 To Len(userInput)
        char = Mid(userInput, i, 1)
        Select Case char
            Case "*"
                pattern = pattern & ".*?"
            Case ".", "\", "+", "?", "[", "]", "^", "$", "(", ")", "{", "}", "|"
                pattern = pattern & "\" & char
            Case Else
                pattern = pattern & char
        End Select
    Next i
    BuildRegexPatternForSearch = pattern
End Function

'----------------------------------------------------------------------------------
' HELPER: Stellt sicher, dass die Quelltabelle aktuell ist
'----------------------------------------------------------------------------------
Public Sub LoadDatabase(Optional showMessageInStatusBar As Boolean = False)
    ' Deaktiviert AutoSave, um Popups während der Aktualisierung zu vermeiden
    Dim wsSource As Worksheet
    Dim tbl As ListObject
    ' Set the worksheet and table
    
    Set wsSource = ThisWorkbook.Sheets("Purchasing Info Records")     ' Replace with the actual sheet name
    Set tbl = wsSource.ListObjects("LoadedData")       ' Replace with the actual table name
    SetCloudAutoSave False
    ' Delete all rows in the DataBodyRange if the table has rows
    If Not tbl.DataBodyRange Is Nothing Then
        application.Calculation = xlCalculationManual
        If showMessageInStatusBar = True Then
            application.StatusBar = "Loading material master database. Please wait... (auto-save will be turned off until you close the file or you unload the database)"
        End If
        With tbl
            .Range.AutoFilter
            If .ListRows.Count = 1 Then
                .QueryTable.Refresh BackgroundQuery:=False
            End If
        End With
        
        If showMessageInStatusBar = True Then
            application.StatusBar = ""
        End If
        application.Calculation = xlCalculationAutomatic
    End If
    
End Sub

'----------------------------------------------------------------------------------
' HELPER: Verwaltet die AutoSave-Einstellung für Cloud-Dateien
'----------------------------------------------------------------------------------
Public Sub SetCloudAutoSave(enableAutoSave As Boolean)
    On Error Resume Next ' Falls nicht in der Cloud gespeichert
    If LCase(Left(ActiveWorkbook.FullName, 4)) = "http" Then
        ActiveWorkbook.AutoSaveOn = enableAutoSave
    End If
    On Error GoTo 0
End Sub

