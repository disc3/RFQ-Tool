Attribute VB_Name = "FilterComponents"
Option Explicit

' --- KONSTANTEN F�R DIE FILTERFUNKTION ---
Private Const LIST_OBJECT_NAME As String = "LoadedData"
Private Const SEARCH_COLUMN_NAME As String = "SearchColumn"
Private Const SOURCE_SHEET_NAME As String = "Purchasing Info Records"
Private Const PLANT_DATA_COLUMN_NAME As String = "Source"

' --- CONSTANTS FOR PREFERENTIAL ALTERNATIVES ---
Private Const ALT_SHEET_NAME As String = "Material Alternatives"
Private Const ALT_TABLE_NAME As String = "AlternativeMaterials"
Private Const ALT_SOURCE_COL As String = "SourceMaterial"
Private Const ALT_TARGET_COL As String = "AlternativeMaterial"

'##################################################################################
'# NAME:          GetFilteredData
'# ZWECK:         Filtert Daten aus einer Tabelle basierend auf einem Suchbegriff
'#                und einer Werksauswahl und gibt die Ergebnisse als Collection zur�ck.
'# PARAMETER:
'#   userInput:     Der Suchbegriff des Benutzers. Wildcard * wird unterst�tzt.
'#   plantsToInclude: Eine Collection mit den Werksnamen, die ber�cksichtigt werden sollen.
'#                    Wenn die Collection leer ist, werden alle Werke durchsucht.
'# R�CKGABE:      Eine Collection, bei der jedes Element ein Array
'#                (eine Ergebniszeile) ist.
'##################################################################################
Function GetFilteredData(userInput As String, plantsToInclude As Collection, searchColumnName As String, Optional ByRef outIsAlternative As Collection = Nothing) As Collection
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
    On Error GoTo 0 ' Fehlerbehandlung zur�cksetzen

    ' Material column index for alternative lookups
    Dim materialColumnIndex As Long
    materialColumnIndex = lo.ListColumns("Material").Index

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

    ' Daten in ein Array laden f�r maximale Geschwindigkeit
    dataArr = lo.DataBodyRange.Value2

    ' --- 3b. Alternatives Setup (only if mapping table exists and search is not empty) ---
    Dim altDict As Object
    Dim matIndexDict As Object
    Dim useAlternatives As Boolean
    Dim addedKeys As Object
    Dim matchedMaterials As Object

    useAlternatives = False
    If Trim(userInput) <> "" Then
        Set altDict = LoadAlternativesDict()
        If Not altDict Is Nothing Then
            useAlternatives = True
            Set matIndexDict = BuildMaterialRowIndex(dataArr, materialColumnIndex)
            Set outIsAlternative = New Collection
            Set addedKeys = CreateObject("Scripting.Dictionary")
            addedKeys.CompareMode = vbTextCompare
            Set matchedMaterials = CreateObject("Scripting.Dictionary")
            matchedMaterials.CompareMode = vbTextCompare
        End If
    End If

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

            If Not plantMatch Then GoTo NextRow ' Wenn Werk nicht passt, n�chste Zeile
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
            results.Add BuildOutputRow(dataArr, i, numSourceColumns, searchColumnIndex)

            ' Track for alternative resolution
            If useAlternatives Then
                outIsAlternative.Add False ' Direct match, not an alternative

                Dim matVal As String
                matVal = LCase(Trim(CStr(dataArr(i, materialColumnIndex))))

                ' Track Material+Plant for deduplication
                Dim dedupKey As String
                If plantColumnIndex > 0 Then
                    dedupKey = matVal & "|" & LCase(Trim(CStr(dataArr(i, plantColumnIndex))))
                Else
                    dedupKey = matVal & "|"
                End If
                If Not addedKeys.Exists(dedupKey) Then addedKeys.Add dedupKey, True

                ' Track unique materials for post-loop alternative lookup
                If Not matchedMaterials.Exists(matVal) Then matchedMaterials.Add matVal, True
            End If
        End If

NextRow:
    Next i

    ' --- 4b. Post-loop: Resolve alternatives for matched materials ---
    If useAlternatives And matchedMaterials.Count > 0 Then
        Dim matKey As Variant
        Dim altMaterials As Collection
        Dim altMat As Variant
        Dim altRowIndices As Collection
        Dim altRowIdx As Variant
        Dim altPlantMatch As Boolean
        Dim altPlantInRow As String
        Dim altDedupKey As String
        Dim altPlantToCompare As Variant

        For Each matKey In matchedMaterials.Keys
            If altDict.Exists(CStr(matKey)) Then
                Set altMaterials = altDict(CStr(matKey))

                For Each altMat In altMaterials
                    If matIndexDict.Exists(LCase(CStr(altMat))) Then
                        Set altRowIndices = matIndexDict(LCase(CStr(altMat)))

                        For Each altRowIdx In altRowIndices
                            ' Apply plant pre-filter
                            If preFilterActive Then
                                altPlantMatch = False
                                altPlantInRow = Trim(CStr(dataArr(CLng(altRowIdx), plantColumnIndex)))
                                For Each altPlantToCompare In plantsToInclude
                                    If LCase(altPlantInRow) = LCase(CStr(altPlantToCompare)) Then
                                        altPlantMatch = True
                                        Exit For
                                    End If
                                Next altPlantToCompare
                                If Not altPlantMatch Then GoTo NextAltRow
                            End If

                            ' Deduplication: skip if this Material+Plant is already in results
                            altDedupKey = LCase(Trim(CStr(dataArr(CLng(altRowIdx), materialColumnIndex)))) & "|"
                            If plantColumnIndex > 0 Then
                                altDedupKey = altDedupKey & LCase(Trim(CStr(dataArr(CLng(altRowIdx), plantColumnIndex))))
                            End If
                            If addedKeys.Exists(altDedupKey) Then GoTo NextAltRow
                            addedKeys.Add altDedupKey, True

                            ' Build output row and add as alternative
                            results.Add BuildOutputRow(dataArr, CLng(altRowIdx), numSourceColumns, searchColumnIndex)
                            outIsAlternative.Add True ' Mark as alternative

NextAltRow:
                        Next altRowIdx
                    End If
                Next altMat
            End If
        Next matKey
    End If

    ' --- 5. Aufr�umen und Ergebnisse zur�ckgeben ---
    Set GetFilteredData = results ' Das Collection-Objekt wird zur�ckgegeben

    Set regex = Nothing
    Set lo = Nothing
    Set searchColumnObj = Nothing
    Set sourceSheet = Nothing
    Set altDict = Nothing
    Set matIndexDict = Nothing
    Set addedKeys = Nothing
    Set matchedMaterials = Nothing
    application.ScreenUpdating = True
End Function

'----------------------------------------------------------------------------------
' HELPER: Loads the AlternativeMaterials mapping table into a Dictionary.
' Returns Nothing if the table does not exist or is empty (feature disabled).
' Key: LCase(SourceMaterial), Value: Collection of AlternativeMaterial strings.
'----------------------------------------------------------------------------------
Private Function LoadAlternativesDict() As Object
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim dict As Object
    Dim altArr As Variant
    Dim i As Long
    Dim srcKey As String, altVal As String
    Dim srcColIdx As Long, altColIdx As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(ALT_SHEET_NAME)
    If ws Is Nothing Then
        Set LoadAlternativesDict = Nothing
        Exit Function
    End If
    Set lo = ws.ListObjects(ALT_TABLE_NAME)
    If lo Is Nothing Then
        Set LoadAlternativesDict = Nothing
        Exit Function
    End If
    On Error GoTo 0

    If lo.ListRows.Count = 0 Or lo.DataBodyRange Is Nothing Then
        Set LoadAlternativesDict = Nothing
        Exit Function
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare ' Case-insensitive keys

    srcColIdx = lo.ListColumns(ALT_SOURCE_COL).Index
    altColIdx = lo.ListColumns(ALT_TARGET_COL).Index
    altArr = lo.DataBodyRange.Value2

    For i = 1 To UBound(altArr, 1)
        srcKey = Trim(CStr(altArr(i, srcColIdx)))
        altVal = Trim(CStr(altArr(i, altColIdx)))

        If srcKey <> "" And altVal <> "" Then
            If Not dict.Exists(srcKey) Then
                dict.Add srcKey, New Collection
            End If
            dict(srcKey).Add altVal
        End If
    Next i

    If dict.Count = 0 Then
        Set LoadAlternativesDict = Nothing
    Else
        Set LoadAlternativesDict = dict
    End If
End Function

'----------------------------------------------------------------------------------
' HELPER: Builds a Dictionary mapping LCase(Material) -> Collection of row indices
' in the data array. Single O(N) pass for fast alternative row lookups.
'----------------------------------------------------------------------------------
Private Function BuildMaterialRowIndex(dataArr As Variant, materialColIdx As Long) As Object
    Dim dict As Object
    Dim i As Long
    Dim matKey As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For i = 1 To UBound(dataArr, 1)
        matKey = Trim(CStr(dataArr(i, materialColIdx)))
        If matKey <> "" Then
            If Not dict.Exists(matKey) Then
                dict.Add matKey, New Collection
            End If
            dict(matKey).Add i
        End If
    Next i

    Set BuildMaterialRowIndex = dict
End Function

'----------------------------------------------------------------------------------
' HELPER: Builds a 1D output array for a single row, excluding the SearchColumn.
' Protects against Excel formula injection.
'----------------------------------------------------------------------------------
Private Function BuildOutputRow(dataArr As Variant, rowIndex As Long, numSourceColumns As Long, searchColumnIndex As Long) As Variant
    Dim tempRowValues As Variant
    Dim outputColIdx As Long
    Dim c As Long
    Dim valueToWrite As Variant

    ReDim tempRowValues(1 To numSourceColumns - 1)
    outputColIdx = 0

    For c = 1 To numSourceColumns
        If c <> searchColumnIndex Then
            outputColIdx = outputColIdx + 1
            valueToWrite = dataArr(rowIndex, c)
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

    BuildOutputRow = tempRowValues
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
    ' Deaktiviert AutoSave, um Popups w�hrend der Aktualisierung zu vermeiden
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
' HELPER: Verwaltet die AutoSave-Einstellung f�r Cloud-Dateien
'----------------------------------------------------------------------------------
Public Sub SetCloudAutoSave(enableAutoSave As Boolean)
    On Error Resume Next ' Falls nicht in der Cloud gespeichert
    If LCase(Left(ActiveWorkbook.FullName, 4)) = "http" Then
        ActiveWorkbook.AutoSaveOn = enableAutoSave
    End If
    On Error GoTo 0
End Sub

