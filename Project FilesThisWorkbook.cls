VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
    ClearTables.ClearComponentDatabase True
End Sub

Private Sub Workbook_Open()
    Dim qry As WorkbookQuery
    Dim excludeQueries As Object
    Dim queryName As String

    ' Optional: Beschleunigt das Öffnen der Datei
    application.Calculation = xlCalculationManual
    
    ThisWorkbook.Sheets("Purchasing Info Records").Visible = xlVeryHidden
    ' Erstellt ein Dictionary-Objekt, um die Namen der auszuschließenden Abfragen zu speichern.
    ' Das Dictionary ermöglicht eine schnelle Überprüfung, ob eine Abfrage ausgeschlossen werden soll.
    Set excludeQueries = CreateObject("Scripting.Dictionary")
    
    ' Hier fügst du die Namen der Abfragen hinzu, die NICHT aktualisiert werden sollen.
    ' Achte auf die genaue Schreibweise der Abfragenamen.
    excludeQueries("LoadedData") = True
    excludeQueries("GetExtendedMMD") = True
    ' Füge hier weitere Abfragen hinzu, falls nötig:
    ' excludeQueries("NochEineAbfrage") = True
    
    
    ThisWorkbook.Queries.FastCombine = True
    On Error Resume Next ' Fehlerbehandlung, falls eine Abfrage nicht aktualisiert werden kann

    ' Durchläuft alle Abfragen in der aktuellen Arbeitsmappe
    For Each qry In ThisWorkbook.Queries
        queryName = qry.name ' Holt den Namen der aktuellen Abfrage
        
        ' Überprüft, ob die aktuelle Abfrage NICHT im Dictionary der auszuschließenden Abfragen enthalten ist.
        If Not excludeQueries.exists(queryName) Then
            ' Wenn die Abfrage nicht ausgeschlossen werden soll, wird sie aktualisiert.
            qry.Refresh
        Else
            ' Optional: Gib eine Meldung aus, dass die Abfrage übersprungen wird (kann für Debugging nützlich sein).
            ' Debug.Print "Abfrage übersprungen: " & queryName
        End If
    Next qry

    On Error GoTo 0 ' Setzt die Fehlerbehandlung zurück

    application.Calculation = xlCalculationAutomatic ' Setzt die Berechnungsmethode zurück
End Sub

