VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCustomerView 
   Caption         =   "Search Customer List for surcharge bonus"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17655
   OleObjectBlob   =   "frmCustomerView.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmCustomerView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ####################################################################
' # Code für die UserForm: frmCustomerView
' # (Version 2 - Mit Such-Button und Spalten-Filter C-F)
' ####################################################################

' --- Modul-Variablen ---
Private m_vData As Variant
Private m_ws As Worksheet
' HINWEIS: Auf 10 Spalten (A-J) gesetzt
Const M_LCOLCOUNT As Long = 10





Private Sub lblCustName_Click()

End Sub

Private Sub lblTPDiscount_Click()

End Sub

' --- Initialisierungs-Ereignis ---
Private Sub UserForm_Initialize()
    On Error Resume Next
    Set m_ws = ThisWorkbook.Sheets("customer list")
    On Error GoTo 0
    
    If m_ws Is Nothing Then
        MsgBox "Fehler: Arbeitsblatt 'customer list' nicht gefunden.", vbCritical
        Unload Me
        Exit Sub
    End If
    
    ' Lade die Daten aus dem Blatt in unsere Array-Variable
    Call LoadDataFromSheet
    
    ' Fülle die ListBox zum ersten Mal mit *allen* Daten
    ' (ApplyFilter mit leerem Text zeigt alles an)
    Call ApplyFilter("")
End Sub

' --- Hilfsprozedur: Daten laden ---
Private Sub LoadDataFromSheet()
    Dim lLastRow As Long
    
    lLastRow = m_ws.Cells(m_ws.Rows.Count, "A").End(xlUp).row
    
    If lLastRow <= 1 Then
        Exit Sub
    End If
    
    ' HINWEIS: Der Bereich wurde auf "J" erweitert
    m_vData = m_ws.Range("A2:J" & lLastRow).Value2
    
    ' *** FEHLERBEHEBUNG FÜR LAUFZEITFEHLER 9 ***
    ' Wenn nur eine Zeile Daten vorhanden ist (lLastRow = 2),
    ' gibt .Value2 ein 2D-Array (1 To 1, 1 To 10) zurück.
    ' Der alte Code hat hier fälschlicherweise ein 1D-Array erwartet.
    ' Der 'If lLastRow = 2'-Block wurde entfernt, da er
    ' bei einem mehrspaltigen Bereich nicht mehr benötigt wird.
    ' m_vData ist jetzt *immer* ein 2D-Array, solange lLastRow > 1 ist.
    
End Sub

' --- EREIGNIS: Such-Button wird geklickt ---
' NEU: Diese Prozedur wird ausgeführt, wenn du auf "Suchen" klickst.
Private Sub cmdSearch_Click()
    ' Rufe die Filter-Logik mit dem aktuellen Suchtext auf
    Call ApplyFilter(Me.txtSearch.Text)
End Sub

' --- EREIGNIS: Echtzeit-Suche (ENTFERNT) ---
' Die Prozedur 'Private Sub txtSearch_Change()' wurde
' absichtlich gelöscht, um die Echtzeit-Suche zu deaktivieren.

' --- KERNLOGIK: Filtern und Anzeigen ---
Private Sub ApplyFilter(ByVal sFilterTerm As String)
    Dim vFiltered() As Variant
    Dim lMatchCount As Long
    Dim i As Long, j As Long
    Dim bRowAdded As Boolean
    
    ' Suchbegriff für den Vergleich in Kleinbuchstaben umwandeln
    sFilterTerm = LCase(sFilterTerm)
    
    ' Prüfen, ob überhaupt Daten geladen wurden
    If IsEmpty(m_vData) Then
        Me.lstData.Clear
        Exit Sub
    End If
    
    ReDim vFiltered(1 To UBound(m_vData, 1), 1 To M_LCOLCOUNT)
    lMatchCount = 0
    
    ' 1. Durchlaufe alle Zeilen im *Original*-Array (m_vData)
    For i = 1 To UBound(m_vData, 1)
        bRowAdded = False
        
        ' Wir durchsuchen die Spalten 3 bis 6 (C, D, E, F)
        For j = 3 To 6
            
            ' Wenn der Suchbegriff leer ist, zeige alles an.
            If sFilterTerm = "" Then
                lMatchCount = lMatchCount + 1
                ' Kopiere die ganze Zeile (alle 10 Spalten)
                Dim k As Long
                For k = 1 To M_LCOLCOUNT
                    vFiltered(lMatchCount, k) = m_vData(i, k)
                Next k
                Exit For ' Gehe zur nächsten Zeile (i)
                
            ' Wenn Suchbegriff nicht leer ist, suche nach Übereinstimmung
            ElseIf InStr(1, LCase(CStr(m_vData(i, j))), sFilterTerm) > 0 Then
                ' Treffer gefunden!
                lMatchCount = lMatchCount + 1
                
                ' Kopiere die *gesamte* Zeile (alle 10 Spalten)
                Dim k_ As Long
                For k_ = 1 To M_LCOLCOUNT
                    vFiltered(lMatchCount, k_) = m_vData(i, k_)
                Next k_
                
                ' Verlasse die innere Spalten-Schleife (j)
                Exit For
            End If
            
        Next j
    Next i
    
    ' 2. Leere die ListBox, bevor wir sie neu füllen
    Me.lstData.Clear
    
    ' 3. Fülle die ListBox mit den gefilterten Ergebnissen
    If lMatchCount > 0 Then
        ' Erstelle das finale Array in der exakt richtigen Größe
        Dim vFinal() As Variant
        ReDim vFinal(1 To lMatchCount, 1 To M_LCOLCOUNT)
        
        Dim x As Long, y As Long
        
        ' Wir kopieren die gefundenen Zeilen aus vFiltered in vFinal
        For x = 1 To lMatchCount ' Schleife für Zeilen
            For y = 1 To M_LCOLCOUNT ' Schleife für Spalten
                
                ' *** NEU: Datenformatierung beim Kopieren ***
                Select Case y
                    Case 7 ' Spalte G (Euro)
                        ' Prüfen, ob es eine Zahl ist
                        If IsNumeric(vFiltered(x, y)) Then
                            ' Format: "€"-Symbol, Tausenderpunkt, keine Dezimalen
                            vFinal(x, y) = Format(vFiltered(x, y), "€ #,##0")
                        Else
                            ' Falls Text (z.B. "N/A"), einfach übernehmen
                            vFinal(x, y) = vFiltered(x, y)
                        End If
                        
                    Case 9, 10 ' Spalte I und J (Prozent)
                        If IsNumeric(vFiltered(x, y)) Then
                            ' Format: "0.#%"
                            ' Zeigt "50%" (bei 0.5) oder "50,1%" (bei 0.501)
                            ' Passt sich automatisch an 0 oder 1 Dezimale an.
                            vFinal(x, y) = Format(vFiltered(x, y), "0.0%")
                        Else
                            vFinal(x, y) = vFiltered(x, y)
                        End If
                        
                    Case Else ' Alle anderen Spalten
                        ' Normal kopieren
                        vFinal(x, y) = vFiltered(x, y)
                End Select
                ' *** ENDE DER FORMATIERUNG ***
                
            Next y
        Next x
        
        ' Weise das *neue*, formatierte Array der ListBox zu
        Me.lstData.List = vFinal
        
    Else
        ' Keine Treffer gefunden
        Me.lstData.AddItem "Keine passenden Kunden gefunden."
    End If
    
End Sub

