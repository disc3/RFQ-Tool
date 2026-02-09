VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ResultsForm 
   Caption         =   "Add Component"
   ClientHeight    =   7005
   ClientLeft      =   105
   ClientTop       =   450
   ClientWidth     =   17670
   OleObjectBlob   =   "ResultsForm.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "ResultsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim originalRowIndices() As Long

Private sortColumnIndex As Long
Private sortAscending As Boolean

' --- Preferential Highlighting ---
Private Const ALT_FORECOLOR As Long = 32768 ' Dark green RGB(0,128,0)
Private m_alternativeFlags As Collection

Const THIS_PLANT = "This plant"
Const THIS_COMPANY = "This company (all plants)"
Const COMPANY_TP_LIST = "This company and Transfer Price List"
Const ALL_PLANTS = "All companies"

'##################################################################################
'# USERFORM INITIALIZATION
'##################################################################################
Private Sub UserForm_Initialize()
    Dim wsSource As Worksheet
    Dim headerRange As Range
    Dim headerCell As Range

    With Me.cmbSearchedPlants
        .AddItem THIS_PLANT
        .AddItem THIS_COMPANY
        .AddItem COMPANY_TP_LIST
        .AddItem ALL_PLANTS
        .ListIndex = 1
    End With
    
    ' Set ListView properties
    With Me.lstResults
        .View = lvwReport
        .Gridlines = True
        .FullRowSelect = True
        .HideColumnHeaders = False
    End With
    
    ' Populate ComboBox with options
    With Me.cmbAlternateOptions
        .AddItem "No" ' Default option
        .AddItem "LAPP component"
        .AddItem "Alternate supplier with stock"
        .ListIndex = 0 ' Set default selection to "No"
    End With

    ' Set default values for controls
    Me.txtSearch.Value = ""
    Me.lblStatus.Caption = "" ' Initialize status label

    ' --- Dynamically load column headers for the ListView ---
    Me.lstResults.columnHeaders.Clear
    
    On Error Resume Next
    Set wsSource = ThisWorkbook.Sheets("Purchasing Info Records")
    If wsSource Is Nothing Then
        MsgBox "Source Sheet 'Purchasing Info Records' could not be found. Header columns could not be loaded.", vbCritical
        Exit Sub
    End If
    
    ' Get column headers from the "LoadedData" table, but exclude "SearchColumn"
    Dim lo As ListObject
    Dim col As ListColumn
    Set lo = wsSource.ListObjects("LoadedData")
    If lo Is Nothing Then
        MsgBox "Table 'LoadedData' could not be found. Header columns could not be loaded.", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    Dim colWidth As Integer
    For Each col In lo.ListColumns
        ' The search column is not displayed in the ListView
        If col.name = "Material description" Or col.name = "Vendor name" Then
            colWidth = 200
        ElseIf col.name = "Plant" Or col.name = "Base unit of component" Or col.name = "Condition Currency" Then
            colWidth = 50
        Else
            colWidth = 100
        End If
        If col.name <> "SearchColumn" Then
             Me.lstResults.columnHeaders.Add Text:=col.name, Width:=colWidth
        End If
    Next col
End Sub

'##################################################################################
'# TRIGGER SEARCH
'##################################################################################
Private Sub btnSearch_Click()
    Dim searchTerm As String
    Dim plantsToInclude As New Collection
    Dim userPlant, searchColumn As String
    Dim foundItems As Collection
    Dim itemData As Variant
    Dim li As listItem
    Dim i As Long
    
    ' --- 1. Preparation ---
    Me.lblStatus.Caption = "Loading Database, please wait... (auto-save will be turned off until you close the file or you unload the database)"
    Me.Repaint ' Ensures the message is displayed immediately
    application.Cursor = xlWait ' Set mouse cursor to "Wait"
    Me.lstResults.ListItems.Clear

    ' --- 2. Update database (optional, but recommended) ---
    ' Ensures that the data from the Power Query is up to date
    LoadDatabase ' This calls the Sub from the module

    ' --- 3. Collect filter settings ---
    searchTerm = Me.txtSearch.Text
    
    Me.lblStatus.Caption = "Search is running, please wait..."
    Me.Repaint
    
    ' Assemble plant filter based on the ComboBox
    If StrComp(Me.cmbSearchedPlants.Text, ALL_PLANTS) <> 0 Then
    
        On Error Resume Next
        If StrComp(Me.cmbSearchedPlants.Text, THIS_PLANT) = 0 Then
            searchColumn = "Plant"
            userPlant = Trim(CStr(ThisWorkbook.Sheets("Global Variables").Range("B3").Value))
        Else
            searchColumn = "Source"
            userPlant = Trim(CStr(ThisWorkbook.Sheets("Global Variables").Range("B2").Value))
        End If
        On Error GoTo 0
        
        If userPlant <> "" Then
            plantsToInclude.Add userPlant
            If StrComp(Me.cmbSearchedPlants.Text, COMPANY_TP_LIST) = 0 Then
                plantsToInclude.Add "Transfer Price List"
            End If
        Else
            MsgBox "No plant defined in 'Global Variables'!B2. The search will be performed for all plants.", vbInformation
        End If
    Else
        searchColumn = "Source"
    End If
    
    ' --- 4. Call filter function (with alternative flags) ---
    Set m_alternativeFlags = New Collection
    Set foundItems = GetFilteredData(searchTerm, plantsToInclude, searchColumn, m_alternativeFlags)

    ' --- 5. Display results in ListView ---
    Dim altCount As Long
    altCount = 0

    If Not foundItems Is Nothing Then
        If foundItems.Count > 0 Then
            Dim itemIndex As Long
            itemIndex = 0
            For Each itemData In foundItems ' itemData is an array with the values of a row
                itemIndex = itemIndex + 1
                ' Add first value as the main item
                Set li = Me.lstResults.ListItems.Add(, , CStr(itemData(1)))
                ' Add the remaining values as subitems
                For i = 2 To UBound(itemData)
                    li.SubItems(i - 1) = CStr(itemData(i))
                Next i

                ' Apply preferential highlighting for alternatives
                Dim isAlt As Boolean
                isAlt = False
                If Not m_alternativeFlags Is Nothing Then
                    If itemIndex <= m_alternativeFlags.Count Then
                        isAlt = CBool(m_alternativeFlags(itemIndex))
                    End If
                End If

                If isAlt Then
                    li.Tag = "ALT"
                    li.ForeColor = ALT_FORECOLOR
                    Dim si As Long
                    For si = 1 To li.ListSubItems.Count
                        li.ListSubItems(si).ForeColor = ALT_FORECOLOR
                    Next si
                    altCount = altCount + 1
                Else
                    li.Tag = ""
                End If
            Next itemData

            ' Show count with alternative breakdown
            If altCount > 0 Then
                Me.lblStatus.Caption = foundItems.Count & " result(s) found (" & altCount & " preferred alternative(s) shown in green)."
            Else
                Me.lblStatus.Caption = foundItems.Count & " result(s) found."
            End If
        Else
            Me.lblStatus.Caption = "No results were found for this search criteria."
        End If
    Else
        Me.lblStatus.Caption = "An error occurred or no data is available."
    End If

    ' --- 6. Cleanup ---
    Set foundItems = Nothing
    Set plantsToInclude = Nothing
    application.Cursor = xlDefault
End Sub

Private Sub btnCancel_Click()
    ' Close the form
    Unload Me
End Sub

Private Sub btnCopy_Click()
    Dim selectedItem As MSComctlLib.listItem
    Dim partNumber As String
    Dim plant As String
    Dim quantity As Double
    Dim alternateInfo As String
    Dim partNumberColumnIndex, plantColumnIndex As Long

    ' Step 1: Check if a row is selected
    Set selectedItem = Me.lstResults.selectedItem
    If selectedItem Is Nothing Then
        MsgBox "Please first select a row to copy.", vbExclamation
        Exit Sub
    End If

    ' Step 2: Find the column index for the part number.
    partNumberColumnIndex = GetColumnIndexByName(Me.lstResults, "Material")
    If partNumberColumnIndex = 0 Then
        MsgBox "The column 'Material' could not be found.", vbCritical
        Exit Sub
    End If

    ' Step 3: Read the part number from the selected row
    If partNumberColumnIndex = 1 Then
        partNumber = selectedItem.Text
    Else
        partNumber = selectedItem.SubItems(partNumberColumnIndex - 1)
    End If

    ' Step 4: Find the column index for the plant.
    plantColumnIndex = GetColumnIndexByName(Me.lstResults, "Plant")
    If plantColumnIndex = 0 Then
        MsgBox "The column 'Plant' could not be found.", vbCritical
        Exit Sub
    End If

    ' Step 5: Read the plant from the selected row
    If plantColumnIndex = 1 Then
        plant = selectedItem.Text
    Else
        plant = selectedItem.SubItems(plantColumnIndex - 1)
    End If

    ' Step 6: Get the alternate info from the ComboBox
    alternateInfo = Me.cmbAlternateOptions.Value

    ' Step 7: Ask the user for the quantity
    Dim piecesStr As String
    piecesStr = InputBox("Please enter the required quantity:", "Number of Pieces")
    If StrPtr(piecesStr) = 0 Then Exit Sub ' User clicked Cancel
    If IsNumeric(piecesStr) And piecesStr <> "" Then
        quantity = CDbl(piecesStr)
    Else
        MsgBox "Invalid input for the number of pieces: '" & piecesStr & "'. Please enter a valid number.", vbExclamation
        Exit Sub
    End If

    ' Step 8: Call the target function with all four parameters
    AddFullComponent partNumber, plant, quantity, alternateInfo
    SortSelectedComponentsByProduct
    Unload Me
End Sub

'----------------------------------------------------------------------------------
' HELPER: Finds the index of a column in a ListView by its name
'----------------------------------------------------------------------------------
Private Function GetColumnIndexByName(lv As MSComctlLib.ListView, ByVal columnName As String) As Long
    Dim i As Long
    For i = 1 To lv.columnHeaders.Count
        If lv.columnHeaders(i).Text = columnName Then
            GetColumnIndexByName = i
            Exit Function
        End If
    Next i
    ' Returns 0 if the column was not found
    GetColumnIndexByName = 0
End Function


' Button: Add NEW Component
Private Sub btnAddNewComponent_Click()
    ' Open the Add Component form
    AddComponentForm.Show
End Sub


Private Sub lstResults_DblClick()
    If Not Me.lstResults.selectedItem Is Nothing Then
        btnCopy_Click
    End If
End Sub

Private Sub lstResults_BeforeLabelEdit(Cancel As Integer)
    Cancel = True ' Disables editing of item labels
End Sub

Private Sub lstResults_ColumnClick(ByVal columnHeader As MSComctlLib.columnHeader)
    ' Toggle ascending/descending if same column is clicked
    If sortColumnIndex = columnHeader.Index Then
        sortAscending = Not sortAscending
    Else
        sortColumnIndex = columnHeader.Index
        sortAscending = True ' default to ascending on new column
    End If

    Call SortListViewByColumn(lstResults, columnHeader.Index, sortAscending)
End Sub

Private Sub SortListViewByColumn(lv As MSComctlLib.ListView, ByVal colIndex As Long, ByVal ascending As Boolean)

    ' --- 1. Exit if there is nothing to sort ---
    If lv.ListItems.Count < 2 Then Exit Sub

    ' Turn off screen updates for the ListView for performance
    application.ScreenUpdating = False

    ' --- 2. Load all ListView data and tags into arrays ---
    ' This is much faster than reading from the ListView repeatedly.
    Dim listData() As String
    Dim listTags() As String
    ReDim listData(1 To lv.ListItems.Count, 1 To lv.columnHeaders.Count)
    ReDim listTags(1 To lv.ListItems.Count)
    Dim i, j, k As Long

    For i = 1 To lv.ListItems.Count
        listData(i, 1) = lv.ListItems(i).Text
        listTags(i) = lv.ListItems(i).Tag
        For j = 2 To lv.columnHeaders.Count
            listData(i, j) = lv.ListItems(i).SubItems(j - 1)
        Next j
    Next i

    ' --- 3. Sort the array using a simple Bubble Sort ---
    ' This version now correctly handles both text and number sorting.
    Dim temp As String
    For i = 1 To UBound(listData, 1) - 1
        For j = i + 1 To UBound(listData, 1)

            Dim val1 As String: val1 = listData(i, colIndex)
            Dim val2 As String: val2 = listData(j, colIndex)

            Dim shouldSwap As Boolean

            ' IMPROVEMENT: Check if values are numeric for proper sorting
            If IsNumeric(val1) And IsNumeric(val2) Then
                ' Compare as numbers
                shouldSwap = IIf(ascending, CDbl(val1) > CDbl(val2), CDbl(val2) > CDbl(val1))
            Else
                ' Compare as case-insensitive text
                shouldSwap = IIf(ascending, LCase(val1) > LCase(val2), LCase(val2) > LCase(val1))
            End If

            If shouldSwap Then
                ' Swap the entire rows in our array (data + tag)
                For k = 1 To UBound(listData, 2)
                    temp = listData(i, k)
                    listData(i, k) = listData(j, k)
                    listData(j, k) = temp
                Next k
                temp = listTags(i)
                listTags(i) = listTags(j)
                listTags(j) = temp
            End If

        Next j
    Next i

    ' --- 4. Clear the ListView and reload the sorted data from the array ---
    lv.ListItems.Clear
    Dim li As listItem
    For i = 1 To UBound(listData, 1)
        Set li = lv.ListItems.Add(, , listData(i, 1))
        For j = 2 To UBound(listData, 2)
            li.SubItems(j - 1) = listData(i, j)
        Next j

        ' Restore tag and re-apply preferential highlighting
        li.Tag = listTags(i)
        If li.Tag = "ALT" Then
            li.ForeColor = ALT_FORECOLOR
            Dim si As Long
            For si = 1 To li.ListSubItems.Count
                li.ListSubItems(si).ForeColor = ALT_FORECOLOR
            Next si
        End If
    Next i

    ' Turn updates back on
    application.ScreenUpdating = True

End Sub
